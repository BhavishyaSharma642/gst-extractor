import os, re, zipfile, tempfile, uuid, io, subprocess, shutil, json, time
import urllib.request, urllib.error
import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template_string
from docx import Document

# ── PDF support ─────────────────────────────────────────────
try:
    import pdfplumber
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

# ── Ollama (local AI — free, no limits, no API key) ──────────
# Model to use. llama3.2 is fast and accurate for extraction.
# You can change to "mistral" or "phi3" if preferred.
OLLAMA_MODEL   = "llama3.2"
OLLAMA_URL     = "http://localhost:11434/api/generate"
OLLAMA_ENABLED = False

def _check_ollama():
    """Check if Ollama is running and the model is available."""
    try:
        req  = urllib.request.urlopen("http://localhost:11434/api/tags", timeout=3)
        data = json.loads(req.read())
        models = [m["name"].split(":")[0] for m in data.get("models", [])]
        if OLLAMA_MODEL in models:
            print(f"✅ Ollama AI fallback: ENABLED  (model: {OLLAMA_MODEL})")
            return True
        else:
            print(f"⚠  Ollama running but model '{OLLAMA_MODEL}' not found.")
            print(f"   Run:  ollama pull {OLLAMA_MODEL}")
            return False
    except Exception:
        print("⚠  Ollama not running. Start it with:  ollama serve")
        print("   AI fallback disabled — regex-only mode active.")
        return False

OLLAMA_ENABLED = _check_ollama()

# ── LibreOffice ──────────────────────────────────────────────
LIBREOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
if not os.path.exists(LIBREOFFICE):
    LIBREOFFICE = shutil.which("libreoffice") or shutil.which("soffice")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024

# ============================================================
# READ HELPERS
# ============================================================

def read_docx(path):
    doc   = Document(path)
    lines = []
    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            lines.append(t)
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                lines.append(" | ".join(cells))
    return "\n".join(lines)


def convert_doc_to_docx(doc_path, out_dir):
    if not LIBREOFFICE:
        return None
    try:
        subprocess.run(
            [LIBREOFFICE, "--headless", "--convert-to", "docx",
             "--outdir", out_dir, doc_path],
            capture_output=True, timeout=30
        )
        base      = os.path.splitext(os.path.basename(doc_path))[0]
        converted = os.path.join(out_dir, base + ".docx")
        return converted if os.path.exists(converted) else None
    except Exception:
        return None


def read_pdf(path):
    if not PDF_SUPPORT:
        return ""
    lines = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.append(text)
    return "\n".join(lines)


def read_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return read_docx(path)
    elif ext == ".doc":
        tmp       = tempfile.mkdtemp()
        converted = convert_doc_to_docx(path, tmp)
        return read_docx(converted) if converted else ""
    elif ext == ".pdf":
        return read_pdf(path)
    return ""

# ============================================================
# CLEAN HELPERS
# ============================================================

def clean_text(text):
    text = text.replace("\xa0", " ")
    text = re.sub(r'[ \t]+', ' ', text)
    return text


def clean_amount(value):
    """
    Clean a rupee amount string.
    Handles: 1,770.00 / 2,360-00 / 4,720.00.00 / 1770/-
    """
    if not value:
        return ""
    value = str(value)
    value = re.sub(r'[,\s]', '', value)
    value = value.replace("/-", "")
    value = re.sub(r'-(?=\d{2}$)', '.', value)          # 2360-00 → 2360.00
    value = re.sub(r'(\.\d{2})\.\d{2}$', r'\1', value)  # 4720.00.00 → 4720.00
    value = value.strip()
    try:
        return "{:.2f}".format(float(value))
    except ValueError:
        return value


def is_blank(v):
    if v is None:
        return True
    return str(v).strip() in ("", "0", "0.0", "0.00", "None", "nan")

# ============================================================
# BILL NUMBER
# 316_UIIC   → 316     (letter after _ separator is NOT part of bill no)
# 319A_NIA   → 319A    (A directly attached, no separator)
# 773B_Chola → 773B
# ============================================================

def extract_bill_no(filename):
    name = os.path.splitext(filename)[0]
    m    = re.match(r'^(\d+)([A-Za-z]?)(?:[_\-\s]|$)', name)
    if m:
        return m.group(1) + m.group(2).upper()
    return ""

# ============================================================
# DETECT COMPANY
# ============================================================

COMPANY_MAP = {
    "GODIGIT": "Go Digit General Insurance Ltd",
    "DIGIT":   "Go Digit General Insurance Ltd",
    "UIIC":    "United India Insurance Company Ltd",
    "NICL":    "National Insurance Company Ltd",
    "NIC":     "National Insurance Company Ltd",
    "NIA":     "New India Assurance Company Ltd",
    "SBI":     "SBI General Insurance Company Ltd",
    "IFFCO":   "Iffco-Tokio General Insurance Company Ltd",
    "ITGI":    "Iffco-Tokio General Insurance Company Ltd",
    "OIC":     "The Oriental Insurance Company Ltd",
    "SGIC":    "Shri Ram General Insurance Company Limited",
    "SGI":     "Shri Ram General Insurance Company Limited",
    "CHOLA":   "Chola MS General Insurance Company Ltd",
    "ACKO":    "Acko General Insurance Limited",
    "ZUNO":    "Zuno General Insurance Limited",
    "RAHEJA":  "Raheja QBE General Insurance Company Limited",
    "USGI":    "Universal Sompo General Insurance Company Ltd",
    "MGHDI":   "Magma HDI General Insurance Company",
    "LIBERTY": "Liberty General Insurance Limited",
    "HDFC":    "HDFC ERGO General Insurance Company Ltd",
    "BAJAJ":   "Bajaj Allianz General Insurance Company Ltd",
    "TATA":    "Tata AIG General Insurance Company Ltd",
    "ICICI":   "ICICI Lombard General Insurance Company Ltd",
    "RELIANCE":"Reliance General Insurance Company Ltd",
}

def detect_company(filename):
    up = filename.upper()
    for key, val in COMPANY_MAP.items():
        if key in up:
            return val
    return ""

# ============================================================
# REGEX EXTRACTION
# ============================================================

def extract_data(filename, text):

    bill_no = extract_bill_no(filename)
    company = detect_company(filename)

    # ── DATE ─────────────────────────────────────────────────
    date = ""
    for pat in [
        r'Invoice\s*Date\s*[:\-]?\s*([0-9]{1,2}[.\/\-][0-9]{1,2}[.\/\-][0-9]{2,4})',
        r'Bill\s*Date\s*[:\-]?\s*([0-9]{1,2}[.\/\-][0-9]{1,2}[.\/\-][0-9]{2,4})',
        r'Bill Date.*?\|\s*([0-9]{1,2}[.\/\-][0-9]{1,2}[.\/\-][0-9]{2,4})',
        r'Date\s*[:\-]?\s*([0-9]{1,2}[.\/\-][0-9]{1,2}[.\/\-][0-9]{2,4})',
        r'Dt\.?\s*[:\-]?\s*([0-9]{1,2}[.\/\-][0-9]{1,2}[.\/\-][0-9]{2,4})',
    ]:
        ms = re.findall(pat, text, re.IGNORECASE)
        if ms:
            date = ms[0]
            break

    # ── CGST ─────────────────────────────────────────────────
    cgst = ""
    for pat in [
        r'Service\s*CGST\s*Amount\s*\|[^\d\n]*([0-9,]+\.[0-9]{2})',
        r'Service\s*CGST\s*Amount[^\d\n]{0,20}([0-9,]+\.[0-9]{2})',
        r'CGST@0?9%\s*\|[^\d\n]*([0-9,]+\.[0-9]{2})',
        r'CGST@0?9%[^\d\n]{0,15}([0-9,]+\.[0-9]{2})',
        # FORMAT A: "C GST @ 9% | C GST @ 9% | C GST @ 9% | 180.00"
        # Use larger window (60 chars) to cross the repeated label cells
        r'C\s*GST\s*@\s*0?9\s*%[^\d\n]{0,80}([0-9,]+\.[0-9]{2})',
        r'CGST\s*@\s*0?9\s*%[^\d\n]{0,30}([0-9,]+\.[0-9]{2})',
        r'CGST[^\d\n]{0,25}([0-9,]+\.[0-9]{2})',
    ]:
        ms = re.findall(pat, text, re.IGNORECASE)
        if ms:
            cgst = ms[-1]
            break

    # ── SGST ─────────────────────────────────────────────────
    sgst = ""
    for pat in [
        r'Service\s*SGST\s*Amount\s*\|[^\d\n]*([0-9,]+\.[0-9]{2})',
        r'Service\s*SGST\s*Amount[^\d\n]{0,20}([0-9,]+\.[0-9]{2})',
        r'SCGST@0?9%\s*\|[^\d\n]*([0-9,]+\.[0-9]{2})',
        r'SCGST@0?9%[^\d\n]{0,15}([0-9,]+\.[0-9]{2})',
        r'S\s*GST\s*@\s*0?9\s*%[^\d\n]{0,80}([0-9,]+\.[0-9]{2})',
        r'SGST\s*@\s*0?9\s*%[^\d\n]{0,30}([0-9,]+\.[0-9]{2})',
        r'SGST[^\d\n]{0,25}([0-9,]+\.[0-9]{2})',
    ]:
        ms = re.findall(pat, text, re.IGNORECASE)
        if ms:
            sgst = ms[-1]
            break

    # ── IGST ─────────────────────────────────────────────────
    igst = ""
    for pat in [
        r'IGST\s*Amount\s*\|[^\d\n]*([0-9,]+\.[0-9]{2})',
        r'IGST\s*Amount[^\d\n]{0,20}([0-9,]+\.[0-9]{2})',
        r'ISGST\s*@\s*18\s*%[^\d\n]{0,80}([0-9,]+\.[0-9]{2})',
        r'ICGST@\s*18\s*%[^\d\n]{0,80}([0-9,]+\.[0-9]{2})',  # Iffco variant
        r'ICGST@18%[^\d\n]{0,80}([0-9,]+\.[0-9]{2})',
        r'GST\s*Tax\s*\|[^\d\n]*@18[^\d\n]*([0-9,]+\.[0-9]{2})',
        r'GST\s*Tax[^\d\n]{0,30}([0-9,]+\.[0-9]{2})',
        r'IGST\s*@\s*18\s*%[^\d\n]{0,30}([0-9,]+\.[0-9]{2})',
        r'IGST[^\d\n]{0,25}([0-9,]+\.[0-9]{2})',
    ]:
        ms = re.findall(pat, text, re.IGNORECASE)
        if ms:
            igst = ms[-1]
            break

    # ── GRAND TOTAL ───────────────────────────────────────────
    grand_total = ""
    for pat in [
        r'Grand\s*Total[^\n]*?([0-9,]+\.[0-9]{2}(?:\.[0-9]{2})?)\s*(?:\|.*)?$',
        r'Grand\s*Total[^\n]*?([0-9,]+\-[0-9]{2})\s*(?:\|.*)?$',  # 2,360-00 variant
        r'Total\s*Amount\s*\(?In\s*Fig[^\n]*?Rs\.?\s*([0-9,]+(?:\.[0-9]{2})?)',
        r'Total\s*Amount\s*\(?In\s*Fig[^\d\n]{0,30}([0-9,]+\.[0-9]{2})',
        r'Total\s*Amount\s*(?:Payable)?[^\d\n]{0,30}([0-9,]+\.[0-9]{2})',
        r'Or\s*Say\s*Rs\.?\s*([0-9,]+(?:\.[0-9]{2})?)',
        r'Invoice\s*(?:Value|Amount)[^\d\n]{0,30}([0-9,]+\.[0-9]{2})',
        r'Net\s*(?:Payable|Amount)[^\d\n]{0,30}([0-9,]+\.[0-9]{2})',
        r'Amount\s*Payable[^\d\n]{0,30}([0-9,]+\.[0-9]{2})',
    ]:
        ms = re.findall(pat, text, re.IGNORECASE | re.MULTILINE)
        if ms:
            grand_total = ms[-1]
            break

    return {
        "Bill No":     bill_no,
        "Company":     company,
        "Date":        date,
        "Grand Total": clean_amount(grand_total),
        "IGST":        clean_amount(igst),
        "CGST":        clean_amount(cgst),
        "SGST":        clean_amount(sgst),
        "_ai_filled":  False,
    }

# ============================================================
# OLLAMA FALLBACK — local AI, free, no limits, no API key
#
# GST LAW NOTE: A bill uses EITHER:
#   IGST  (inter-state)   OR   CGST + SGST (intra-state)
# Never both. So blank CGST/SGST when IGST exists = correct.
# ============================================================

def _ollama_ask(prompt):
    """Send a prompt to local Ollama and return the response text."""
    payload = json.dumps({
        "model":  OLLAMA_MODEL,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0,      # deterministic — critical for data extraction
            "num_predict": 200,    # short answer = faster response
        }
    }).encode("utf-8")

    req = urllib.request.Request(
        OLLAMA_URL,
        data    = payload,
        headers = {"Content-Type": "application/json"},
        method  = "POST"
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = json.loads(resp.read())
        return data.get("response", "").strip()


def ollama_fill_missing(filename, text, row):
    """Use local Ollama ONLY to fill fields that regex couldn't find."""
    if not OLLAMA_ENABLED:
        return row

    igst_filled = not is_blank(row.get("IGST"))
    cgst_filled = not is_blank(row.get("CGST"))
    sgst_filled = not is_blank(row.get("SGST"))

    missing = []

    if is_blank(row.get("Date")):
        missing.append("invoice_date")

    if is_blank(row.get("Grand Total")):
        missing.append("grand_total")

    # Only request CGST/SGST if IGST is also blank (inter-state bill)
    if not igst_filled:
        if not cgst_filled: missing.append("cgst")
        if not sgst_filled: missing.append("sgst")

    # Only request IGST if both CGST and SGST are blank (intra-state bill)
    if not cgst_filled and not sgst_filled:
        if not igst_filled: missing.append("igst")

    # Sanity-check grand total value
    try:
        total   = float(row.get("Grand Total", "0") or "0")
        gst_sum = sum(float(row.get(k, "0") or "0") for k in ("CGST", "SGST", "IGST"))
        if total < 100 or (gst_sum > 0 and gst_sum > total):
            if "grand_total" not in missing:
                missing.append("grand_total")
    except Exception:
        if "grand_total" not in missing:
            missing.append("grand_total")

    if not missing:
        return row   # ✅ Regex got everything — skip AI entirely

    print(f"  🤖 Ollama fixing [{filename}] → {missing}")

    prompt = f"""You are a GST invoice data extraction assistant.
Extract ONLY the following fields from the invoice text: {', '.join(missing)}

STRICT RULES:
- grand_total: the FINAL total amount payable including ALL taxes. Never the subtotal.
- invoice_date: the bill/invoice date exactly as written e.g. 12.09.2025
- cgst: CGST tax amount as a plain number e.g. 180.00
- sgst: SGST tax amount as a plain number e.g. 180.00
- igst: IGST tax amount as a plain number e.g. 360.00
- If a field is not present in the invoice, use empty string "".
- Return ONLY a valid JSON object with these exact keys. No explanation, no markdown.

Invoice filename: {filename}

Invoice text:
{text[:4000]}

JSON output:"""

    try:
        raw  = _ollama_ask(prompt)
        # Strip markdown fences if model adds them
        raw  = re.sub(r'```(?:json)?', '', raw).strip().rstrip('`').strip()
        # Extract first JSON object from response
        m    = re.search(r'\{.*\}', raw, re.DOTALL)
        if not m:
            print(f"    ✗ No JSON found in response: {raw[:80]}")
            return row
        data = json.loads(m.group())

        field_map = {
            "invoice_date": ("Date",        False),   # False = plain text
            "grand_total":  ("Grand Total", True),    # True  = clean_amount
            "cgst":         ("CGST",        True),
            "sgst":         ("SGST",        True),
            "igst":         ("IGST",        True),
        }

        ai_helped = False
        for ai_key, (excel_key, is_amount) in field_map.items():
            if ai_key not in missing:
                continue
            val = str(data.get(ai_key, "")).strip()
            if val:
                row[excel_key] = clean_amount(val) if is_amount else val
                ai_helped      = True
                print(f"    ✓ {excel_key} = {row[excel_key]}")

        if ai_helped:
            row["_ai_filled"] = True

    except json.JSONDecodeError as e:
        print(f"    ✗ JSON parse error [{filename}]: {e} | raw: {raw[:80]}")
    except Exception as e:
        print(f"    ✗ Ollama error [{filename}]: {str(e)[:120]}")

    return row

# ============================================================
# PROCESS ZIP
# ============================================================

def process_zip(zip_path):
    results  = []
    errors   = []
    seen_nos = {}

    with tempfile.TemporaryDirectory() as tmp:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(tmp)

        all_files = []
        for root, _, files in os.walk(tmp):
            for fname in files:
                if fname.startswith("~$"):
                    continue
                ext = os.path.splitext(fname)[1].lower()
                if ext in (".docx", ".doc", ".pdf"):
                    all_files.append((root, fname, ext))

        # .docx before .doc for same base name
        all_files.sort(key=lambda x: (x[1], 0 if x[2] == ".docx" else 1))

        seen_files = set()
        for root, fname, ext in all_files:
            base = os.path.splitext(fname)[0]
            if base in seen_files:
                continue
            seen_files.add(base)

            fpath = os.path.join(root, fname)
            try:
                raw  = read_file(fpath)
                text = clean_text(raw)
                if not text.strip():
                    errors.append({"File": fname, "Error": "Empty / unreadable — install LibreOffice for .doc"})
                    continue

                row = extract_data(fname, text)
                row = ollama_fill_missing(fname, text, row)

                bill_no = row["Bill No"]
                if bill_no and bill_no in seen_nos:
                    existing = results[seen_nos[bill_no]]
                    if sum(1 for v in row.values() if v) > sum(1 for v in existing.values() if v):
                        results[seen_nos[bill_no]] = row
                else:
                    seen_nos[bill_no] = len(results)
                    results.append(row)

            except Exception as e:
                errors.append({"File": fname, "Error": str(e)})

    def sort_key(r):
        m = re.match(r'(\d+)', r.get("Bill No",""))
        return (int(m.group(1)) if m else 9999999, r.get("Bill No",""))

    results.sort(key=sort_key)
    return results, errors

# ============================================================
# HTML TEMPLATE
# ============================================================

HTML = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>GST Invoice Extractor</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@300;400;500&display=swap" rel="stylesheet"/>
<style>
  :root{--bg:#0a0a0f;--panel:#13131a;--border:#2a2a3a;--accent:#00e5a0;--accent2:#7c6aff;--text:#e8e8f0;--muted:#6b6b80;--danger:#ff4f6a;}
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:var(--bg);color:var(--text);font-family:'DM Mono',monospace;min-height:100vh;overflow-x:hidden;}
  body::before{content:'';position:fixed;inset:0;background-image:url("data:image/svg+xml,%3Csvg viewBox='0 0 200 200' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.04'/%3E%3C/svg%3E");pointer-events:none;z-index:0;opacity:.4;}
  .wrap{max-width:1300px;margin:0 auto;padding:2rem;position:relative;z-index:1;}
  header{display:flex;align-items:center;gap:1rem;margin-bottom:3rem;}
  .logo{font-family:'Syne',sans-serif;font-weight:800;font-size:2rem;letter-spacing:-.03em;}
  .logo span{color:var(--accent);}
  .tag{font-size:.7rem;color:var(--muted);border:1px solid var(--border);padding:.2rem .6rem;border-radius:999px;}
  .ai-badge{font-size:.7rem;background:rgba(124,106,255,.15);color:var(--accent2);border:1px solid rgba(124,106,255,.3);padding:.2rem .6rem;border-radius:999px;margin-left:auto;}
  #drop-zone{border:2px dashed var(--border);border-radius:16px;padding:4rem 2rem;text-align:center;cursor:pointer;transition:border-color .25s,background .25s;background:var(--panel);position:relative;overflow:hidden;}
  #drop-zone:hover,#drop-zone.drag-over{border-color:var(--accent);background:rgba(0,229,160,.04);}
  #drop-zone .icon{font-size:3rem;margin-bottom:1rem;display:block;}
  #drop-zone h2{font-family:'Syne',sans-serif;font-size:1.4rem;font-weight:700;margin-bottom:.5rem;}
  #drop-zone p{color:var(--muted);font-size:.8rem;}
  #drop-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;}
  #file-chip{display:none;align-items:center;gap:.75rem;margin-top:1.5rem;background:rgba(0,229,160,.08);border:1px solid rgba(0,229,160,.25);border-radius:12px;padding:.75rem 1.25rem;font-size:.85rem;}
  #file-chip.show{display:flex;}
  .fname{color:var(--accent);font-weight:500;}.fsize{color:var(--muted);font-size:.75rem;}
  .btn{display:inline-flex;align-items:center;gap:.5rem;background:var(--accent);color:#000;font-family:'Syne',sans-serif;font-weight:700;font-size:.9rem;border:none;border-radius:10px;padding:.8rem 1.8rem;cursor:pointer;transition:transform .15s,box-shadow .15s;}
  .btn:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(0,229,160,.3);}
  .btn:disabled{opacity:.4;cursor:not-allowed;transform:none;box-shadow:none;}
  .btn.secondary{background:transparent;color:var(--accent);border:1px solid var(--accent);}
  .btn.secondary:hover{background:rgba(0,229,160,.08);}
  .btn.dl{background:var(--accent2);}
  .btn.dl:hover{box-shadow:0 8px 24px rgba(124,106,255,.35);}
  #progress-wrap{display:none;margin-top:2rem;}
  #progress-wrap.show{display:block;}
  .prog-label{font-size:.8rem;color:var(--muted);margin-bottom:.5rem;display:flex;justify-content:space-between;}
  .prog-bar-bg{height:6px;background:var(--border);border-radius:999px;overflow:hidden;}
  .prog-bar{height:100%;width:0%;background:linear-gradient(90deg,var(--accent),var(--accent2));border-radius:999px;transition:width .4s ease;}
  #status-text{font-size:.78rem;color:var(--muted);margin-top:.5rem;}
  #results{display:none;margin-top:3rem;}
  #results.show{display:block;}
  .results-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:1.5rem;flex-wrap:wrap;gap:1rem;}
  .results-header h3{font-family:'Syne',sans-serif;font-size:1.2rem;font-weight:700;}
  .stats{display:flex;gap:1rem;flex-wrap:wrap;}
  .stat{background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:.6rem 1.1rem;text-align:center;}
  .stat .val{font-size:1.3rem;font-weight:700;font-family:'Syne',sans-serif;color:var(--accent);}
  .stat .lbl{color:var(--muted);font-size:.7rem;margin-top:.1rem;}
  .info-box{background:rgba(124,106,255,.07);border:1px solid rgba(124,106,255,.2);border-radius:10px;padding:.8rem 1.2rem;font-size:.75rem;color:var(--muted);margin-top:1rem;line-height:1.6;}
  .info-box strong{color:var(--accent2);}
  .table-wrap{overflow-x:auto;border-radius:12px;border:1px solid var(--border);}
  table{width:100%;border-collapse:collapse;font-size:.78rem;}
  thead tr{background:rgba(255,255,255,.03);}
  th{padding:.8rem 1rem;text-align:left;color:var(--muted);font-weight:500;border-bottom:1px solid var(--border);white-space:nowrap;font-size:.68rem;letter-spacing:.08em;text-transform:uppercase;}
  td{padding:.7rem 1rem;border-bottom:1px solid rgba(255,255,255,.04);vertical-align:middle;}
  tr:last-child td{border-bottom:none;}
  tr:hover td{background:rgba(255,255,255,.025);}
  .c-billno{font-family:'Syne',sans-serif;font-weight:800;font-size:.9rem;color:var(--accent);}
  .c-company{color:var(--text);max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
  .c-date{color:#aaa;white-space:nowrap;}
  .c-total{font-weight:600;color:#fff;}
  .c-gst{color:#7adbb5;}
  .rs{font-size:.65rem;color:var(--muted);margin-right:1px;}
  .empty{color:var(--muted);font-style:italic;font-size:.72rem;}
  .na{color:#3a3a50;font-size:.7rem;}
  .ai-row td:first-child{border-left:2px solid var(--accent2);}
  .ai-tag{font-size:.58rem;background:rgba(124,106,255,.25);color:var(--accent2);padding:.1rem .35rem;border-radius:4px;margin-left:.4rem;}
  #errors-box{display:none;margin-top:1.5rem;background:rgba(255,79,106,.05);border:1px solid rgba(255,79,106,.2);border-radius:12px;padding:1.25rem;}
  #errors-box.show{display:block;}
  #errors-box h4{color:var(--danger);font-family:'Syne',sans-serif;font-size:.9rem;margin-bottom:.75rem;}
  .err-item{font-size:.75rem;color:var(--muted);padding:.3rem 0;border-bottom:1px solid rgba(255,79,106,.08);}
  .err-item:last-child{border-bottom:none;}
  .err-item strong{color:var(--text);}
  .actions{display:flex;gap:1rem;margin-top:2rem;flex-wrap:wrap;}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="logo">GST<span>.</span>extract</div>
    <div class="tag">LOCAL · v6.0</div>
    <div class="ai-badge">🦙 Ollama AI fallback</div>
  </header>

  <div id="drop-zone">
    <input type="file" id="file-input" accept=".zip"/>
    <span class="icon">📦</span>
    <h2>Drop your ZIP file here</h2>
    <p>Pack all bills (.docx / .doc / .pdf) into one ZIP &nbsp;·&nbsp; Handles 500+ bills &nbsp;·&nbsp; AI fills missing fields automatically</p>
  </div>

  <div class="info-box">
    <strong>Powered by local Ollama AI</strong> — runs on your PC, no internet needed, zero limits, completely free forever.
    Regex extracts fields first (instant). Ollama fills only what's missing. &nbsp;·&nbsp;
    <strong>GST Law:</strong> Bills use IGST (inter-state) OR CGST+SGST (intra-state) — never both. Blank GST fields on the other type = correct.
  </div>

  <div id="file-chip">
    <span>📄</span><span class="fname" id="chip-name"></span><span class="fsize" id="chip-size"></span>
    <span style="margin-left:auto"><button class="btn secondary" style="padding:.4rem .9rem;font-size:.78rem" onclick="clearFile()">✕ Clear</button></span>
  </div>
  <div style="margin-top:1.5rem;">
    <button class="btn" id="process-btn" disabled onclick="processFile()">⚡ Extract GST Data</button>
  </div>

  <div id="progress-wrap">
    <div class="prog-label"><span>Processing bills…</span><span id="prog-pct">0%</span></div>
    <div class="prog-bar-bg"><div class="prog-bar" id="prog-bar"></div></div>
    <div id="status-text">Initialising…</div>
  </div>

  <div id="results">
    <div class="results-header">
      <h3>Extracted GST Data</h3>
      <div class="stats">
        <div class="stat"><div class="val" id="stat-total">0</div><div class="lbl">Total Bills</div></div>
        <div class="stat"><div class="val" id="stat-ok">0</div><div class="lbl">Extracted</div></div>
        <div class="stat"><div class="val" id="stat-ai" style="color:var(--accent2)">0</div><div class="lbl">AI Assisted</div></div>
        <div class="stat"><div class="val" id="stat-err" style="color:var(--danger)">0</div><div class="lbl">Errors</div></div>
      </div>
    </div>

    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Bill No</th><th>Company Name</th><th>Date</th>
            <th>Grand Total (₹)</th><th>IGST (₹)</th><th>CGST (₹)</th><th>SGST (₹)</th>
          </tr>
        </thead>
        <tbody id="table-body"></tbody>
      </table>
    </div>

    <div id="errors-box"><h4>⚠ Files with issues</h4><div id="errors-list"></div></div>
    <div class="actions">
      <button class="btn dl" onclick="downloadExcel()">⬇ Download Excel</button>
      <button class="btn secondary" onclick="resetAll()">↺ Process Another ZIP</button>
    </div>
  </div>
</div>

<script>
  let sessionId=null;
  const dropZone=document.getElementById('drop-zone'),fileInput=document.getElementById('file-input'),
    fileChip=document.getElementById('file-chip'),processBtn=document.getElementById('process-btn'),
    progWrap=document.getElementById('progress-wrap'),progBar=document.getElementById('prog-bar'),
    progPct=document.getElementById('prog-pct'),statusText=document.getElementById('status-text'),
    resultsDiv=document.getElementById('results');

  dropZone.addEventListener('dragover',e=>{e.preventDefault();dropZone.classList.add('drag-over');});
  dropZone.addEventListener('dragleave',()=>dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop',e=>{
    e.preventDefault();dropZone.classList.remove('drag-over');
    const f=e.dataTransfer.files[0];
    if(f&&f.name.endsWith('.zip'))setFile(f);else alert('Please drop a .zip file.');
  });
  fileInput.addEventListener('change',()=>{if(fileInput.files[0])setFile(fileInput.files[0]);});

  function setFile(f){
    document.getElementById('chip-name').textContent=f.name;
    document.getElementById('chip-size').textContent=(f.size/1024/1024).toFixed(1)+' MB';
    fileChip.classList.add('show');processBtn.disabled=false;processBtn._file=f;
  }
  function clearFile(){
    fileInput.value='';fileChip.classList.remove('show');processBtn.disabled=true;processBtn._file=null;
  }

  let progInterval=null;
  function startFakeProgress(){
    let pct=0;progWrap.classList.add('show');
    const msgs=['Reading ZIP…','Extracting files…','Running regex…','AI fixing missing fields…','Deduplicating…','Wrapping up…'];
    let mi=0;
    progInterval=setInterval(()=>{
      pct=Math.min(pct+Math.random()*1.5,90);
      progBar.style.width=pct+'%';progPct.textContent=Math.floor(pct)+'%';
      if(mi<msgs.length&&pct>mi*15)statusText.textContent=msgs[mi++];
    },400);
  }
  function finishProgress(){clearInterval(progInterval);progBar.style.width='100%';progPct.textContent='100%';statusText.textContent='Done!';}

  async function processFile(){
    const f=processBtn._file;if(!f)return;
    processBtn.disabled=true;resultsDiv.classList.remove('show');startFakeProgress();
    const form=new FormData();form.append('file',f);
    try{
      const res=await fetch('/upload',{method:'POST',body:form});
      const data=await res.json();finishProgress();
      if(data.error){alert('Server error: '+data.error);return;}
      sessionId=data.session_id;
      renderResults(data.results,data.errors,data.ai_count||0);
    }catch(err){finishProgress();alert('Connection failed: '+err.message);}
    finally{processBtn.disabled=false;}
  }

  function fmtAmt(val){return val?'<span class="rs">₹</span>'+val:'<span class="empty">—</span>';}
  function fmtTxt(val){return val||'<span class="empty">—</span>';}
  // Show N/A (greyed) for GST fields that are legitimately absent (other type used)
  function fmtGst(val,otherFilled){
    if(val)return'<span class="rs">₹</span>'+val;
    if(otherFilled)return'<span class="na">N/A</span>';
    return'<span class="empty">—</span>';
  }

  function renderResults(rows,errors,aiCount){
    document.getElementById('stat-total').textContent=rows.length+(errors||[]).length;
    document.getElementById('stat-ok').textContent=rows.length;
    document.getElementById('stat-ai').textContent=aiCount;
    document.getElementById('stat-err').textContent=(errors||[]).length;
    const tbody=document.getElementById('table-body');
    tbody.innerHTML='';
    rows.forEach(r=>{
      const aiTag=r._ai_filled?'<span class="ai-tag">AI</span>':'';
      const hasIgst=!!r.IGST, hasCgst=!!(r.CGST||r.SGST);
      const tr=document.createElement('tr');
      if(r._ai_filled)tr.classList.add('ai-row');
      tr.innerHTML=
        `<td class="c-billno">${fmtTxt(r['Bill No'])}</td>`+
        `<td class="c-company" title="${r.Company||''}">${fmtTxt(r.Company)}</td>`+
        `<td class="c-date">${fmtTxt(r.Date)}${aiTag}</td>`+
        `<td class="c-total">${fmtAmt(r['Grand Total'])}</td>`+
        `<td class="c-gst">${fmtGst(r.IGST,hasCgst)}</td>`+
        `<td class="c-gst">${fmtGst(r.CGST,hasIgst)}</td>`+
        `<td class="c-gst">${fmtGst(r.SGST,hasIgst)}</td>`;
      tbody.appendChild(tr);
    });
    const errBox=document.getElementById('errors-box'),errList=document.getElementById('errors-list');
    if(errors&&errors.length){
      errList.innerHTML=errors.map(e=>`<div class="err-item"><strong>${e.File}</strong> — ${e.Error}</div>`).join('');
      errBox.classList.add('show');
    }else errBox.classList.remove('show');
    resultsDiv.classList.add('show');
  }

  function downloadExcel(){if(sessionId)window.location.href='/download/'+sessionId;}
  function resetAll(){
    clearFile();progWrap.classList.remove('show');
    progBar.style.width='0%';progPct.textContent='0%';statusText.textContent='Initialising…';
    resultsDiv.classList.remove('show');
    document.getElementById('errors-box').classList.remove('show');
    sessionId=null;
  }
</script>
</body>
</html>
"""

# ============================================================
# ROUTES
# ============================================================

SESSIONS = {}

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if not f.filename.endswith(".zip"):
        return jsonify({"error": "Please upload a .zip file"}), 400
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    f.save(tmp.name); tmp.close()
    results, errors = process_zip(tmp.name)
    os.unlink(tmp.name)
    ai_count = sum(1 for r in results if r.get("_ai_filled"))
    sid = str(uuid.uuid4())
    SESSIONS[sid] = results
    return jsonify({"session_id": sid, "results": results,
                    "errors": errors, "count": len(results), "ai_count": ai_count})


@app.route("/download/<session_id>")
def download(session_id):
    rows = SESSIONS.get(session_id)
    if rows is None:
        return "Session not found", 404
    export = [{k:v for k,v in r.items() if not k.startswith("_")} for r in rows]
    cols   = ["Bill No","Company","Date","Grand Total","IGST","CGST","SGST"]
    df     = pd.DataFrame(export, columns=cols)
    buf    = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="GST Details")
        ws = writer.sheets["GST Details"]
        from openpyxl.styles import Font, PatternFill, Alignment
        hdr_fill = PatternFill("solid", fgColor="1A1A2E")
        hdr_font = Font(bold=True, color="00E5A0", size=10)
        for cell in ws[1]:
            cell.fill=hdr_fill; cell.font=hdr_font
            cell.alignment=Alignment(horizontal="center")
        for col in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len+4, 45)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="gst_details.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    print("\n" + "="*50)
    print("  GST Extractor  →  http://localhost:5000")
    print(f"  Ollama AI: {'✅ ENABLED  (model: '+OLLAMA_MODEL+')' if OLLAMA_ENABLED else '❌ OFF — regex only'}")
    print("="*50 + "\n")
    app.run(debug=True, port=5000)
