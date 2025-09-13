import os
import re
import json
import threading
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- External libraries ---
import pdfplumber
import easyocr
from docx import Document
import dateparser
from dateparser.search import search_dates
import spacy
from spacy_download import load_spacy
from pdf2image import convert_from_path

# Initialize
reader = easyocr.Reader(['en', 'af']) # English + Afrikaans
nlp = load_spacy("en_core_web_sm")

# ----------------------------- Text extraction -----------------------------
def extract_text_from_pdf(path: Path) -> str:
    """Extract text from a PDF using pdfplumber, fallback to EasyOCR."""
    text_chunks = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    text_chunks.append(text)
    except Exception:
        pass

    text = "\n".join(text_chunks).strip()

    # Use EasyOCR if no text extracted
    if not text:
        images = convert_from_path(str(path))
        ocr_texts = []
        for img in images:
            result = reader.readtext(img)
            page_text = " ".join([item[1] for item in result])
            ocr_texts.append(page_text)
        text = "\n".join(ocr_texts)
    return text


def extract_text_from_docx(path: Path) -> str:
    doc = Document(str(path))
    return "\n".join([p.text for p in doc.paragraphs])

# ----------------------------- SA-specific regex -----------------------------
DATE_REGEX = re.compile(r"\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b")
MONEY_REGEX = re.compile(r"\b(?:R|ZAR)?\s?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?\b")

CLAUSE_KEYWORDS = {
    'rent': ['rent', 'monthly rental', 'base rent', 'huur'],
    'deposit': ['security deposit', 'deposit', 'borg'],
    'termination': ['termination', 'cancel', 'notice period', 'opzegging'],
    'maintenance': ['maintenance', 'repair', 'landlord shall', 'tenant shall'],
    'utilities': ['utilities', 'electricity', 'water', 'municipal rates'],
    'law': ['Rental Housing Act', 'Consumer Protection Act', 'CPA', 'South African law'],
    'governing_law': ['governing law', 'jurisdiction', 'South Africa']
}

# ----------------------------- Extraction logic -----------------------------
def find_dates(text: str):
    parsed = search_dates(text, settings={'DATE_ORDER': 'DMY'})
    return list({d[1].date().isoformat() for d in parsed}) if parsed else []

def find_money(text: str):
    return list(set([m.group(0) for m in MONEY_REGEX.finditer(text)]))

def extract_parties(text: str):
    pattern = re.compile(r"between\s+(.*?)\s+and\s+(.*?)\s", re.I)
    m = pattern.search(text[:2000])
    if m:
        return [m.group(1).strip(), m.group(2).strip()]
    return []

def clause_search(text: str, keywords: list):
    text_lower = text.lower()
    results = []
    for kw in keywords:
        if kw.lower() in text_lower:
            start = max(0, text_lower.find(kw.lower()) - 100)
            end = min(len(text), text_lower.find(kw.lower()) + 300)
            results.append(text[start:end])
    return results

def extract_lease_fields(text: str) -> dict:
    result = {
        'dates': find_dates(text),
        'monetary_values': find_money(text),
        'parties': extract_parties(text),
        'clauses': {},
        'compliance_flags': []
    }
    for key, kws in CLAUSE_KEYWORDS.items():
        result['clauses'][key] = clause_search(text, kws)

    # Check deposit compliance
    deposits = [float(re.sub(r'[^\d.]', '', val)) for val in result['monetary_values'] if 'R' in val or 'ZAR' in val]
    if deposits:
        max_deposit = max(deposits)
        rent_candidates = [float(re.sub(r'[^\d.]', '', val)) for val in result['monetary_values']]
        if rent_candidates and max_deposit > (2 * max(rent_candidates)):
            result['compliance_flags'].append('Deposit exceeds 2 months rent')

    # Ensure governing law is specified
    if not result['clauses']['governing_law']:
        result['compliance_flags'].append('Missing governing law clause')

    return result

def compute_health_score(extracted: dict) -> int:
    score = 100
    if not extracted['parties']:
        score -= 15
    if not extracted['clauses']['rent']:
        score -= 20
    if not extracted['clauses']['deposit']:
        score -= 10
    if not extracted['clauses']['termination']:
        score -= 10
    if 'Deposit exceeds 2 months rent' in extracted['compliance_flags']:
        score -= 15
    if 'Missing governing law clause' in extracted['compliance_flags']:
        score -= 10
    return max(0, score)

# ----------------------------- Processing -----------------------------
def process_file(path: Path):
    ext = path.suffix.lower()
    if ext == '.pdf':
        text = extract_text_from_pdf(path)
    elif ext == '.docx':
        text = extract_text_from_docx(path)
    else:
        text = path.read_text(encoding='utf-8', errors='ignore')

    extracted = extract_lease_fields(text)
    extracted['file_name'] = path.name
    extracted['health_score'] = compute_health_score(extracted)
    return extracted

def process_all_files(input_folder, output_folder, progress_callback=None):
    in_path = Path(input_folder)
    out_path = Path(output_folder)
    out_path.mkdir(parents=True, exist_ok=True)

    results = []
    files = [p for p in in_path.iterdir() if p.is_file() and p.suffix.lower() in ['.pdf', '.docx', '.txt']]
    for i, p in enumerate(files):
        if progress_callback:
            progress_callback(f"Processing: {p.name}")
        data = process_file(p)
        results.append(data)
        with open(out_path / (p.stem + '.json'), 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    df = pd.DataFrame(results)
    df.to_csv(out_path / 'summary.csv', index=False)
    if progress_callback:
        progress_callback(f"Processing complete! {len(results)} files analyzed.")
    return df

# ----------------------------- Tkinter GUI -----------------------------
class LeaseAnalyticsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("South African Lease Analytics Tool")

        # Heading
        tk.Label(root, text="South African Lease Analytics Tool", font=("Arial", 16, "bold")).pack(pady=10)

        # Input folder
        frame_input = tk.Frame(root)
        frame_input.pack(pady=5)
        tk.Label(frame_input, text="Input Folder: ").pack(side=tk.LEFT)
        self.input_path_var = tk.StringVar()
        tk.Entry(frame_input, textvariable=self.input_path_var, width=50).pack(side=tk.LEFT)
        tk.Button(frame_input, text="Browse", command=self.select_input_folder).pack(side=tk.LEFT, padx=5)

        # Output folder
        frame_output = tk.Frame(root)
        frame_output.pack(pady=5)
        tk.Label(frame_output, text="Output Folder: ").pack(side=tk.LEFT)
        self.output_path_var = tk.StringVar()
        tk.Entry(frame_output, textvariable=self.output_path_var, width=50).pack(side=tk.LEFT)
        tk.Button(frame_output, text="Browse", command=self.select_output_folder).pack(side=tk.LEFT, padx=5)

        # Status
        self.status_var = tk.StringVar(value="Idle")
        tk.Label(root, textvariable=self.status_var, fg="blue").pack(pady=5)

        # Start button
        tk.Button(root, text="Start Processing", command=self.start_processing, bg="green", fg="white").pack(pady=10)

        # Table for results
        self.table_columns = (
            "File", "Health Score", "Parties", "Dates", "Monetary Values",
            "Rent Clause", "Deposit Clause", "Termination Clause", "Compliance Flags"
        )
        self.table = ttk.Treeview(root, columns=self.table_columns, show="headings", height=12)

        for col in self.table_columns:
            self.table.heading(col, text=col)
            self.table.column(col, anchor=tk.W, width=150)

        # Horizontal scrollbar
        scroll_x = ttk.Scrollbar(root, orient="horizontal", command=self.table.xview)
        self.table.configure(xscrollcommand=scroll_x.set)
        scroll_x.pack(fill=tk.X, side=tk.BOTTOM)

        self.table.pack(fill=tk.BOTH, expand=True, pady=10)

    def select_input_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_path_var.set(folder)

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_path_var.set(folder)

    def start_processing(self):
        input_folder = self.input_path_var.get()
        output_folder = self.output_path_var.get()

        if not input_folder or not output_folder:
            messagebox.showerror("Error", "Please select both input and output folders")
            return

        threading.Thread(target=self.run_processing, args=(input_folder, output_folder), daemon=True).start()

    def run_processing(self, input_folder, output_folder):
        def update_status(msg):
            self.status_var.set(msg)
            self.root.update_idletasks()

        df = process_all_files(input_folder, output_folder, progress_callback=update_status)

        # Clear old table rows
        for row in self.table.get_children():
            self.table.delete(row)

        # Insert new rows
        for _, r in df.iterrows():
            self.table.insert("", tk.END, values=(
                r.get('file_name', ''),
                r.get('health_score', ''),
                ", ".join(r.get('parties', [])),
                ", ".join(r.get('dates', [])),
                ", ".join(r.get('monetary_values', [])),
                " | ".join(r.get('clauses', {}).get('rent', []))[:100],
                " | ".join(r.get('clauses', {}).get('deposit', []))[:100],
                " | ".join(r.get('clauses', {}).get('termination', []))[:100],
                ", ".join(r.get('compliance_flags', []))
            ))


# ----------------------------- Main -----------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = LeaseAnalyticsGUI(root)
    root.geometry("900x600")
    root.mainloop()

