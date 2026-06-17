import json, os
def export_markdown(t, p):
    with open(p, 'w') as f: f.write(t)
def export_json(d, t, p):
    out = d.to_dict(); out['text'] = t
    with open(p, 'w') as f: json.dump(out, f)
def build_docx(t, p, d, sig_c=None, sig_cl=None, stamp_c=None, logo_path=None):
    from docx import Document
    doc = Document(); doc.add_paragraph(t); doc.save(p)
def build_pdf(t, p, d, sig_c=None, sig_cl=None):
    from fpdf import FPDF
    pdf = FPDF(); pdf.add_page(); pdf.set_font("Arial", size=12); pdf.multi_cell(0, 10, t); pdf.output(p)
