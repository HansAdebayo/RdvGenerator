import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
import unicodedata

def sanitize_filename(name):
    return name.replace(" ", "_").replace("/", "-")

def normalize(text):
    return ''.join(c for c in unicodedata.normalize('NFD', str(text)) if unicodedata.category(c) != 'Mn').lower()

def detect_column(columns, keyword):
    keyword_norm = normalize(keyword)
    for col in columns:
        if keyword_norm in normalize(col):
            return col
    return None

def load_rdv_data(path):
    df = pd.read_excel(path)
    col_year = detect_column(df.columns, 'annee')
    col_month = detect_column(df.columns, 'mois')
    col_day = detect_column(df.columns, 'jour')
    col_com = detect_column(df.columns, 'commercial')
    col_reason = detect_column(df.columns, 'raison')
    col_address = detect_column(df.columns, 'adresse')

    df['Date_RDV'] = pd.to_datetime(df[[col_year, col_month, col_day]])
    return df, col_com, col_reason, col_address

def filter_rdv(df, year, month, day_start, day_end):
    df = df[(df['Date_RDV'].dt.year == year) & (df['Date_RDV'].dt.month == month)]
    df = df[(df['Date_RDV'].dt.day >= day_start) & (df['Date_RDV'].dt.day <= day_end)]
    return df

def generate_rdv_report(commercial, df_rdv, logo_path, report_date, col_reason, col_address, output_dir):
    doc = Document()
    header = doc.sections[0].header.paragraphs[0]
    if logo_path and os.path.exists(logo_path):
        run = header.add_run()
        run.add_picture(logo_path, width=Inches(1.5))
    header.add_run(f"   Compte rendu du {report_date.strftime('%d %B %Y')} – Réunion commerciale {commercial}")
    header.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    doc.add_heading("1 - Agenda", level=1)
    doc.add_paragraph(f"Agenda des RDV du {df_rdv['Date_RDV'].min().strftime('%d %B %Y')} au {df_rdv['Date_RDV'].max().strftime('%d %B %Y')}")

    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    headers = ['Date', 'Raison du RDV', 'Adresse du RDV']
    for i, cell in enumerate(table.rows[0].cells):
        cell.text = headers[i]
        cell.paragraphs[0].runs[0].font.bold = True
        shade = OxmlElement('w:shd')
        shade.set(qn('w:fill'), 'D9E1F2')
        cell._tc.get_or_add_tcPr().append(shade)

    for _, row in df_rdv.iterrows():
        r = table.add_row().cells
        r[0].text = row['Date_RDV'].strftime('%d/%m/%Y')
        r[1].text = str(row[col_reason])
        r[2].text = str(row[col_address]) if pd.notnull(row[col_address]) else 'Non précisé'

    os.makedirs(output_dir, exist_ok=True)
    filename = f"{output_dir}/RDV_{sanitize_filename(commercial)}_{report_date.strftime('%Y-%m-%d')}.docx"
    doc.save(filename)
    return filename
