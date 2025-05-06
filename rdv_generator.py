# === rdv_generator.py ===
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
import unicodedata


def sanitize_filename(name):
    return name.replace(" ", "_").replace("/", "-")

def normalize(text):
    return ''.join(c for c in unicodedata.normalize('NFD', str(text)) if unicodedata.category(c) != 'Mn').lower().replace('_', ' ').replace('-', ' ')

def detect_column(columns, keyword):
    keyword_norm = normalize(keyword)
    for col in columns:
        if keyword_norm in normalize(col):
            return col
    return None

def load_rdv_data(path):
    df = pd.read_excel(path)

    col_year = detect_column(df.columns, 'annee') or detect_column(df.columns, 'year')
    col_month = detect_column(df.columns, 'mois') or detect_column(df.columns, 'month')
    col_day = detect_column(df.columns, 'jour') or detect_column(df.columns, 'day')

    if not (col_year and col_month and col_day):
        raise ValueError(f"Colonnes date manquantes : annee={col_year}, mois={col_month}, jour={col_day}")

    df['Date_RDV'] = pd.to_datetime(df[[col_year, col_month, col_day]])
    return df

def filter_rdv_between_dates(df, date_start, date_end):
    return df[(df['Date_RDV'] >= date_start) & (df['Date_RDV'] <= date_end)]

def add_logo_and_header(doc, logo_path, commercial_name, report_date):
    header = doc.sections[0].header
    paragraph = header.paragraphs[0]
    if logo_path and os.path.exists(logo_path):
        run = paragraph.add_run()
        run.add_picture(logo_path, width=Inches(1.5))
    paragraph.add_run(f"   Compte rendu du {report_date.strftime('%d %B %Y')} OBJET : réunion périodique commerciale {commercial_name}")
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

def add_agenda_section(doc, start_date, end_date, rdv_data):
    doc.add_paragraph(f"Date : {start_date.strftime('%d %B %Y')}\nRédacteur : Adeline et Laurent\nLieu : Visio")
    doc.add_paragraph()
    doc.add_heading("1- Agenda", level=1)
    doc.add_paragraph(f"RDV Gecko Agenda Semaine du {start_date.strftime('%d %B %Y')} au {end_date.strftime('%d %B %Y')} :")

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    headers = ['Date', 'Raison du RDV', 'Adresse du RDV']

    for i, cell in enumerate(hdr_cells):
        cell.text = headers[i]
        cell.paragraphs[0].runs[0].font.bold = True
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), 'D9E1F2')
        cell._tc.get_or_add_tcPr().append(shading_elm)

    for _, row in rdv_data.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = row['Date_RDV'].strftime('%d/%m/%Y')
        row_cells[1].text = str(row.get('Raison_du_RDV', ''))
        row_cells[2].text = str(row.get('Adresse_du_rdv', '')) if pd.notnull(row.get('Adresse_du_rdv')) else 'Non précisé'
        for cell in row_cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)

def create_word_report(commercial, rdv_df, logo_path, report_date, output_dir):
    doc = Document()
    add_logo_and_header(doc, logo_path, commercial, report_date)
    doc.add_paragraph("")
    start_date = report_date
    end_date = report_date + timedelta(days=15)
    add_agenda_section(doc, start_date, end_date, rdv_df)

    sanitized_name = sanitize_filename(commercial)
    filename = f"RDV_{sanitized_name}_{report_date.strftime('%Y-%m-%d')}.docx"
    output_path = os.path.join(output_dir, filename)
    doc.save(output_path)
    return output_path
