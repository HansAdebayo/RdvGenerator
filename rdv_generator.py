import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
import locale
import unicodedata

locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')


def normalize(text):
    return ''.join(c for c in unicodedata.normalize('NFD', str(text)) if unicodedata.category(c) != 'Mn').lower()


def detect_column(columns, keyword):
    keyword_norm = normalize(keyword)
    for col in columns:
        if keyword_norm in normalize(col):
            return col
    return None


def sanitize_filename(name):
    return name.replace(" ", "_").replace("/", "-")


def load_rdv_data(path):
    df = pd.read_excel(path)
    year_col = detect_column(df.columns, "annee")
    month_col = detect_column(df.columns, "mois")
    day_col = detect_column(df.columns, "jour")

    if not all([year_col, month_col, day_col]):
        raise ValueError("Les colonnes année, mois et jour sont requises.")

    df['Date_RDV'] = pd.to_datetime(dict(
        year=df[year_col],
        month=df[month_col],
        day=df[day_col]
    ))

    return df


def filter_rdv(df, start_date, end_date):
    return df[(df['Date_RDV'] >= start_date) & (df['Date_RDV'] <= end_date)]


def add_logo_and_header(doc, logo_path, commercial_name, report_date):
    header = doc.sections[0].header
    paragraph = header.paragraphs[0]

    if logo_path and os.path.exists(logo_path):
        run = paragraph.add_run()
        run.add_picture(logo_path, width=Inches(1.5))

    paragraph.add_run(f"   Compte rendu du {report_date.strftime('%d %B %Y')} – Réunion commerciale {commercial_name}")
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT


def add_agenda_section(doc, start_date, end_date, rdv_data):
    doc.add_paragraph(f"Date : {start_date.strftime('%d %B %Y')}\nRédacteur : Adeline et Laurent\nLieu : Visio")
    doc.add_paragraph()
    doc.add_heading("1- Agenda", level=1)
    doc.add_paragraph(f"RDV Gecko Agenda du {start_date.strftime('%d %B %Y')} au {end_date.strftime('%d %B %Y')} :")

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = 'Table Grid'

    headers = ['Date', 'Raison du RDV', 'Adresse du RDV']
    hdr_cells = table.rows[0].cells
    for i, head in enumerate(headers):
        hdr_cells[i].text = head
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), 'D9E1F2')
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading_elm)
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    col_raison = detect_column(rdv_data.columns, "raison")
    col_adresse = detect_column(rdv_data.columns, "adresse")

    for _, row in rdv_data.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = row['Date_RDV'].strftime('%d/%m/%Y')
        row_cells[1].text = str(row[col_raison]) if col_raison else ''
        row_cells[2].text = str(row[col_adresse]) if col_adresse and pd.notnull(row[col_adresse]) else 'Non précisé'
        for cell in row_cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def create_word_report(commercial, rdv_df, logo_path, report_date, output_dir):
    doc = Document()
    add_logo_and_header(doc, logo_path, commercial, report_date)
    doc.add_paragraph()
    end_date = report_date + timedelta(days=15)
    add_agenda_section(doc, report_date, end_date, rdv_df)

    os.makedirs(output_dir, exist_ok=True)
    filename = f"RDV_{sanitize_filename(commercial)}_{report_date.strftime('%Y-%m-%d')}.docx"
    path = os.path.join(output_dir, filename)
    doc.save(path)
    return path
