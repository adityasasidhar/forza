import os
import re
import json
import docx
import pandas as pd
from groq import Groq
from pydantic import BaseModel

# Init Groq client
api_key = os.getenv("GROQ_API_KEY")
if not api_key:
    raise EnvironmentError("GROQ_API_KEY environment variable not set.")
client = Groq(api_key=api_key)

def read_docx(path):
    doc = docx.Document(path)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text.strip())
    return "\n".join(full_text)

def classify_balance_sheet(text: str):
    prompt = f"""
    You are a financial assistant.
    Input: an unstructured analytical balance sheet.
    Task: Reclassify all items into the EU Article 2424 balance sheet schema (CEE).
    Return the result as JSON with two main keys: ATTIVO and PASSIVO.
    Each section should have subcategories (A, B, C, D, E) with items and amounts.
    Example:
    {{
      "ATTIVO": {{
        "B) Immobilizzazioni immateriali": [
          {{"label": "Software", "amount": 3041.40}},
          ...
        ],
        "B) Immobilizzazioni materiali": [
          ...
        ]
      }},
      "PASSIVO": {{
        "A) Patrimonio netto": [...],
        "D) Debiti": [...]
      }}
    }}
    Balance sheet text:
    {text}
    """

    response = client.chat.completions.create(
        model="openai/gpt-oss-20b",  # or smaller one if needed
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )

    result = response.choices[0].message.content
    print(result)
    return result

import json


def extract_json(json_str):
    # Remove surrounding Python triple quotes if any
    cleaned = json_str.strip("'''").strip('"""').strip()

    # Extract the JSON block
    match = re.search(r"\{.*\}", cleaned, re.DOTALL)
    if match:
        return json.loads(match.group())
    else:
        raise ValueError("No JSON found in the Groq response")

def write_excel_from_json(json_str, output_file="reclassified.xlsx"):
    data = json.loads(json_str)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for section, items in data.items():  # ATTIVO, PASSIVO
            # Flatten subcategories
            rows = []
            for subcat, records in items.items():
                for rec in records:
                    rows.append({
                        "section": section,
                        "subcategory": subcat,
                        "label": rec.get("label"),
                        "amount": rec.get("amount")
                    })
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=section, index=False)

    print(f"✅ Reclassified Excel saved to {output_file}")


from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_RIGHT


def write_to_pdf(data_dict, output_path):
    """
    Write the reclassified balance sheet to a PDF file.

    Args:
        data_dict: Dictionary containing ATTIVO and PASSIVO sections
        output_path: Path where the PDF will be saved
    """
    # Create PDF document
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm
    )

    # Container for PDF elements
    elements = []

    # Define styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=20,
        alignment=TA_CENTER
    )

    section_style = ParagraphStyle(
        'SectionTitle',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#2c5f9e'),
        spaceAfter=12,
        spaceBefore=12
    )

    subsection_style = ParagraphStyle(
        'SubsectionTitle',
        parent=styles['Heading3'],
        fontSize=11,
        textColor=colors.HexColor('#444444'),
        spaceAfter=8,
        spaceBefore=8
    )

    # Add title
    title = Paragraph("Reclassified Balance Sheet - Article 2424 CEE", title_style)
    elements.append(title)
    elements.append(Spacer(1, 0.5 * cm))

    # Process ATTIVO (Assets)
    if "ATTIVO" in data_dict:
        section_title = Paragraph("ATTIVO (Assets)", section_style)
        elements.append(section_title)

        for category, items in data_dict["ATTIVO"].items():
            if items:  # Only process if there are items
                # Add category header
                category_header = Paragraph(category, subsection_style)
                elements.append(category_header)

                # Create table data
                table_data = [["Description", "Amount (€)"]]

                for item in items:
                    label = item.get("label", "")
                    amount = item.get("amount", 0)
                    table_data.append([label, f"{amount:,.2f}"])

                # Calculate subtotal
                subtotal = sum(item.get("amount", 0) for item in items)
                table_data.append(["Subtotal", f"{subtotal:,.2f}"])

                # Create and style table
                table = Table(table_data, colWidths=[12 * cm, 4 * cm])
                table.setStyle(TableStyle([
                    # Header row
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c5f9e')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),

                    # Data rows
                    ('BACKGROUND', (0, 1), (-1, -2), colors.beige),
                    ('TEXTCOLOR', (0, 1), (-1, -2), colors.black),
                    ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
                    ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -2), 9),
                    ('GRID', (0, 0), (-1, -2), 0.5, colors.grey),

                    # Subtotal row
                    ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#d9e8f5')),
                    ('TEXTCOLOR', (0, -1), (-1, -1), colors.black),
                    ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, -1), (-1, -1), 10),
                    ('LINEABOVE', (0, -1), (-1, -1), 1.5, colors.black),
                ]))

                elements.append(table)
                elements.append(Spacer(1, 0.3 * cm))

        # Calculate total ATTIVO
        total_attivo = 0
        for category, items in data_dict["ATTIVO"].items():
            total_attivo += sum(item.get("amount", 0) for item in items)

        total_table = Table([["TOTAL ATTIVO", f"{total_attivo:,.2f}"]], colWidths=[12 * cm, 4 * cm])
        total_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#1f4788')),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.whitesmoke),
            ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
        ]))
        elements.append(total_table)

    # Add page break before PASSIVO
    elements.append(PageBreak())

    # Process PASSIVO (Liabilities)
    if "PASSIVO" in data_dict:
        section_title = Paragraph("PASSIVO (Liabilities & Equity)", section_style)
        elements.append(section_title)

        for category, items in data_dict["PASSIVO"].items():
            if items:  # Only process if there are items
                # Add category header
                category_header = Paragraph(category, subsection_style)
                elements.append(category_header)

                # Create table data
                table_data = [["Description", "Amount (€)"]]

                for item in items:
                    label = item.get("label", "")
                    amount = item.get("amount", 0)
                    table_data.append([label, f"{amount:,.2f}"])

                # Calculate subtotal
                subtotal = sum(item.get("amount", 0) for item in items)
                table_data.append(["Subtotal", f"{subtotal:,.2f}"])

                # Create and style table
                table = Table(table_data, colWidths=[12 * cm, 4 * cm])
                table.setStyle(TableStyle([
                    # Header row
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c5f9e')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),

                    # Data rows
                    ('BACKGROUND', (0, 1), (-1, -2), colors.beige),
                    ('TEXTCOLOR', (0, 1), (-1, -2), colors.black),
                    ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
                    ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -2), 9),
                    ('GRID', (0, 0), (-1, -2), 0.5, colors.grey),

                    # Subtotal row
                    ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#d9e8f5')),
                    ('TEXTCOLOR', (0, -1), (-1, -1), colors.black),
                    ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, -1), (-1, -1), 10),
                    ('LINEABOVE', (0, -1), (-1, -1), 1.5, colors.black),
                ]))

                elements.append(table)
                elements.append(Spacer(1, 0.3 * cm))

        # Calculate total PASSIVO
        total_passivo = 0
        for category, items in data_dict["PASSIVO"].items():
            total_passivo += sum(item.get("amount", 0) for item in items)

        total_table = Table([["TOTAL PASSIVO", f"{total_passivo:,.2f}"]], colWidths=[12 * cm, 4 * cm])
        total_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#1f4788')),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.whitesmoke),
            ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
        ]))
        elements.append(total_table)
    doc.build(elements)
    print(f"PDF successfully created: {output_path}")

# 4. Run end-to-end
if __name__ == "__main__":
    if __name__ == "__main__":
        input_path = "data/file.docx"
        text = read_docx(input_path)
        result_json = classify_balance_sheet(text)
        result_json = extract_json(result_json)
        write_to_pdf(result_json, "reclassified_output.pdf")
