import os
from get_data import get_pn_for, get_sdr_for, load_weight_table, get_discount
from price_calculator import calculate_total_mass, calculate_price, calculate_length_from_mass, calculate_price_per_kg_from_total
from datetime import datetime
import requests
from num2words import num2words
from reportlab.lib.pagesizes import A4, landscape

# Define PAGE_WIDTH, PAGE_HEIGHT for landscape A4
PAGE_WIDTH, PAGE_HEIGHT = landscape(A4)
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from bidi.algorithm import get_display
from arabic_reshaper import reshape
from khayyam import JalaliDate
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from persiantools import digits
import json
import csv
from typing import Optional

# Use "program files" instead of "dependencies"
font_path = os.path.join(os.path.dirname(__file__), "program files", "DejaVuSans.ttf")
if os.path.exists(font_path):
    pdfmetrics.registerFont(TTFont("Persian", font_path))
    DEFAULT_FONT = "Persian"
else:
    DEFAULT_FONT = "Helvetica"

# Ensure program files directory exists
DEPENDENCIES_DIR = os.path.join(os.path.dirname(__file__), "program files")
os.makedirs(DEPENDENCIES_DIR, exist_ok=True)
# Constants
INVOICE_COUNTER_FILE = os.path.join(DEPENDENCIES_DIR, "invoice_counter.json")
COMPANY_NAME = "شرکت پلی غرب"

# Default company information used when optional fields are not provided
DEFAULT_COMPANY_INFO = {
    "name": "شرکت پلی غرب اتصال ایرانیان (سهامی خاص)",
    "address": "میرداماد میدان محسنی خیابان شاه نظری کوچه دوم پلاک ۳۴ واحد اول شمالی",
    "phone": "02122922999",
    "fax": "02122278331",
}


def fetch_current_jalali_date() -> str:
    """Return today's date in Jalali format using an internet time service."""
    try:
        resp = requests.get("https://worldtimeapi.org/api/ip", timeout=5)
        resp.raise_for_status()
        iso_dt = resp.json().get("datetime")
        dt = datetime.fromisoformat(iso_dt.rstrip("Z")) if iso_dt else datetime.utcnow()
    except Exception:
        dt = datetime.utcnow()
    return JalaliDate(dt).strftime("%Y/%m/%d")


def number_to_words(value: float) -> str:
    """Convert a numeric value to Persian words."""
    try:
        return num2words(int(value), lang="fa")
    except Exception:
        return str(value)


def append_total_words(elements, total: float):
    """Append the total amount in words to the elements list."""
    words = number_to_words(total)
    line = str(get_display(reshape(f"مبلغ به حروف: {words} تومان")))
    style = ParagraphStyle(name="TotalWords", fontName=DEFAULT_FONT, fontSize=10, alignment=TA_RIGHT)
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(line, style))


# Helper for invoice numbering
def _get_next_invoice_number():
    """Return the next invoice number as an unpadded string.

    The function reads ``invoice_counter.json`` from the ``program files``
    directory, increments the stored ``counter`` value and writes it back.
    The resulting invoice number is returned directly without any zero
    padding.
    """
    if os.path.exists(INVOICE_COUNTER_FILE):
        try:
            with open(INVOICE_COUNTER_FILE, 'r') as f:
                data = json.load(f)
        except (ValueError, KeyError):
            data = {'counter': 0}
    else:
        data = {'counter': 0}
    data['counter'] += 1
    with open(INVOICE_COUNTER_FILE, 'w') as f:
        json.dump(data, f)
    # Return as plain integer string (no leading zeros)
    return str(data['counter'])


def to_persian_digits(text):
    return digits.en_to_fa(str(text))

def generate_pdf(customer_name: str, invoice_number: str, items: list[dict], output_dir: Optional[str] = None, explanation_text: Optional[str] = None):
    """
    Generates a PDF invoice for a customer.

    Args:
        customer_name (str): The name of the customer.
        invoice_number (str): The invoice number.
        items (list[dict]): A list of item dictionaries with invoice details.
        output_dir (str, optional): Path to the directory where the PDF will be saved. Defaults to None, which uses the module's 'خروجی' directory.
        explanation_text (str, optional): Additional explanation text to include in the PDF.
    """
    # Determine and create output directory if it doesn't exist
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)
    date_jalali = fetch_current_jalali_date()
    now = datetime.now().strftime("%d-%m")
    # Prepare and shape header texts for RTL display
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    # Shape the customer name text for RTL
    sh_customer_name = str(get_display(reshape(customer_name)))
    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)

    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    # Helper to draw the logo directly onto the canvas so it doesn't affect flowables
    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            # Coordinates origin is at lower‑left; place near top‑left inside page margins
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60  # 60 is the logo height
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    elements = []
    # Add company title to the PDF
    title_style = ParagraphStyle(
        name="CompanyTitle",
        fontName=DEFAULT_FONT,
        fontSize=18,
        alignment=TA_CENTER,
        leading=22
    )
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))

    # Add customer and invoice information
    header = [
        ["", sh_date],
        ["", sh_inv],
        ["", f"{sh_customer_name}{label_cust}"]
    ]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))
    # Build the items table
    headers_text = ["شماره","قطر (mm)","SDR","گرید","طول (m)","وزن/متر (kg)","وزن کل (kg)","قیمت/kg","قیمت کل (تومان)"]
    headers = [ str(get_display(reshape(h))) for h in headers_text ]
    headers = list(reversed(headers))
    data = [headers]
    total_price_all = 0.0
    for idx, itm in enumerate(items, start=1):
        total_price_all += itm["total_price"]
        # Ensure grade uses English digits
        grade_val = itm["pe_grade"]
        grade_val = digits.fa_to_en(grade_val)
        row = [
            str(idx),
            str(int(itm['diameter'])),
            str(int(itm['sdr'])),
            grade_val,
            f"{itm['length']:.2f}",
            f"{itm['weight_per_meter']:.3f}",
            f"{itm['total_weight']:.3f}",
            f"{int(itm['price_per_kg']):,}",
            f"{int(itm['total_price']):,}",
        ]
        data.append(list(reversed(row)))
    # Add total price row
    sh_total_label = str(get_display(reshape("جمع کل:")))
    sh_total_price = str(get_display(reshape(f"{int(total_price_all):,}")))
    total_row = ["", "", "", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))
    tbl = Table(data, repeatRows=1, colWidths=[100,60,60,60,50,50,30,60,40])
    tbl.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 9),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('BACKGROUND', (0,-1), (1,-1), colors.lightgrey),
        ('ALIGN', (1, len(data)-1), (1, len(data)-1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))  # More space before explanation section

        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))

        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )

        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    # Write total amount in words
    append_total_words(elements, total_price_all)

    # Generate PDF with logo on pages
    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    # Invoice counter is managed by the GUI application
    print(f"Invoice PDF saved to: {pdf_file}")
    return pdf_file

def generate_pdf_with_added_value(
    customer_name: str, invoice_number: str, items: list[dict],
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None
):
    """
    Generates a PDF invoice for a customer with 10% added value applied.
    """
    # Determine output directory
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    # Dates and header shaping
    date_jalali = fetch_current_jalali_date()
    now = datetime.now().strftime("%d-%m")
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    # Safe filename
    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")

    # Document setup
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    # Build flowables
    elements = []
    title_style = ParagraphStyle(name="CompanyTitle", fontName=DEFAULT_FONT, fontSize=18, alignment=TA_CENTER, leading=22)
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))

    # Customer and invoice info table
    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Item table headers
    headers_text = ["شماره","قطر (mm)","SDR","گرید","طول (m)","وزن/متر (kg)","وزن کل (kg)","قیمت/kg","قیمت کل (تومان)"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]

    # Populate items and compute total
    total_price_all = 0.0
    for idx, itm in enumerate(items, start=1):
        total_price_all += itm["total_price"]
        # Ensure grade uses English digits
        grade_val = itm["pe_grade"]
        grade_val = digits.fa_to_en(grade_val)
        row = [
            str(idx),
            str(int(itm['diameter'])),
            str(int(itm['sdr'])),
            grade_val,
            f"{itm['length']:.2f}",
            f"{itm['weight_per_meter']:.3f}",
            f"{itm['total_weight']:.3f}",
            f"{int(itm['price_per_kg']):,}",
            f"{int(itm['total_price']):,}",
        ]
        data.append(list(reversed(row)))

    # Calculate and add value-added row (10%)
    added_value = total_price_all * 0.10
    sh_added_label = str(get_display(reshape("مالیات بر ارزش افزوده:")))
    sh_added_value = str(get_display(reshape(f"{int(added_value):,}")))
    added_row = ["", "", "", "", "", "", "", sh_added_label, sh_added_value]
    data.append(list(reversed(added_row)))

    # Final total including added value
    final_total = total_price_all + added_value
    sh_total_label = str(get_display(reshape("جمع کل:")))
    sh_total_price = str(get_display(reshape(f"{int(final_total):,}")))
    total_row = ["", "", "", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    # Create table with styles
    tbl = Table(data, repeatRows=1, colWidths=[100, 100, 60, 60, 50, 50, 30, 60, 40])
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 9),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        # Highlight only the added-value label and its number
        ('BACKGROUND', (0, -2), (1, -2), colors.lightgrey),
        # Highlight only the total label and its number
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN', (1, len(data)-1), (1, len(data)-1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))
    
    # Write total in words
    append_total_words(elements, final_total)

    # Build PDF
    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    # Invoice counter is managed by the GUI application
    print(f"Invoice PDF with added value saved to: {pdf_file}")
    return pdf_file


def generate_pdf_with_discount(
    customer_name: str,
    invoice_number: str,
    items: list[dict],
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None
):
    """
    Generates a PDF invoice for a customer with tiered discounts applied.
    """
    # Determine output directory
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    # Dates and header shaping
    date_jalali = fetch_current_jalali_date()
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    # Safe filename
    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")

    # Document setup
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    # Build flowables
    elements = []
    title_style = ParagraphStyle(name="CompanyTitle", fontName=DEFAULT_FONT, fontSize=18, alignment=TA_CENTER, leading=22)
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))
    # Header table
    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Item table headers
    headers_text = ["شماره","قطر (mm)","SDR","گرید","طول (m)","وزن/متر (kg)","وزن کل (kg)","قیمت/kg","قیمت کل (تومان)"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]

    # Populate items and compute total
    total_price_all = 0.0
    for idx, itm in enumerate(items, start=1):
        total_price_all += itm["total_price"]
        grade_val = digits.fa_to_en(itm["pe_grade"])
        row = [
            str(idx),
            str(int(itm['diameter'])),
            str(int(itm['sdr'])),
            grade_val,
            f"{itm['length']:.2f}",
            f"{itm['weight_per_meter']:.3f}",
            f"{itm['total_weight']:.3f}",
            f"{int(itm['price_per_kg']):,}",
            f"{int(itm['total_price']):,}",
        ]
        data.append(list(reversed(row)))

    # Load discount thresholds
    base_dir = os.path.dirname(__file__)
    discount_csv_path = os.path.join(base_dir, "program files", "discount.csv")
    thresholds = []
    if os.path.exists(discount_csv_path):
        with open(discount_csv_path, newline='') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 2:
                    try:
                        thresholds.append((float(row[0]), float(row[1])))
                    except ValueError:
                        continue
    thresholds.sort(key=lambda x: x[0])
    thresholds.insert(0, (0.0, 0.0))  # zero-threshold for no discount below first step

    # Calculate tiered discount amount
    discount_amount = 0.0
    for i in range(len(thresholds)):
        min_price, pct = thresholds[i]
        next_price = thresholds[i+1][0] if i+1 < len(thresholds) else total_price_all
        if total_price_all > min_price:
            segment = min(total_price_all, next_price) - min_price
        else:
            segment = 0.0
        discount_amount += segment * (pct / 100.0)

    # Add discount row
    sh_discount_label = str(get_display(reshape("تخفیف :")))
    sh_discount_value = str(get_display(reshape(f"{int(discount_amount):,}")))
    discount_row = ["", "", "", "", "", "", "", sh_discount_label, sh_discount_value]
    data.append(list(reversed(discount_row)))

    # Final total after discount
    final_total = total_price_all - discount_amount
    sh_total_label = str(get_display(reshape("جمع کل :")))
    sh_total_price = str(get_display(reshape(f"{int(final_total):,}")))
    total_row = ["", "", "", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    # Table styling
    tbl = Table(data, repeatRows=1, colWidths=[100,60,60,60,50,50,30,60,40])
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 9),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, -2), (1, -2), colors.lightgrey),
        ('ALIGN',    (1, -2), (1, -2), 'LEFT'),
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN',    (1, -1), (1, -1), 'LEFT'),
        ('ALIGN', (1, len(data)-1), (1, len(data)-1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    # Write total in words
    append_total_words(elements, final_total)

    # Build PDF
    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    # Invoice counter is managed by the GUI application
    print(f"Invoice PDF with discount saved to: {pdf_file}")
    return pdf_file


def generate_pdf_with_custom_discount(
    customer_name: str,
    invoice_number: str,
    items: list[dict],
    discount: float,
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None):
    """
    Generates a PDF invoice for a customer applying a custom discount.
    If discount is <= 100, treat it as a percentage; if > 100, treat it as an absolute amount.
    """
    # Determine output directory
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    # Dates and header shaping
    date_jalali = fetch_current_jalali_date()
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    # Safe filename
    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")

    # Document setup
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    elements = []
    title_style = ParagraphStyle(name="CompanyTitle", fontName=DEFAULT_FONT, fontSize=18, alignment=TA_CENTER, leading=22)
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))

    # Customer and invoice info
    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Item table headers
    headers_text = ["شماره","قطر (mm)","SDR","گرید","طول (m)","وزن/متر (kg)","وزن کل (kg)","قیمت/kg","قیمت کل (تومان)"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]

    # Populate items and compute total
    total_price_all = 0.0
    for idx, itm in enumerate(items, start=1):
        total_price_all += itm["total_price"]
        grade_val = digits.fa_to_en(itm["pe_grade"])
        row = [
            str(idx),
            str(int(itm['diameter'])),
            str(int(itm['sdr'])),
            grade_val,
            f"{itm['length']:.2f}",
            f"{itm['weight_per_meter']:.3f}",
            f"{itm['total_weight']:.3f}",
            f"{int(itm['price_per_kg']):,}",
            f"{int(itm['total_price']):,}",
        ]
        data.append(list(reversed(row)))

    # Determine discount amount
    if discount <= 100:
        discount_amount = total_price_all * (discount / 100.0)
    else:
        discount_amount = discount

    # Add discount row
    sh_discount_label = str(get_display(reshape("تخفیف :")))
    sh_discount_value = str(get_display(reshape(f"{int(discount_amount):,}")))
    discount_row = ["", "", "", "", "", "", "", sh_discount_label, sh_discount_value]
    data.append(list(reversed(discount_row)))

    # Final total after discount
    final_total = total_price_all - discount_amount
    sh_total_label = str(get_display(reshape("جمع کل :")))
    sh_total_price = str(get_display(reshape(f"{int(final_total):,}")))
    total_row = ["", "", "", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    # Create table with styles
    tbl = Table(data, repeatRows=1, colWidths=[100,60,60,60,50,50,30,60,40])
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 9),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, -2), (1, -2), colors.lightgrey),
        ('ALIGN', (1, -2), (1, -2), 'LEFT'),
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN', (1, -1), (1, -1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    # Write total in words
    append_total_words(elements, final_total)

    # Build PDF
    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    # Invoice counter is managed by the GUI application
    print(f"Invoice PDF with custom discount saved to: {pdf_file}")
    return pdf_file


def generate_pdf_with_discount_and_added_value(
    customer_name: str,
    invoice_number: str,
    items: list[dict],
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None
):
    """
    Generates a PDF invoice for a customer with tiered discounts and 10% added-value applied on the net amount.
    """
    # Determine output directory
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    # Dates and header shaping
    date_jalali = fetch_current_jalali_date()
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    # Safe filename
    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")

    # Document setup
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    # Build flowables
    elements = []
    title_style = ParagraphStyle(name="CompanyTitle", fontName=DEFAULT_FONT, fontSize=18, alignment=TA_CENTER, leading=22)
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))
    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Item table headers
    headers_text = ["شماره","قطر (mm)","SDR","گرید","طول (m)","وزن/متر (kg)","وزن کل (kg)","قیمت/kg","قیمت کل (تومان)"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]

    # Populate items and compute total
    total_price_all = 0.0
    for idx, itm in enumerate(items, start=1):
        total_price_all += itm["total_price"]
        grade_val = digits.fa_to_en(itm["pe_grade"])
        row = [
            str(idx),
            str(int(itm['diameter'])),
            str(int(itm['sdr'])),
            grade_val,
            f"{itm['length']:.2f}",
            f"{itm['weight_per_meter']:.3f}",
            f"{itm['total_weight']:.3f}",
            f"{int(itm['price_per_kg']):,}",
            f"{int(itm['total_price']):,}",
        ]
        data.append(list(reversed(row)))

    # Load discount thresholds
    base_dir = os.path.dirname(__file__)
    discount_csv_path = os.path.join(base_dir, "program files", "discount.csv")
    thresholds = []
    if os.path.exists(discount_csv_path):
        with open(discount_csv_path, newline='') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 2:
                    try:
                        thresholds.append((float(row[0]), float(row[1])))
                    except ValueError:
                        continue
    thresholds.sort(key=lambda x: x[0])
    thresholds.insert(0, (0.0, 0.0))

    # Calculate tiered discount amount
    discount_amount = 0.0
    for i in range(len(thresholds)):
        min_price, pct = thresholds[i]
        next_price = thresholds[i+1][0] if i+1 < len(thresholds) else total_price_all
        if total_price_all > min_price:
            segment = min(total_price_all, next_price) - min_price
        else:
            segment = 0.0
        discount_amount += segment * (pct / 100.0)

    # Add discount row
    sh_discount_label = str(get_display(reshape("تخفیف :")))
    sh_discount_value = str(get_display(reshape(f"{int(discount_amount):,}")))
    discount_row = ["", "", "", "", "", "", "", sh_discount_label, sh_discount_value]
    data.append(list(reversed(discount_row)))

    # Net amount after discount
    net_amount = total_price_all - discount_amount

    # Calculate added-value (10% on net amount)
    added_value_amount = net_amount * 0.10

    # Add added-value row
    sh_added_label = str(get_display(reshape("مالیات بر ارزش افزوده :")))
    sh_added_value = str(get_display(reshape(f"{int(added_value_amount):,}")))
    added_row = ["", "", "", "", "", "", "", sh_added_label, sh_added_value]
    data.append(list(reversed(added_row)))

    # Final total after discount and added value
    final_total = net_amount + added_value_amount
    sh_total_label = str(get_display(reshape("جمع کل :")))
    sh_total_price = str(get_display(reshape(f"{int(final_total):,}")))
    total_row = ["", "", "", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    # Table styling
    tbl = Table(data, repeatRows=1, colWidths=[100,105,60,60,50,50,30,60,40])
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 9),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, -3), (1, -3), colors.lightgrey),
        ('ALIGN', (1, -3), (1, -3), 'LEFT'),
        ('BACKGROUND', (0, -2), (1, -2), colors.lightgrey),
        ('ALIGN',    (1, -2), (1, -2), 'LEFT'),
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN', (1, len(data)-1), (1, len(data)-1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    # Write total in words
    append_total_words(elements, final_total)

    # Build PDF
    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    # Invoice counter is managed by the GUI application
    print(f"Invoice PDF with discount and added value saved to: {pdf_file}")
    return pdf_file


def generate_pdf_with_custom_discount_and_added_value(
    customer_name: str,
    invoice_number: str,
    items: list[dict],
    discount: float,
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None):
    """
    Generates a PDF invoice applying a custom discount (percentage if <=100, absolute if >100) and then adds 10% added-value on the net amount.
    """
    # Determine output directory
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    # Dates and header shaping
    date_jalali = fetch_current_jalali_date()
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    # Safe filename and document setup
    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    elements = []
    title_style = ParagraphStyle(name="CompanyTitle", fontName=DEFAULT_FONT, fontSize=18, alignment=TA_CENTER, leading=22)
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))

    # Customer and invoice info
    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Populate items and compute total
    headers_text = ["شماره","قطر (mm)","SDR","گرید","طول (m)","وزن/متر (kg)","وزن کل (kg)","قیمت/kg","قیمت کل (تومان)"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]
    total_price_all = 0.0
    for idx, itm in enumerate(items, start=1):
        total_price_all += itm["total_price"]
        grade_val = digits.fa_to_en(itm["pe_grade"])
        row = [
            str(idx), str(int(itm['diameter'])), str(int(itm['sdr'])), grade_val,
            f"{itm['length']:.2f}", f"{itm['weight_per_meter']:.3f}", f"{itm['total_weight']:.3f}",
            f"{int(itm['price_per_kg']):,}", f"{int(itm['total_price']):,}" ]
        data.append(list(reversed(row)))

    # Determine discount amount
    if discount <= 100:
        discount_amount = total_price_all * (discount / 100.0)
    else:
        discount_amount = discount

    # Add discount row
    sh_discount_label = str(get_display(reshape("تخفیف :")))
    sh_discount_value = str(get_display(reshape(f"{int(discount_amount):,}")))
    discount_row = ["", "", "", "", "", "", "", sh_discount_label, sh_discount_value]
    data.append(list(reversed(discount_row)))

    # Net amount after discount and 10% added-value
    net_amount = total_price_all - discount_amount
    added_value_amount = net_amount * 0.10
    sh_added_label = str(get_display(reshape("مالیات بر ارزش افزوده :")))
    sh_added_value = str(get_display(reshape(f"{int(added_value_amount):,}")))
    added_row = ["", "", "", "", "", "", "", sh_added_label, sh_added_value]
    data.append(list(reversed(added_row)))

    # Final total
    final_total = net_amount + added_value_amount
    sh_total_label = str(get_display(reshape("جمع کل :")))
    sh_total_price = str(get_display(reshape(f"{int(final_total):,}")))
    total_row = ["", "", "", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    tbl = Table(data, repeatRows=1, colWidths=[100,105,60,60,50,50,30,60,40])
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 9),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, -3), (1, -3), colors.lightgrey),
        ('ALIGN',  (1, -3), (1, -3), 'LEFT'),
        ('BACKGROUND', (0, -2), (1, -2), colors.lightgrey),
        ('ALIGN',  (1, -2), (1, -2), 'LEFT'),
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN',  (1, -1), (1, -1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    # Write total in words
    append_total_words(elements, final_total)

    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)

    # Invoice counter is managed by the GUI application
    print(f"Invoice PDF with custom discount and added value saved to: {pdf_file}")
    return pdf_file

if __name__ == "__main__":
    def main():
        customer_name = 'مشتری آزمایشی'
        if not customer_name:
            customer_name = "مشتری نمونه"

        invoice_number = _get_next_invoice_number()

        # Sample data (you can replace with dynamic input)
        items = [
            {
                "diameter": 110,
                "sdr": 17,
                "pe_grade": "PE80",
                "length": 120.5,
                "weight_per_meter": 0.95,
                "total_weight": 114.475,
                "price_per_kg": 600000,
                "total_price": 6868500,
            },
            {
                "diameter": 90,
                "sdr": 13.6,
                "pe_grade": "PE100",
                "length": 80,
                "weight_per_meter": 0.75,
                "total_weight": 60,
                "price_per_kg": 62000,
                "total_price": 3720000,
            }
        ]

        explanation = "این یک متن توضیحات نمونه است که در پایین فاکتور نمایش داده می‌شود.\nمی‌تواند شامل چندین خط باشد."

        generate_pdf_with_custom_discount_and_added_value(
            customer_name, invoice_number, items, discount=10, explanation_text=explanation
        )


    main()


def generate_connection_invoice_pdf(
    customer_name: str,
    invoice_number: str,
    items: list[dict],
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None
):
    """
    Generates a PDF invoice for connection items (اتصالات), with columns:
    نوع اتصال | محصول | سایز | تعداد | قیمت واحد | قیمت کل
    All headers are in Persian, RTL-shaped.
    """
    # Determine output directory
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    date_jalali = fetch_current_jalali_date()
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    elements = []
    title_style = ParagraphStyle(
        name="CompanyTitle",
        fontName=DEFAULT_FONT,
        fontSize=18,
        alignment=TA_CENTER,
        leading=22
    )
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))

    # Customer and invoice info table
    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Table headers: نوع اتصال | محصول | فشار قابل تحمل | سایز | تعداد | قیمت واحد | قیمت کل (RTL order, so reverse for display)
    headers_text = ["نوع اتصال", "محصول", "فشار قابل تحمل", "سایز", "تعداد", "قیمت واحد", "قیمت کل"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]

    total_price_all = 0.0
    for itm in items:
        # All values as string, prices as Persian digits with thousands separator
        type_val   = str(get_display(reshape(str(itm.get("type", "")))))
        product_val = str(get_display(reshape(str(itm.get("product", "")))))
        pn_val = str(get_display(reshape(str(itm.get("pn", "")))))
        size_val   = str(get_display(reshape(str(itm.get("size", "")))))
        quantity = itm.get("quantity", 1)
        quantity_val = str(get_display(reshape(str(int(quantity)))))
        unit_price = itm.get("unit_price", 0)
        total_price = itm.get("total_price", 0)
        total_price_all += total_price
        unit_price_str = str(get_display(reshape(f"{int(unit_price):,}")))
        total_price_str = str(get_display(reshape(f"{int(total_price):,}")))
        # Order: type, product, pn, size, quantity, unit_price, total_price
        row = [
            type_val,
            product_val,
            pn_val,
            size_val,
            quantity_val,
            unit_price_str,
            total_price_str,
        ]
        data.append(list(reversed(row)))

    # Add total row (align with new columns: [نوع اتصال, محصول, فشار قابل تحمل, سایز, تعداد, قیمت واحد, قیمت کل])
    sh_total_label = str(get_display(reshape("جمع کل")))
    sh_total_price = str(get_display(reshape(f"{int(total_price_all):,}")))
    total_row = ["", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    # --- Dynamically calculate column widths for all columns ---
    from reportlab.pdfbase.pdfmetrics import stringWidth

    col_font = DEFAULT_FONT
    col_font_size = 10

    # `data` already contains all rows (header, items, total) in display order: [قیمت کل, قیمت واحد, سایز, محصول, نوع اتصال]
    num_cols = len(data[0])
    col_widths = []
    for idx in range(num_cols):
        max_w = max(stringWidth(str(row[idx]), col_font, col_font_size) for row in data)
        col_widths.append(max_w + 20)  # add small margin

    # Create the table with auto-sized columns
    tbl = Table(
        data,
        repeatRows=1,
        colWidths=col_widths
    )
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 10),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN', (1, len(data)-1), (1, len(data)-1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    # Write total in words
    append_total_words(elements, total_price_all)

    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    print(f"Connection invoice PDF saved to: {pdf_file}")
    return pdf_file  


# --- New function: generate_connection_invoice_pdf_with_added_value ---
def generate_connection_invoice_pdf_with_added_value(
    customer_name: str,
    invoice_number: str,
    items: list[dict],
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None
):
    """
    Generates a PDF invoice for connection items with 10% added value applied.
    Columns: نوع اتصال | محصول | سایز | تعداد | قیمت واحد | قیمت کل
    """
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    from bidi.algorithm import get_display
    from arabic_reshaper import reshape
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from persiantools import digits
    from khayyam import JalaliDate

    date_jalali = fetch_current_jalali_date()
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    elements = []
    title_style = ParagraphStyle(
        name="CompanyTitle",
        fontName=DEFAULT_FONT,
        fontSize=18,
        alignment=TA_CENTER,
        leading=22
    )
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))

    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    headers_text = ["نوع اتصال", "محصول", "فشار قابل تحمل", "سایز", "تعداد", "قیمت واحد", "قیمت کل"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]

    total_price_all = 0.0
    for itm in items:
        type_val   = str(get_display(reshape(str(itm.get("type", "")))))
        product_val = str(get_display(reshape(str(itm.get("product", "")))))
        pn_val = str(get_display(reshape(str(itm.get("pn", "")))))
        size_val   = str(get_display(reshape(str(itm.get("size", "")))))
        quantity = itm.get("quantity", 1)
        quantity_val = str(get_display(reshape(str(int(quantity)))))
        unit_price = itm.get("unit_price", 0)
        total_price = itm.get("total_price", 0)
        total_price_all += total_price
        unit_price_str = str(get_display(reshape(f"{int(unit_price):,}")))
        total_price_str = str(get_display(reshape(f"{int(total_price):,}")))
        row = [
            type_val,
            product_val,
            pn_val,
            size_val,
            quantity_val,
            unit_price_str,
            total_price_str,
        ]
        data.append(list(reversed(row)))

    # Add 10% added value row
    added_value = total_price_all * 0.10
    sh_added_label = str(get_display(reshape("مالیات بر ارزش افزوده")))
    sh_added_value = str(get_display(reshape(f"{int(added_value):,}")))
    added_row = ["", "", "", "", "", sh_added_label, sh_added_value]
    data.append(list(reversed(added_row)))

    # Final total including added value
    final_total = total_price_all + added_value
    sh_total_label = str(get_display(reshape("جمع کل")))
    sh_total_price = str(get_display(reshape(f"{int(final_total):,}")))
    total_row = ["", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    from reportlab.pdfbase.pdfmetrics import stringWidth
    col_font = DEFAULT_FONT
    col_font_size = 10
    num_cols = len(data[0])
    col_widths = []
    for idx in range(num_cols):
        max_w = max(stringWidth(str(row[idx]), col_font, col_font_size) for row in data)
        col_widths.append(max_w + 20)

    tbl = Table(
        data,
        repeatRows=1,
        colWidths=col_widths
    )
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 10),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, -2), (1, -2), colors.lightgrey),
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN', (1, len(data)-1), (1, len(data)-1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    # Write total in words
    append_total_words(elements, final_total)

    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    print(f"Connection invoice PDF with added value saved to: {pdf_file}")
    return pdf_file


# --- New function: generate_connection_invoice_pdf_with_discount ---
def generate_connection_invoice_pdf_with_discount(
    customer_name: str,
    invoice_number: str,
    items: list[dict],
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None
):
    """
    Generates a PDF invoice for connection items with tiered discount applied.
    Columns: نوع اتصال | محصول | سایز | تعداد | قیمت واحد | قیمت کل
    Discount tiers loaded from discount.csv in program files.
    """
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    from bidi.algorithm import get_display
    from arabic_reshaper import reshape
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from persiantools import digits
    from khayyam import JalaliDate
    import csv

    date_jalali = fetch_current_jalali_date()
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    elements = []
    title_style = ParagraphStyle(
        name="CompanyTitle",
        fontName=DEFAULT_FONT,
        fontSize=18,
        alignment=TA_CENTER,
        leading=22
    )
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))

    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    headers_text = ["نوع اتصال", "محصول", "فشار قابل تحمل", "سایز", "تعداد", "قیمت واحد", "قیمت کل"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]

    total_price_all = 0.0
    for itm in items:
        type_val   = str(get_display(reshape(str(itm.get("type", "")))))
        product_val = str(get_display(reshape(str(itm.get("product", "")))))
        pn_val = str(get_display(reshape(str(itm.get("pn", "")))))
        size_val   = str(get_display(reshape(str(itm.get("size", "")))))
        quantity = itm.get("quantity", 1)
        quantity_val = str(get_display(reshape(str(int(quantity)))))
        unit_price = itm.get("unit_price", 0)
        total_price = itm.get("total_price", 0)
        total_price_all += total_price
        unit_price_str = str(get_display(reshape(f"{int(unit_price):,}")))
        total_price_str = str(get_display(reshape(f"{int(total_price):,}")))
        row = [
            type_val,
            product_val,
            pn_val,
            size_val,
            quantity_val,
            unit_price_str,
            total_price_str,
        ]
        data.append(list(reversed(row)))

    # --- Discount calculation using discount.csv ---
    discount_csv_path = os.path.join(DEPENDENCIES_DIR, "discount.csv")
    thresholds = []
    if os.path.exists(discount_csv_path):
        with open(discount_csv_path, newline='') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 2:
                    try:
                        thresholds.append((float(row[0]), float(row[1])))
                    except ValueError:
                        continue
    thresholds.sort(key=lambda x: x[0])
    thresholds.insert(0, (0.0, 0.0))  # for no discount below first threshold

    # Calculate tiered discount
    discount_amount = 0.0
    for i in range(len(thresholds)):
        min_price, pct = thresholds[i]
        next_price = thresholds[i+1][0] if i+1 < len(thresholds) else total_price_all
        if total_price_all > min_price:
            segment = min(total_price_all, next_price) - min_price
        else:
            segment = 0.0
        discount_amount += segment * (pct / 100.0)

    sh_discount_label = str(get_display(reshape("تخفیف")))
    sh_discount_value = str(get_display(reshape(f"{int(discount_amount):,}")))
    discount_row = ["", "", "", "", "", sh_discount_label, sh_discount_value]
    data.append(list(reversed(discount_row)))

    # Final total after discount
    final_total = total_price_all - discount_amount
    sh_total_label = str(get_display(reshape("جمع کل")))
    sh_total_price = str(get_display(reshape(f"{int(final_total):,}")))
    total_row = ["", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    from reportlab.pdfbase.pdfmetrics import stringWidth
    col_font = DEFAULT_FONT
    col_font_size = 10
    num_cols = len(data[0])
    col_widths = []
    for idx in range(num_cols):
        max_w = max(stringWidth(str(row[idx]), col_font, col_font_size) for row in data)
        col_widths.append(max_w + 20)

    tbl = Table(
        data,
        repeatRows=1,
        colWidths=col_widths
    )
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 10),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, -2), (1, -2), colors.lightgrey),
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN', (1, len(data)-1), (1, len(data)-1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    # Write total in words
    append_total_words(elements, final_total)

    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    print(f"Connection invoice PDF with discount saved to: {pdf_file}")
    return pdf_file


# --- New function: generate_connection_invoice_pdf_with_custom_discount ---
def generate_connection_invoice_pdf_with_custom_discount(
    customer_name: str,
    invoice_number: str,
    items: list[dict],
    discount: float,
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None
):
    """
    Generates a PDF invoice for connection items applying a custom discount.
    If discount is <= 100, treat as percent; if > 100, treat as absolute amount.
    Columns: نوع اتصال | محصول | سایز | تعداد | قیمت واحد | قیمت کل
    """
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    from bidi.algorithm import get_display
    from arabic_reshaper import reshape
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from persiantools import digits
    from khayyam import JalaliDate

    date_jalali = fetch_current_jalali_date()
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    elements = []
    title_style = ParagraphStyle(
        name="CompanyTitle",
        fontName=DEFAULT_FONT,
        fontSize=18,
        alignment=TA_CENTER,
        leading=22
    )
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))

    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Updated headers: add "فشار قابل تحمل"
    headers_text = ["نوع اتصال", "محصول", "فشار قابل تحمل", "سایز", "تعداد", "قیمت واحد", "قیمت کل"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]

    total_price_all = 0.0
    for itm in items:
        type_val   = str(get_display(reshape(str(itm.get("type", "")))))
        product_val = str(get_display(reshape(str(itm.get("product", "")))))
        pn_val = str(get_display(reshape(str(itm.get("pn", "")))))
        size_val   = str(get_display(reshape(str(itm.get("size", "")))))
        quantity = itm.get("quantity", 1)
        quantity_val = str(get_display(reshape(str(int(quantity)))))
        unit_price = itm.get("unit_price", 0)
        total_price = itm.get("total_price", 0)
        total_price_all += total_price
        unit_price_str = str(get_display(reshape(f"{int(unit_price):,}")))
        total_price_str = str(get_display(reshape(f"{int(total_price):,}")))
        # Order: type, product, pn, size, quantity, unit_price, total_price
        row = [
            type_val,
            product_val,
            pn_val,
            size_val,
            quantity_val,
            unit_price_str,
            total_price_str,
        ]
        data.append(list(reversed(row)))

    # Custom discount
    if discount <= 100:
        discount_amount = total_price_all * (discount / 100.0)
    else:
        discount_amount = discount

    sh_discount_label = str(get_display(reshape("تخفیف")))
    sh_discount_value = str(get_display(reshape(f"{int(discount_amount):,}")))
    discount_row = ["", "", "", "", "", sh_discount_label, sh_discount_value]
    data.append(list(reversed(discount_row)))

    # Final total after discount
    final_total = total_price_all - discount_amount
    sh_total_label = str(get_display(reshape("جمع کل")))
    sh_total_price = str(get_display(reshape(f"{int(final_total):,}")))
    total_row = ["", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    from reportlab.pdfbase.pdfmetrics import stringWidth
    col_font = DEFAULT_FONT
    col_font_size = 10
    num_cols = len(data[0])
    col_widths = []
    for idx in range(num_cols):
        max_w = max(stringWidth(str(row[idx]), col_font, col_font_size) for row in data)
        col_widths.append(max_w + 20)

    tbl = Table(
        data,
        repeatRows=1,
        colWidths=col_widths
    )
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 10),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, -2), (1, -2), colors.lightgrey),
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN', (1, len(data)-1), (1, len(data)-1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    # Write total in words
    append_total_words(elements, final_total)

    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    print(f"Connection invoice PDF with custom discount saved to: {pdf_file}")
    return pdf_file


# --- New function: generate_connection_invoice_pdf_with_discount_and_added_value ---
def generate_connection_invoice_pdf_with_discount_and_added_value(
    customer_name: str,
    invoice_number: str,
    items: list[dict],
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None
):
    """
    Generates a PDF invoice for connection items with tiered discount and 10% added value.
    Columns: نوع اتصال | محصول | سایز | تعداد | قیمت واحد | قیمت کل
    Discount tiers loaded from discount.csv in program files.
    """
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    from bidi.algorithm import get_display
    from arabic_reshaper import reshape
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from persiantools import digits
    from khayyam import JalaliDate
    import csv

    date_jalali = fetch_current_jalali_date()
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(DEPENDENCIES_DIR, "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = PAGE_HEIGHT - _doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    elements = []
    title_style = ParagraphStyle(
        name="CompanyTitle",
        fontName=DEFAULT_FONT,
        fontSize=18,
        alignment=TA_CENTER,
        leading=22
    )
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))

    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 12),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Updated headers: add "فشار قابل تحمل"
    headers_text = ["نوع اتصال", "محصول", "فشار قابل تحمل", "سایز", "تعداد", "قیمت واحد", "قیمت کل"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]

    total_price_all = 0.0
    for itm in items:
        type_val   = str(get_display(reshape(str(itm.get("type", "")))))
        product_val = str(get_display(reshape(str(itm.get("product", "")))))
        pn_val = str(get_display(reshape(str(itm.get("pn", "")))))
        size_val   = str(get_display(reshape(str(itm.get("size", "")))))
        quantity = itm.get("quantity", 1)
        quantity_val = str(get_display(reshape(str(int(quantity)))))
        unit_price = itm.get("unit_price", 0)
        total_price = itm.get("total_price", 0)
        total_price_all += total_price
        unit_price_str = str(get_display(reshape(f"{int(unit_price):,}")))
        total_price_str = str(get_display(reshape(f"{int(total_price):,}")))
        # Order: type, product, pn, size, quantity, unit_price, total_price
        row = [
            type_val,
            product_val,
            pn_val,
            size_val,
            quantity_val,
            unit_price_str,
            total_price_str,
        ]
        data.append(list(reversed(row)))

    # --- Discount calculation using discount.csv ---
    discount_csv_path = os.path.join(DEPENDENCIES_DIR, "discount.csv")
    thresholds = []
    if os.path.exists(discount_csv_path):
        with open(discount_csv_path, newline='') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 2:
                    try:
                        thresholds.append((float(row[0]), float(row[1])))
                    except ValueError:
                        continue
    thresholds.sort(key=lambda x: x[0])
    thresholds.insert(0, (0.0, 0.0))  # for no discount below first threshold

    # Calculate tiered discount
    discount_amount = 0.0
    for i in range(len(thresholds)):
        min_price, pct = thresholds[i]
        next_price = thresholds[i+1][0] if i+1 < len(thresholds) else total_price_all
        if total_price_all > min_price:
            segment = min(total_price_all, next_price) - min_price
        else:
            segment = 0.0
        discount_amount += segment * (pct / 100.0)

    sh_discount_label = str(get_display(reshape("تخفیف")))
    sh_discount_value = str(get_display(reshape(f"{int(discount_amount):,}")))
    discount_row = ["", "", "", "", "", sh_discount_label, sh_discount_value]
    data.append(list(reversed(discount_row)))

    # Net after discount
    net_after_discount = total_price_all - discount_amount

    # Add 10% added value row on net amount
    added_value = net_after_discount * 0.10
    sh_added_label = str(get_display(reshape("مالیات بر ارزش افزوده")))
    sh_added_value = str(get_display(reshape(f"{int(added_value):,}")))
    added_row = ["", "", "", "", "", sh_added_label, sh_added_value]
    data.append(list(reversed(added_row)))

    # Final total including added value
    final_total = net_after_discount + added_value
    sh_total_label = str(get_display(reshape("جمع کل")))
    sh_total_price = str(get_display(reshape(f"{int(final_total):,}")))
    total_row = ["", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    from reportlab.pdfbase.pdfmetrics import stringWidth
    col_font = DEFAULT_FONT
    col_font_size = 10
    num_cols = len(data[0])
    col_widths = []
    for idx in range(num_cols):
        max_w = max(stringWidth(str(row[idx]), col_font, col_font_size) for row in data)
        col_widths.append(max_w + 20)

    tbl = Table(
        data,
        repeatRows=1,
        colWidths=col_widths
    )
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), DEFAULT_FONT, 10),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, -3), (1, -3), colors.lightgrey),
        ('BACKGROUND', (0, -2), (1, -2), colors.lightgrey),
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN', (1, len(data)-1), (1, len(data)-1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName=DEFAULT_FONT, fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    print(f"Connection invoice PDF with discount and added value saved to: {pdf_file}")
    return pdf_file


# --- New function: generate_connection_invoice_pdf_with_custom_discount_and_added_value ---
def generate_connection_invoice_pdf_with_custom_discount_and_added_value(
    customer_name: str,
    invoice_number: str,
    items: list[dict],
    discount: float,
    output_dir: Optional[str] = None,
    explanation_text: Optional[str] = None
):
    """
    Generates a PDF invoice for connection items applying a custom discount (percent if <=100, absolute if >100)
    and then adds 10% added value tax on the net amount.
    Columns: نوع اتصال | محصول | فشار قابل تحمل | سایز | تعداد | قیمت واحد | قیمت کل
    """
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(__file__), "خروجی")
    os.makedirs(output_dir, exist_ok=True)

    from bidi.algorithm import get_display
    from arabic_reshaper import reshape
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from persiantools import digits
    from khayyam import JalaliDate

    date_jalali = fetch_current_jalali_date()
    sh_company = str(get_display(reshape(COMPANY_NAME)))
    sh_date    = str(get_display(reshape(f"تاریخ: {date_jalali}")))
    sh_inv     = str(get_display(reshape(f"شماره پیش‌فاکتور: {invoice_number}")))
    label_cust = str(get_display(reshape("نام مشتری:")))
    sh_customer_name = str(get_display(reshape(customer_name)))

    pdf_file = os.path.join(output_dir, f"{invoice_number}.pdf")
    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
    logo_path = os.path.join(os.path.dirname(__file__), "program files", "logo.png")

    def _draw_logo(canvas, _doc):
        if os.path.exists(logo_path):
            x = _doc.leftMargin
            y = doc.pagesize[1] - doc.topMargin - 60
            canvas.drawImage(logo_path, x, y, width=90, height=60, preserveAspectRatio=True, mask='auto')

    elements = []
    title_style = ParagraphStyle(
        name="CompanyTitle",
        fontName="Persian",
        fontSize=18,
        alignment=TA_CENTER,
        leading=22
    )
    elements.append(Paragraph(sh_company, title_style))
    elements.append(Spacer(1, 12))

    header = [["", sh_date], ["", sh_inv], ["", f"{sh_customer_name}{label_cust}"]]
    table = Table(header, colWidths=[100, 400])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONT', (0, 0), (-1, -1), "Persian", 12),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 20))

    # Updated headers: add "فشار قابل تحمل"
    headers_text = ["نوع اتصال", "محصول", "فشار قابل تحمل", "سایز", "تعداد", "قیمت واحد", "قیمت کل"]
    headers = [str(get_display(reshape(h))) for h in headers_text]
    headers = list(reversed(headers))
    data = [headers]

    total_price_all = 0.0
    for itm in items:
        type_val    = str(get_display(reshape(str(itm.get("type", "")))))
        product_val = str(get_display(reshape(str(itm.get("product", "")))))
        pn_val      = str(get_display(reshape(str(itm.get("pn", "")))))
        size_val    = str(get_display(reshape(str(itm.get("size", "")))))
        quantity = itm.get("quantity", 1)
        quantity_val = str(get_display(reshape(str(int(quantity)))))
        unit_price = itm.get("unit_price", 0)
        total_price = itm.get("total_price", 0)
        total_price_all += total_price
        unit_price_str = str(get_display(reshape(f"{int(unit_price):,}")))
        total_price_str = str(get_display(reshape(f"{int(total_price):,}")))
        row = [
            type_val,
            product_val,
            pn_val,
            size_val,
            quantity_val,
            unit_price_str,
            total_price_str,
        ]
        data.append(list(reversed(row)))

    # Determine custom discount amount
    if discount <= 100:
        discount_amount = total_price_all * (discount / 100.0)
    else:
        discount_amount = discount

    sh_discount_label = str(get_display(reshape("تخفیف")))
    sh_discount_value = str(get_display(reshape(f"{int(discount_amount):,}")))
    discount_row = ["", "", "", "", "", sh_discount_label, sh_discount_value]
    data.append(list(reversed(discount_row)))

    # Net after discount
    net_after_discount = total_price_all - discount_amount

    # Add 10% added value row on net amount
    added_value = net_after_discount * 0.10
    sh_added_label = str(get_display(reshape("مالیات بر ارزش افزوده")))
    sh_added_value = str(get_display(reshape(f"{int(added_value):,}")))
    added_row = ["", "", "", "", "", sh_added_label, sh_added_value]
    data.append(list(reversed(added_row)))

    # Final total including added value
    final_total = net_after_discount + added_value
    sh_total_label = str(get_display(reshape("جمع کل")))
    sh_total_price = str(get_display(reshape(f"{int(final_total):,}")))
    total_row = ["", "", "", "", "", sh_total_label, sh_total_price]
    data.append(list(reversed(total_row)))

    from reportlab.pdfbase.pdfmetrics import stringWidth
    col_font = "Persian"
    col_font_size = 10
    num_cols = len(data[0])
    col_widths = []
    for idx in range(num_cols):
        max_w = max(stringWidth(str(row[idx]), col_font, col_font_size) for row in data)
        col_widths.append(max_w + 20)

    tbl = Table(
        data,
        repeatRows=1,
        colWidths=col_widths
    )
    tbl.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONT', (0, 0), (-1, -1), "Persian", 10),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, -3), (1, -3), colors.lightgrey),
        ('BACKGROUND', (0, -2), (1, -2), colors.lightgrey),
        ('BACKGROUND', (0, -1), (1, -1), colors.lightgrey),
        ('ALIGN', (1, len(data)-1), (1, len(data)-1), 'LEFT'),
    ]))
    elements.append(tbl)

    # Add explanation text if provided
    if explanation_text and explanation_text.strip():
        elements.append(Spacer(1, 24))
        sh_explanation_label = str(get_display(reshape("توضیحات:")))
        sh_explanation = str(get_display(reshape(explanation_text)))
        explanation_label_style = ParagraphStyle(
            name="ExplanationLabel", fontName="Persian", fontSize=10,
            alignment=TA_RIGHT, leading=14, spaceBefore=6
        )
        explanation_text_style = ParagraphStyle(
            name="ExplanationText", fontName="Persian", fontSize=10,
            alignment=TA_RIGHT, leading=14, rightIndent=0
        )
        elements.append(Paragraph(sh_explanation_label, explanation_label_style))
        elements.append(Paragraph(sh_explanation, explanation_text_style))
        elements.append(Spacer(1, 12))

    # Write total in words
    append_total_words(elements, final_total)

    doc.build(elements, onFirstPage=_draw_logo, onLaterPages=_draw_logo)
    print(f"Connection invoice PDF with custom discount and added value saved to: {pdf_file}")
    return pdf_file