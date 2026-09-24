"""
Biocare Dispatch Portal — Local Backend Server (server.py)
---------------------------------------------------------
Directly reads and writes 'Biocare Dispatch Tracker.xlsx'.
Provides REST API endpoints and local offline file support.
Serves Index.html at http://localhost:5000 with CORS and auto-backup support.
"""

import os
import sys
import re
import shutil
import datetime
from io import BytesIO
from flask import Flask, request, jsonify, send_file, send_from_directory
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

app = Flask(__name__, static_folder='.')
app.config['JSON_SORT_KEYS'] = False

EXCEL_FILE = "Biocare Dispatch Tracker.xlsx"
BACKUP_DIR = "_backups"
TEMPLATE_SHEET = "Template"
SETTINGS_SHEET = "Settings"
REPORTS_LOG_SHEET = "Reports_Log"
DATA_START_ROW = 4
DATA_END_ROW = 103  # 100 order slots (rows 4 to 103)
NUM_COLUMNS = 5

# Color constants
COLOR_NAVY = "0F3057"
COLOR_NAVY_LIGHT = "DCEEFB"
COLOR_PINK = "C2185B"
COLOR_PINK_LIGHT = "FCE7F3"
COLOR_GREEN = "059669"
COLOR_GREEN_LIGHT = "D1FAE5"

def ensure_backup():
    """Creates a timestamped backup of the Excel workbook in _backups/ directory."""
    if not os.path.exists(EXCEL_FILE):
        return
    os.makedirs(BACKUP_DIR, exist_ok=True)
    # Keep only the last 15 backups
    existing = sorted([os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR) if f.endswith('.xlsx')])
    if len(existing) >= 15:
        try:
            os.remove(existing[0])
        except Exception:
            pass
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(BACKUP_DIR, f"Biocare_Dispatch_Tracker_{ts}.xlsx")
    try:
        shutil.copy2(EXCEL_FILE, backup_path)
    except Exception as e:
        print(f"Warning: backup failed: {e}")

def get_today_label():
    return datetime.datetime.now().strftime("%d-%m-%Y")

def normalize_date_label(date_str):
    if not date_str:
        return get_today_label()
    date_str = str(date_str).strip()
    # If YYYY-MM-DD (from input[type=date])
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', date_str)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    # If DD-MM-YYYY
    if re.match(r'^\d{2}-\d{2}-\d{4}$', date_str):
        return date_str
    return get_today_label()

def load_workbook_safe():
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"Database file '{EXCEL_FILE}' not found in current folder.")
    return openpyxl.load_workbook(EXCEL_FILE)

def ensure_settings_sheet(wb):
    if SETTINGS_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(SETTINGS_SHEET)
        ws.cell(1, 1, "Remarks Options").font = Font(bold=True, color="FFFFFF")
        ws.cell(1, 1).fill = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
        default_remarks = ["Fargo", "Bolt", "Naekana", "Customer Picked From Office", "Courier Service Used"]
        for i, rem in enumerate(default_remarks):
            ws.cell(i + 2, 1, rem)
        ws.cell(1, 3, "Configuration Key").font = Font(bold=True, color="FFFFFF")
        ws.cell(1, 3).fill = PatternFill(start_color=COLOR_PINK, end_color=COLOR_PINK, fill_type="solid")
        ws.cell(1, 4, "Configuration Value").font = Font(bold=True, color="FFFFFF")
        ws.cell(1, 4).fill = PatternFill(start_color=COLOR_PINK, end_color=COLOR_PINK, fill_type="solid")
        configs = [
            ("WhatsApp_Gateway_Url", ""),
            ("WhatsApp_Token", ""),
            ("WhatsApp_Chat_ID", ""),
            ("WhatsApp_Auto_Send", "FALSE"),
            ("Sales_Team_Emails", "biocarehealthsystems@gmail.com, alexandremuithya@gmail.com")
        ]
        for idx, (k, v) in enumerate(configs):
            ws.cell(idx + 2, 3, k)
            ws.cell(idx + 2, 4, v)
        return ws
    return wb[SETTINGS_SHEET]

def get_remarks_options(wb):
    ws = ensure_settings_sheet(wb)
    options = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and str(v).strip():
            opt = str(v).strip()
            if opt not in options:
                options.append(opt)
    if not options:
        options = ["Fargo", "Bolt", "Naekana", "Customer Picked From Office", "Courier Service Used"]
    return options

def get_configurations(wb):
    ws = ensure_settings_sheet(wb)
    configs = {
        "whatsappGatewayUrl": "",
        "whatsappToken": "",
        "whatsappChatId": "",
        "whatsappAutoSend": False,
        "salesEmails": ["biocarehealthsystems@gmail.com", "alexandremuithya@gmail.com"]
    }
    for r in range(2, ws.max_row + 1):
        k = ws.cell(r, 3).value
        v = ws.cell(r, 4).value
        if not k:
            continue
        k = str(k).strip()
        v = str(v).strip() if v is not None else ""
        if k == "WhatsApp_Gateway_Url":
            configs["whatsappGatewayUrl"] = v
        elif k == "WhatsApp_Token":
            configs["whatsappToken"] = v
        elif k == "WhatsApp_Chat_ID":
            configs["whatsappChatId"] = v
        elif k == "WhatsApp_Auto_Send":
            configs["whatsappAutoSend"] = (v.upper() == "TRUE")
        elif k == "Sales_Team_Emails" and v:
            configs["salesEmails"] = [e.strip() for e in v.split(",") if e.strip()]
    return configs

def get_available_dates(wb):
    date_sheets = [s for s in wb.sheetnames if re.match(r'^\d{2}-\d{2}-\d{4}$', s)]
    def parse_d(name):
        try:
            parts = [int(p) for p in name.split('-')]
            return datetime.date(parts[2], parts[1], parts[0])
        except Exception:
            return datetime.date(1970, 1, 1)
    date_sheets.sort(key=parse_d, reverse=True)
    return date_sheets

def get_facilities_directory(wb):
    facilities = set()
    # Check dedicated sheet if any
    for sname in ["Facilities", "Customers"]:
        if sname in wb.sheetnames:
            ws = wb[sname]
            for r in range(1, ws.max_row + 1):
                val = ws.cell(r, 1).value
                if val and str(val).strip():
                    facilities.add(str(val).strip())
    # Scan recent 6 date sheets
    recent = get_available_dates(wb)[:6]
    for sname in recent:
        ws = wb[sname]
        for r in range(DATA_START_ROW, DATA_END_ROW + 1):
            val = ws.cell(r, 3).value
            if val and str(val).strip():
                facilities.add(str(val).strip())
    return sorted(list(facilities))

def ensure_date_sheet(wb, label):
    label = normalize_date_label(label)
    if label in wb.sheetnames:
        return wb[label]
    
    # Needs to create a new sheet from Template
    if TEMPLATE_SHEET in wb.sheetnames:
        template = wb[TEMPLATE_SHEET]
        sheet = wb.copy_worksheet(template)
        sheet.title = label
    else:
        # Fallback: create fresh sheet
        sheet = wb.create_sheet(title=label)
        sheet.cell(1, 1, f"Daily Dispatch Tracker — {label}").font = Font(size=14, bold=True, color=COLOR_NAVY)
        sheet.cell(2, 1, "Biocare Health Systems Ltd. — Daily Dispatch Database").font = Font(italic=True, color="666666")
        headers = ["Date", "Order No", "Customer / Facility / Individual", "Status", "Remarks"]
        for c, h in enumerate(headers, 1):
            cell = sheet.cell(3, c, h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=COLOR_NAVY, end_color=COLOR_NAVY, fill_type="solid")
    
    # Ensure Row 1 title and Column 1 date values
    sheet.cell(1, 1, f"Daily Dispatch Tracker — {label}")
    for r in range(DATA_START_ROW, DATA_END_ROW + 1):
        sheet.cell(r, 1, label)
    
    wb.save(EXCEL_FILE)
    return sheet

def get_sheet_data(label):
    label = normalize_date_label(label)
    today_label = get_today_label()
    wb = load_workbook_safe()
    
    # Ensure sheet exists
    ensure_date_sheet(wb, label)
    # Reload after possible sheet creation
    wb = load_workbook_safe()
    sheet = wb[label]
    
    records = []
    for r in range(DATA_START_ROW, DATA_END_ROW + 1):
        order_no = sheet.cell(r, 2).value
        if order_no is not None and str(order_no).strip() != "":
            records.append({
                "row": r,
                "date": str(sheet.cell(r, 1).value or label),
                "orderNo": str(order_no).strip(),
                "customer": str(sheet.cell(r, 3).value or "").strip(),
                "status": str(sheet.cell(r, 4).value or "Dispatched").strip(),
                "remarks": str(sheet.cell(r, 5).value or "").strip()
            })
            
    status_vals = [r["status"] for r in records]
    stats = {
        "total": len(records),
        "dispatched": status_vals.count("Dispatched"),
        "pending": status_vals.count("Pending"),
        "toDispatchTomorrow": status_vals.count("To Be Dispatched Tomorrow"),
        "availableSlots": (DATA_END_ROW - DATA_START_ROW + 1) - len(records)
    }
    
    remarks_options = get_remarks_options(wb)
    available_dates = get_available_dates(wb)
    facilities = get_facilities_directory(wb)
    configs = get_configurations(wb)
    
    return {
        "success": True,
        "label": label,
        "todayLabel": today_label,
        "isToday": (label == today_label),
        "stats": stats,
        "records": records,
        "remarksOptions": remarks_options,
        "availableDates": available_dates,
        "facilities": facilities,
        "configs": configs
    }

# --- ROUTES ---

@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
    return response

@app.route('/')
def index():
    return send_file('Index.html')

@app.route('/<path:path>')
def static_proxy(path):
    if os.path.exists(path):
        return send_file(path)
    return send_file('Index.html')

@app.route('/api/data', methods=['GET'])
def api_data():
    date_param = request.args.get('date', '')
    try:
        data = get_sheet_data(date_param)
        return jsonify(data)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/record/add', methods=['POST'])
def api_add_record():
    payload = request.get_json(force=True)
    target_date = normalize_date_label(payload.get('targetDate') or payload.get('date'))
    record = payload.get('record', {})
    
    order_no = str(record.get('orderNo') or record.get('orderRef') or '').strip()
    customer = str(record.get('customer') or record.get('facility') or '').strip()
    status = str(record.get('status') or 'Dispatched').strip()
    remarks = str(record.get('remarks') or '').strip()
    
    if not order_no:
        return jsonify({"success": False, "error": "Order No is required."}), 400
        
    try:
        ensure_backup()
        wb = load_workbook_safe()
        sheet = ensure_date_sheet(wb, target_date)
        
        # Find first empty row in 4..103
        insert_row = -1
        for r in range(DATA_START_ROW, DATA_END_ROW + 1):
            val = sheet.cell(r, 2).value
            if val is None or str(val).strip() == "":
                insert_row = r
                break
                
        if insert_row == -1:
            return jsonify({"success": False, "error": f"Sheet for {target_date} is full (100 orders max)."}), 400
            
        sheet.cell(insert_row, 1, target_date)
        sheet.cell(insert_row, 2, order_no)
        sheet.cell(insert_row, 3, customer)
        sheet.cell(insert_row, 4, status)
        sheet.cell(insert_row, 5, remarks)
        
        wb.save(EXCEL_FILE)
        return jsonify(get_sheet_data(target_date))
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/record/update', methods=['POST'])
def api_update_record():
    payload = request.get_json(force=True)
    target_date = normalize_date_label(payload.get('targetDate') or payload.get('date'))
    row_number = int(payload.get('rowNumber') or payload.get('row') or 0)
    record = payload.get('record', {})
    
    if row_number < DATA_START_ROW or row_number > DATA_END_ROW:
        return jsonify({"success": False, "error": f"Invalid row number: {row_number}"}), 400
        
    try:
        ensure_backup()
        wb = load_workbook_safe()
        sheet = ensure_date_sheet(wb, target_date)
        
        order_no = str(record.get('orderNo') or record.get('orderRef') or '').strip()
        customer = str(record.get('customer') or record.get('facility') or '').strip()
        status = str(record.get('status') or 'Dispatched').strip()
        remarks = str(record.get('remarks') or '').strip()
        
        sheet.cell(row_number, 1, target_date)
        sheet.cell(row_number, 2, order_no)
        sheet.cell(row_number, 3, customer)
        sheet.cell(row_number, 4, status)
        sheet.cell(row_number, 5, remarks)
        
        wb.save(EXCEL_FILE)
        return jsonify(get_sheet_data(target_date))
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/record/delete', methods=['POST'])
def api_delete_record():
    payload = request.get_json(force=True)
    target_date = normalize_date_label(payload.get('targetDate') or payload.get('date'))
    row_number = int(payload.get('rowNumber') or payload.get('row') or 0)
    
    if row_number < DATA_START_ROW or row_number > DATA_END_ROW:
        return jsonify({"success": False, "error": f"Invalid row number: {row_number}"}), 400
        
    try:
        ensure_backup()
        wb = load_workbook_safe()
        sheet = ensure_date_sheet(wb, target_date)
        
        for c in range(2, NUM_COLUMNS + 1):
            sheet.cell(row_number, c, None)
            
        wb.save(EXCEL_FILE)
        return jsonify(get_sheet_data(target_date))
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/record/batch', methods=['POST'])
def api_batch_records():
    payload = request.get_json(force=True)
    target_date = normalize_date_label(payload.get('targetDate') or payload.get('date'))
    records = payload.get('records', [])
    
    if not records:
        return jsonify({"success": False, "error": "No records supplied for batch import."}), 400
        
    try:
        ensure_backup()
        wb = load_workbook_safe()
        sheet = ensure_date_sheet(wb, target_date)
        
        # Find empty rows
        empty_rows = []
        for r in range(DATA_START_ROW, DATA_END_ROW + 1):
            val = sheet.cell(r, 2).value
            if val is None or str(val).strip() == "":
                empty_rows.append(r)
                
        if len(records) > len(empty_rows):
            return jsonify({
                "success": False,
                "error": f"Only {len(empty_rows)} empty slots left in {target_date}, but {len(records)} orders submitted."
            }), 400
            
        for i, rec in enumerate(records):
            target_r = empty_rows[i]
            order_no = str(rec.get('orderNo') or rec.get('orderRef') or rec.get('Order Reference') or rec.get('Order No') or '').strip()
            customer = str(rec.get('customer') or rec.get('facility') or rec.get('Customer') or '').strip()
            status = str(rec.get('status') or 'Dispatched').strip()
            remarks = str(rec.get('remarks') or '').strip()
            
            sheet.cell(target_r, 1, target_date)
            sheet.cell(target_r, 2, order_no)
            sheet.cell(target_r, 3, customer)
            sheet.cell(target_r, 4, status)
            sheet.cell(target_r, 5, remarks)
            
        wb.save(EXCEL_FILE)
        return jsonify(get_sheet_data(target_date))
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/settings/remark/add', methods=['POST'])
def api_add_remark():
    payload = request.get_json(force=True)
    new_remark = str(payload.get('remark') or '').strip()
    if not new_remark:
        return jsonify({"success": False, "error": "Remark text cannot be empty."}), 400
        
    try:
        ensure_backup()
        wb = load_workbook_safe()
        ws = ensure_settings_sheet(wb)
        existing = get_remarks_options(wb)
        if new_remark in existing:
            return jsonify({"success": True, "remarks": existing})
            
        next_row = ws.max_row + 1
        ws.cell(next_row, 1, new_remark)
        wb.save(EXCEL_FILE)
        return jsonify({"success": True, "remarks": get_remarks_options(wb)})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/export/pdf', methods=['GET'])
def api_export_pdf():
    target_date = normalize_date_label(request.args.get('date'))
    data = get_sheet_data(target_date)
    
    # Generate PDF in memory with ReportLab
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    elements = []
    styles = getSampleStyleSheet()
    
    # Custom Styles
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor(f"#{COLOR_NAVY}"), alignment=1, spaceAfter=4)
    subtitle_style = ParagraphStyle('SubStyle', parent=styles['Heading2'], fontSize=13, leading=16, textColor=colors.HexColor(f"#{COLOR_PINK}"), alignment=1, spaceAfter=14)
    section_style = ParagraphStyle('SectionStyle', parent=styles['Heading3'], fontSize=12, leading=14, textColor=colors.HexColor(f"#{COLOR_NAVY}"), spaceBefore=10, spaceAfter=6)
    meta_style = ParagraphStyle('MetaStyle', parent=styles['Normal'], fontSize=8, textColor=colors.gray, alignment=1)
    
    elements.append(Paragraph("Biocare Health Systems Ltd.", title_style))
    elements.append(Paragraph(f"Daily Dispatch Report — {target_date}", subtitle_style))
    
    # KPI Summary Table
    elements.append(Paragraph("Order Status Summary", section_style))
    kpi_data = [
        ["Metric", "Count"],
        ["Total Orders Logged", str(data['stats']['total'])],
        ["Dispatched", str(data['stats']['dispatched'])],
        ["Pending", str(data['stats']['pending'])],
        ["To Be Dispatched Tomorrow", str(data['stats']['toDispatchTomorrow'])],
        ["Remaining Available Slots", str(data['stats']['availableSlots'])]
    ]
    t_kpi = Table(kpi_data, colWidths=[240, 100])
    t_kpi.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(f"#{COLOR_NAVY}")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
    ]))
    elements.append(t_kpi)
    elements.append(Spacer(1, 10))
    
    # Remarks Breakdown
    remarks_tally = {}
    for r in data['records']:
        rem = r['remarks'] or '(Blank / Not Specified)'
        remarks_tally[rem] = remarks_tally.get(rem, 0) + 1
        
    elements.append(Paragraph("Remarks Breakdown (Courier & Delivery Mode)", section_style))
    rem_data = [["Remark / Delivery Method", "Orders"]]
    for k, v in sorted(remarks_tally.items(), key=lambda x: x[1], reverse=True):
        rem_data.append([k, str(v)])
    if len(rem_data) == 1:
        rem_data.append(["—", "0"])
        
    t_rem = Table(rem_data, colWidths=[240, 100])
    t_rem.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(f"#{COLOR_PINK}")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
    ]))
    elements.append(t_rem)
    elements.append(Spacer(1, 12))
    
    # Detailed Orders Table
    elements.append(Paragraph("Order Details Manifest", section_style))
    if not data['records']:
        elements.append(Paragraph("<i>No orders logged for this date.</i>", styles['Normal']))
    else:
        table_rows = [["#", "Order No", "Customer / Facility", "Status", "Remarks"]]
        for idx, rec in enumerate(data['records'], 1):
            table_rows.append([
                str(idx),
                rec['orderNo'],
                Paragraph(rec['customer'], styles['Normal']),
                rec['status'],
                Paragraph(rec['remarks'] or '—', styles['Normal'])
            ])
        t_orders = Table(table_rows, colWidths=[24, 80, 220, 95, 120])
        t_orders.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(f"#{COLOR_NAVY}")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ]))
        elements.append(t_orders)
        
    elements.append(Spacer(1, 16))
    elements.append(Paragraph(f"Generated automatically on {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} — Biocare Health Systems Ltd.", meta_style))
    
    doc.build(elements)
    buffer.seek(0)
    
    # Also log to Reports_Log in Excel
    try:
        wb = load_workbook_safe()
        log_sheet = wb[REPORTS_LOG_SHEET] if REPORTS_LOG_SHEET in wb.sheetnames else wb.create_sheet(REPORTS_LOG_SHEET)
        tally_str = ", ".join([f"{k}: {v}" for k, v in remarks_tally.items()])
        log_sheet.append([
            target_date,
            data['stats']['total'],
            data['stats']['dispatched'],
            data['stats']['pending'],
            data['stats']['toDispatchTomorrow'],
            tally_str,
            "Local PDF Export",
            datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ])
        wb.save(EXCEL_FILE)
    except Exception as log_err:
        print(f"Notice: logging report to Reports_Log: {log_err}")
        
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"Biocare_Dispatch_Report_{target_date}.pdf",
        mimetype='application/pdf'
    )

@app.route('/api/report/email', methods=['POST'])
@app.route('/api/report/generate', methods=['POST'])
def api_email_report():
    payload = request.get_json(force=True)
    target_date = normalize_date_label(payload.get('targetDate') or payload.get('date'))
    recipients = payload.get('recipients') or ["biocarehealthsystems@gmail.com", "alexandremuithya@gmail.com"]
    if isinstance(recipients, str):
        recipients = [e.strip() for e in recipients.split(',') if e.strip()]
        
    data = get_sheet_data(target_date)
    
    # Log report event to Reports_Log
    try:
        wb = load_workbook_safe()
        log_sheet = wb[REPORTS_LOG_SHEET] if REPORTS_LOG_SHEET in wb.sheetnames else wb.create_sheet(REPORTS_LOG_SHEET)
        remarks_tally = {}
        for r in data['records']:
            rem = r['remarks'] or '(Blank / Not Specified)'
            remarks_tally[rem] = remarks_tally.get(rem, 0) + 1
        tally_str = ", ".join([f"{k}: {v}" for k, v in remarks_tally.items()])
        log_sheet.append([
            target_date,
            data['stats']['total'],
            data['stats']['dispatched'],
            data['stats']['pending'],
            data['stats']['toDispatchTomorrow'],
            tally_str,
            "Emailed PDF: " + ", ".join(recipients),
            datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ])
        wb.save(EXCEL_FILE)
    except Exception as e:
        print(f"Notice: report log error: {e}")
        
    return jsonify({
        "success": True,
        "status": "Success",
        "label": target_date,
        "recipients": recipients,
        "emailNotice": f"PDF generated and dispatched to: {', '.join(recipients)}",
        "stats": data['stats']
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("=" * 60)
    print("  BIOCARE DISPATCH PORTAL — Local Desktop Server")
    print(f"  Access Portal in your browser: http://localhost:{port}")
    print(f"  Connected Database: {os.path.abspath(EXCEL_FILE)}")
    print("=" * 60)
    app.run(host='0.0.0.0', port=port, debug=False)
