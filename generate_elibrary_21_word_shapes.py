"""Generate an e-Library System interface document using editable Word shapes only."""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

import generate_22_interfaces_doc as base

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "e_Library_System_21_Editable_Word_Interfaces.docx"
FONT = "Calibri"

# Exactly 21 interfaces derived from the supplied project report, Context Diagram,
# DFD and FDD. Tuple: title, role, section, inputs, action, process, popup, output, kind.
SCREENS = [
    ("Secure Login", "Member / Librarian", "Access", ["User ID", "Password", "Role"], "SIGN IN",
     "Authenticate the account, role and active membership status.",
     "Login successful. Welcome to the e-Library.",
     "The correct dashboard opens according to the account role.", "auth"),
    ("Library Dashboard", "Librarian", "Overview", ["Date range", "Branch"], "REFRESH",
     "Retrieve live member, loan, inventory and fine totals.",
     "Dashboard data has been refreshed.",
     "Summary cards show active members, books on loan, overdue books and collected fines.", "dashboard"),
    ("Register New Member", "Librarian", "Member Management", ["Full name", "Matrix / Member ID", "Email", "Phone"], "REGISTER MEMBER",
     "Validate the details, create a member ID and store the member record.",
     "New member registered successfully.",
     "A member profile and registration confirmation are displayed.", "form"),
    ("Update Member Profile", "Member / Librarian", "Member Management", ["Member ID", "Email", "Phone", "Address"], "SAVE PROFILE",
     "Find the member record and save only the changed profile fields.",
     "Member profile updated successfully.",
     "The latest contact details and account status are displayed.", "form"),
    ("Delete Membership", "Librarian", "Member Management", ["Member ID", "Reason", "Confirm checkbox"], "DELETE MEMBERSHIP",
     "Check for active loans or unpaid fines before deactivating the membership.",
     "Delete this membership permanently?", 
     "The account is deactivated and an audit entry is created.", "confirm"),
    ("Search Book Catalogue", "Member", "Book Management", ["Title / ISBN / author", "Category", "Format"], "SEARCH BOOKS",
     "Search the catalogue index and apply the selected filters.",
     "12 matching books were found.",
     "Book cards show title, author, call number, format and availability.", "catalog"),
    ("Book Details & Availability", "Member", "Book Management", ["Select copy", "Select branch"], "REQUEST BOOK",
     "Read the selected book record and check real-time copy availability.",
     "Book request placed successfully.",
     "Full metadata, shelf location and available-copy status are displayed.", "detail"),
    ("Add New Book", "Librarian", "Book Management", ["ISBN", "Title", "Author", "Category"], "ADD BOOK",
     "Check for duplicate ISBN, create the book record and add its copies to inventory.",
     "New book and copy records added successfully.",
     "The catalogue displays the new title with its generated call number.", "form"),
    ("Update Book Information", "Librarian", "Book Management", ["Book ID / ISBN", "Title", "Shelf", "Availability"], "UPDATE BOOK",
     "Locate the book record, validate changes and update catalogue and inventory data.",
     "Book information updated successfully.",
     "The revised title, shelf, copy count and status are displayed.", "form"),
    ("Book Purchase Order", "Librarian", "Supplier Management", ["Supplier", "ISBN / Title", "Quantity", "Unit price"], "SEND ORDER",
     "Calculate the total, create a purchase order and send it to the supplier.",
     "Purchase order sent to the selected supplier.",
     "An order reference, ordered titles, quantity and delivery status are displayed.", "transaction"),
    ("Issue / Borrow Book", "Librarian", "Circulation", ["Member ID", "Book barcode", "Loan period"], "ISSUE BOOK",
     "Verify the member, check availability and create the loan record and due date.",
     "Book issued successfully. Due date: 28 July 2026.",
     "A borrowing receipt shows member, book, issue date and due date.", "transaction"),
    ("Current Loans", "Member", "Circulation", ["Loan status", "Sort by due date"], "VIEW LOAN",
     "Retrieve all active and previous loan records for the signed-in member.",
     "Loan record opened.",
     "The loan list shows book title, issue date, due date and current status.", "list"),
    ("Renew Book Loan", "Member / Librarian", "Circulation", ["Loan ID", "New loan period"], "RENEW LOAN",
     "Check renewal limit, reservations and fines before extending the due date.",
     "Loan renewed. The new due date is 11 August 2026.",
     "The loan record and renewal count are updated with the new due date.", "transaction"),
    ("Return Book", "Librarian", "Circulation", ["Book barcode", "Return condition", "Return date"], "RETURN BOOK",
     "Close the loan, update book availability and calculate any overdue charge.",
     "Return recorded. An overdue fine of RM 3.00 applies.",
     "A return receipt shows loan closure, condition, days late and fine amount.", "transaction"),
    ("Fine Calculation", "Librarian", "Fine Management", ["Loan ID", "Days overdue", "Adjustment reason"], "CALCULATE FINE",
     "Apply the daily overdue rate and any approved fine adjustment.",
     "Fine calculated: RM 3.00.",
     "Fine breakdown shows overdue days, rate, adjustment and final amount.", "calculator"),
    ("Record Fine Payment", "Librarian", "Fine Management", ["Member ID", "Fine reference", "Payment method", "Amount"], "RECORD PAYMENT",
     "Validate the amount, store the payment record and update the fine balance.",
     "Payment received. Receipt FP-2026-0182 created.",
     "A payment receipt shows amount, method, balance and transaction time.", "transaction"),
    ("Overdue & Fine Notification", "Member / Librarian", "Notifications", ["Member ID", "Channel", "Template"], "SEND NOTICE",
     "Compile overdue loan details and send an email or SMS notification.",
     "Overdue and fine notice sent successfully.",
     "Delivery status, recipient, overdue title, due date and fine are displayed.", "notification"),
    ("Overdue Books Report", "Librarian", "Report Generation", ["Start date", "End date", "Branch"], "GENERATE REPORT",
     "Aggregate open overdue loans and calculate current outstanding fines.",
     "Overdue books report generated successfully.",
     "A report table lists member, book, due date, days overdue and fine.", "report"),
    ("Member Activity Report", "Librarian", "Report Generation", ["Member ID / All", "Start date", "End date"], "GENERATE REPORT",
     "Summarize registrations, searches, borrowing, renewals, returns and payments.",
     "Member activity report generated successfully.",
     "Activity totals and a transaction history table are displayed.", "report"),
    ("Borrowing & Inventory Report", "Librarian", "Report Generation", ["Period", "Category", "Format"], "GENERATE REPORT",
     "Combine loan trends with available, borrowed, lost and damaged copy counts.",
     "Borrowing and inventory report is ready.",
     "Charts and tables show popular titles, circulation totals and stock status.", "report"),
    ("System Logs & Error Alerts", "Administrator", "Security & Monitoring", ["Severity", "Module", "Date range"], "FILTER LOGS",
     "Read audit logs and system errors, then apply severity and module filters.",
     "Critical error alert acknowledged.",
     "The log table shows timestamp, user, action, severity and error details.", "logs"),
]

TEAM = [
    ("Muhammad Haziq Bin Jalaludin", "13DIT24F2029", "Interface & Circulation"),
    ("Muhammad Danish Bin Osman", "13DIT24F2030", "Member & Book Management"),
    ("Seri Maisarah Binti Helmy Rozario", "13DIT24F1083", "Fine, Reports & Documentation"),
]

# Aliases to the native DrawingML primitives in the established shape engine.
box = base.box
label = base.label
circle = base.circle
rule = base.rule
marker = base.marker


def ptext(doc, text, size=11, bold=False, align="left", before=0, after=4, color=(34, 34, 34)):
    p = doc.add_paragraph()
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "right": WD_ALIGN_PARAGRAPH.RIGHT}[align]
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = RGBColor(*color)
    return p


def set_cell(cell, text, bold=False, fill=None, size=9):
    cell.text = text
    for run in cell.paragraphs[0].runs:
        run.font.name = FONT
        run.font.size = Pt(size)
        run.font.bold = bold
    if fill:
        base.shading(cell, fill)
    base.cell_margins(cell, 55, 75, 55, 75)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def cover(doc):
    ptext(doc, "DEPARTMENT OF INFORMATION AND COMMUNICATION TECHNOLOGY", 11, True, "center", 28, 6)
    ptext(doc, "DFC40343: SYSTEM ANALYSIS AND DESIGN FUNDAMENTALS", 11, True, "center", 0, 20)
    ptext(doc, "e-LIBRARY SYSTEM", 23, True, "center", 0, 4)
    ptext(doc, "21 Editable Interface Sketches", 15, False, "center", 0, 4)
    ptext(doc, "Input  |  Process  |  Popup Message  |  Output", 10, False, "center", 0, 18, (90, 90, 90))

    # Redesigned details: separate project card followed by a three-column team directory.
    details = doc.add_table(rows=4, cols=2)
    details.style = "Table Grid"
    details.alignment = WD_TABLE_ALIGNMENT.CENTER
    values = [("Project", "e-Library System"), ("Lecturer", "Pn Murniyati Binti Abdul"),
              ("Class", "DIT4-S1"), ("Document", "21 Interface Designs")]
    for row, (key, value) in zip(details.rows, values):
        set_cell(row.cells[0], key, True, "E7E7E7", 9.5)
        set_cell(row.cells[1], value, False, None, 9.5)

    ptext(doc, "TEAM DIRECTORY", 12, True, "left", 14, 5)
    team = doc.add_table(rows=1, cols=4)
    team.style = "Table Grid"
    team.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, text in enumerate(["No.", "Student name", "Matrix number", "Assigned focus"]):
        set_cell(team.rows[0].cells[idx], text, True, "D9D9D9", 8.5)
    for number, member in enumerate(TEAM, 1):
        cells = team.add_row().cells
        for idx, text in enumerate((str(number),) + member):
            set_cell(cells[idx], text, False, "F8F8F8" if number % 2 == 0 else None, 8.2)

    ptext(doc, "All interface sketches in this document are built from native Microsoft Word shapes. "
               "They are not PNG images and every rectangle, line, label, field, button and popup can be edited.",
          9, False, "center", 14, 0, (85, 85, 85))
    doc.add_page_break()


def index_page(doc):
    ptext(doc, "21 INTERFACE INDEX", 17, True, "left", 0, 2)
    ptext(doc, "Functional coverage mapped from the supplied Context Diagram, DFD and FDD.", 9.5, False,
          "left", 0, 8, (90, 90, 90))
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, text in enumerate(["ID", "Interface", "Primary user", "Functional area"]):
        set_cell(table.rows[0].cells[idx], text, True, "D9D9D9", 8.5)
    for number, s in enumerate(SCREENS, 1):
        cells = table.add_row().cells
        for idx, text in enumerate((f"EL-{number:02d}", s[0], s[1], s[2])):
            set_cell(cells[idx], text, False, "F8F8F8" if number % 2 == 0 else None, 8.2)
    doc.add_page_break()


def guide_page(doc):
    ptext(doc, "INTERFACE READING GUIDE", 17, True, "left", 0, 3)
    ptext(doc, "Every interface page contains one complete e-Library screen plus four numbered workflow notes.",
          9.5, False, "left", 0, 10, (90, 90, 90))
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    for idx, text in enumerate(["Marker", "Stage", "Meaning"]):
        set_cell(table.rows[0].cells[idx], text, True, "D9D9D9", 9)
    rows = [("1", "INPUT", "The member, librarian or administrator enters or selects data."),
            ("2", "PROCESS", "The e-Library validates, searches, calculates or saves the data."),
            ("3", "POPUP", "The system displays a success, warning or confirmation message."),
            ("4", "OUTPUT", "The requested record, receipt, report or status is displayed.")]
    for values in rows:
        cells = table.add_row().cells
        for idx, text in enumerate(values):
            set_cell(cells[idx], text, idx == 1, None, 9)
    ptext(doc, "Design style", 11, True, "left", 16, 4)
    ptext(doc, "A full e-Library application window is used on each page. The left navigation identifies the "
               "functional modules. Input controls occupy the upper workspace, the process and output occupy "
               "the lower workspace, and a popup dialog appears on the right. The notes below the window explain "
               "the exact data flow.", 9.5, False, "left", 0, 0, (80, 80, 80))
    doc.add_page_break()


# A4 page coordinates in points. All shapes remain inside 595 x 842.
PAGE_X, PAGE_Y, PAGE_W, PAGE_H = 22, 52, 551, 456
SIDE_W = 108
MAIN_X = PAGE_X + SIDE_W
MAIN_W = PAGE_W - SIDE_W


def chrome(p, number, s):
    title, role, section = s[0], s[1], s[2]
    box(p, 22, 18, 551, 26, "", "E3E3E3", "444444")
    box(p, 29, 23, 52, 16, f"EL-{number:02d}", "FFFFFF", "444444", 7.5, True, rounded=True)
    label(p, 89, 22, 320, 18, title, 11.5, True, "left")
    label(p, 420, 23, 145, 16, role, 7.5, False, "right", "555555")

    box(p, PAGE_X, PAGE_Y, PAGE_W, PAGE_H, "", "FFFFFF", "333333", line_width=1.3)
    box(p, PAGE_X, PAGE_Y, SIDE_W, PAGE_H, "", "EFEFEF", "555555")
    label(p, PAGE_X + 10, PAGE_Y + 12, SIDE_W - 20, 20, "e-LIBRARY", 12, True, "center")
    label(p, PAGE_X + 10, PAGE_Y + 34, SIDE_W - 20, 14, "MANAGEMENT SYSTEM", 5.8, False, "center", "666666")
    rule(p, PAGE_X + 9, PAGE_Y + 54, SIDE_W - 18, color="999999")
    nav = ["Dashboard", "Members", "Catalogue", "Circulation", "Fines", "Reports", "System Logs"]
    active_map = {"Overview": 0, "Member Management": 1, "Book Management": 2,
                  "Supplier Management": 2, "Circulation": 3, "Fine Management": 4,
                  "Notifications": 4, "Report Generation": 5, "Security & Monitoring": 6,
                  "Access": 0}
    active = active_map.get(section, 0)
    for idx, item in enumerate(nav):
        y = PAGE_Y + 70 + idx * 33
        box(p, PAGE_X + 8, y, SIDE_W - 16, 25, item,
            "D8D8D8" if idx == active else "F7F7F7", "555555" if idx == active else "AAAAAA",
            7.2, idx == active, "left", rounded=True)
    box(p, PAGE_X + 9, PAGE_Y + PAGE_H - 59, SIDE_W - 18, 43, "", "FFFFFF", "999999")
    circle(p, PAGE_X + 15, PAGE_Y + PAGE_H - 49, 24, "U", "EFEFEF", "777777", 8, True)
    label(p, PAGE_X + 44, PAGE_Y + PAGE_H - 51, 53, 16, role.split(" /")[0], 6.5, True, "left")
    label(p, PAGE_X + 44, PAGE_Y + PAGE_H - 36, 53, 12, "Sign out", 6.3, False, "left", "666666")

    box(p, MAIN_X, PAGE_Y, MAIN_W, 36, "", "F8F8F8", "777777")
    label(p, MAIN_X + 12, PAGE_Y + 7, 260, 20, title, 10.5, True, "left")
    box(p, MAIN_X + MAIN_W - 156, PAGE_Y + 8, 110, 20, "Search library", "FFFFFF", "999999",
        6.8, False, "left", rounded=True, color="777777")
    circle(p, MAIN_X + MAIN_W - 35, PAGE_Y + 7, 22, "?", "FFFFFF", "777777", 8, True)


def input_area(p, s):
    inputs, action = s[3], s[4]
    label(p, MAIN_X + 14, 96, 200, 16, "INPUT CONTROLS", 7.5, True, "left", "555555")
    rule(p, MAIN_X + 14, 113, 405, color="AAAAAA")
    for idx, field in enumerate(inputs[:4]):
        col, row = idx % 2, idx // 2
        x = MAIN_X + 14 + col * 205
        y = 121 + row * 46
        label(p, x, y, 185, 13, field, 7.1, False, "left", "333333")
        if "checkbox" in field.lower():
            box(p, x, y + 17, 14, 14, "", "FFFFFF", "666666")
            label(p, x + 20, y + 16, 165, 16, "Selected", 7, False, "left")
        elif any(word in field.lower() for word in ("role", "status", "branch", "category", "format", "method", "channel", "severity", "module", "condition", "period")):
            box(p, x, y + 15, 185, 24, "Select option  v", "FFFFFF", "888888", 7, False, "left")
        else:
            box(p, x, y + 15, 185, 24, "Enter value", "FFFFFF", "888888", 7, False, "left", color="888888")
    box(p, MAIN_X + 14, 218, 142, 27, action, "DDDDDD", "333333", 7.7, True, rounded=True)
    box(p, MAIN_X + 164, 218, 78, 27, "CLEAR", "FFFFFF", "777777", 7.5, False, rounded=True)
    marker(p, MAIN_X + 396, 216, 1)


def dashboard_content(p):
    cards = [("Active members", "1,248"), ("Books on loan", "386"),
             ("Overdue books", "27"), ("Fines collected", "RM 820")]
    for idx, (name, value) in enumerate(cards):
        x = MAIN_X + 14 + idx * 101
        box(p, x, 121, 94, 62, "", "FFFFFF", "888888", rounded=True)
        label(p, x + 7, 127, 80, 13, name, 6.4, False, "left", "666666")
        label(p, x + 7, 145, 80, 28, value, 12, True, "left")
    label(p, MAIN_X + 14, 194, 200, 15, "Circulation this week", 7.5, True, "left")
    for idx, height in enumerate([28, 52, 42, 67, 59, 34, 48]):
        box(p, MAIN_X + 28 + idx * 30, 280 - height, 18, height, "", "E0E0E0", "777777")
        label(p, MAIN_X + 24 + idx * 30, 284, 26, 12, ["M", "T", "W", "T", "F", "S", "S"][idx],
              6.2, False, "center", "666666")
    box(p, MAIN_X + 260, 194, 158, 103, "", "F8F8F8", "999999")
    label(p, MAIN_X + 269, 201, 140, 15, "Attention required", 7.5, True, "left")
    for idx, text in enumerate(["27 overdue loans", "8 unpaid fines", "3 low-stock titles"]):
        box(p, MAIN_X + 269, 224 + idx * 23, 140, 19, text, "FFFFFF", "AAAAAA", 6.7, False, "left")
    marker(p, MAIN_X + 396, 192, 1)


def list_content(p, kind):
    label(p, MAIN_X + 14, 121, 200, 15, "FILTER AND RECORD LIST", 7.5, True, "left")
    for idx, text in enumerate(["Status  v", "Date  v", "Sort by  v"]):
        box(p, MAIN_X + 14 + idx * 116, 143, 108, 22, text, "FFFFFF", "888888", 6.8, False, "left")
    box(p, MAIN_X + 365, 143, 54, 22, "APPLY", "DDDDDD", "444444", 6.8, True, rounded=True)
    headers = ["Record / Title", "Reference", "Date", "Status"]
    widths = [165, 90, 80, 75]
    x = MAIN_X + 14
    for h, w in zip(headers, widths):
        box(p, x, 177, w, 22, h, "E3E3E3", "555555", 6.8, True, "left")
        x += w
    rows = {"logs": ["Login accepted", "Book record updated", "Payment posted"],
            "report": ["The Silent Patient", "Database Systems", "Atomic Habits"],
            "catalog": ["Digital resource", "Printed book", "Institutional journal"],
            "list": ["Current loan 1", "Current loan 2", "Previous loan 3"]}.get(kind, ["Record 1", "Record 2", "Record 3"])
    for row, name in enumerate(rows):
        y = 199 + row * 27
        x = MAIN_X + 14
        values = [name, f"REF-{row+1:03d}", "19 Jul 2026", "Active" if row < 2 else "Closed"]
        for value, w in zip(values, widths):
            box(p, x, y, w, 27, value, "FFFFFF", "BBBBBB", 6.5, False, "left")
            x += w
    marker(p, MAIN_X + 396, 141, 1)


def output_and_popup(p, s):
    process, popup_text, output = s[5], s[6], s[7]
    # Processing card at lower left.
    box(p, MAIN_X + 12, 319, 197, 126, "", "F8F8F8", "666666", dashed=True)
    box(p, MAIN_X + 22, 327, 82, 19, "PROCESS", "E2E2E2", "555555", 7, True, rounded=True)
    circle(p, MAIN_X + 22, 356, 24, "1", "FFFFFF", "666666", 7, True)
    label(p, MAIN_X + 52, 357, 143, 18, "Receive submitted data", 6.6, False, "left")
    rule(p, MAIN_X + 34, 380, 1, 12, color="777777")
    circle(p, MAIN_X + 22, 395, 24, "2", "FFFFFF", "666666", 7, True)
    label(p, MAIN_X + 52, 396, 143, 18, "Validate and process", 6.6, False, "left")
    label(p, MAIN_X + 22, 423, 173, 17, process, 6.2, False, "left", "555555")
    marker(p, MAIN_X + 194, 311, 2)

    # Output summary at lower centre.
    box(p, MAIN_X + 219, 319, 196, 126, "", "FFFFFF", "666666")
    box(p, MAIN_X + 229, 327, 78, 19, "OUTPUT", "E2E2E2", "555555", 7, True, rounded=True)
    label(p, MAIN_X + 229, 353, 176, 52, output, 6.7, False, "left")
    for idx in range(3):
        rule(p, MAIN_X + 229, 412 + idx * 10, 155 - idx * 25, color="AAAAAA")
    marker(p, MAIN_X + 400, 428, 4)

    # Popup floats over the upper right workspace but does not hide input controls.
    box(p, MAIN_X + 188, 253, 228, 58, "", "FFFFFF", "222222", line_width=1.4)
    box(p, MAIN_X + 188, 253, 228, 17, "SYSTEM MESSAGE", "E2E2E2", "222222", 6.5, True, "left")
    circle(p, MAIN_X + 196, 276, 20, "!", "FFFFFF", "555555", 8, True)
    label(p, MAIN_X + 222, 273, 153, 29, popup_text, 6.5, False, "left")
    box(p, MAIN_X + 379, 282, 29, 18, "OK", "E2E2E2", "444444", 6.3, True, rounded=True)
    marker(p, MAIN_X + 179, 245, 3)


def notes(p, number, s):
    stages = [("1", "INPUT", "; ".join(s[3])), ("2", "PROCESS", s[5]),
              ("3", "POPUP", s[6]), ("4", "OUTPUT", s[7])]
    label(p, 22, 520, 551, 16, "WORKFLOW NOTES", 8, True, "left", "444444")
    widths = [28, 65, 458]
    for idx, text in enumerate(["No.", "Stage", "Description"]):
        x = 22 + sum(widths[:idx])
        box(p, x, 539, widths[idx], 22, text, "D9D9D9", "555555", 7.2, True,
            "center" if idx == 0 else "left")
    for row, values in enumerate(stages):
        y = 561 + row * 39
        for idx, text in enumerate(values):
            x = 22 + sum(widths[:idx])
            box(p, x, y, widths[idx], 39, text, "FFFFFF" if row % 2 == 0 else "F7F7F7",
                "AAAAAA", 7, idx == 1, "center" if idx == 0 else "left")
    label(p, 22, 724, 551, 17,
          f"Data flow: {s[1]}  >  {s[4]}  >  e-Library processing  >  message  >  {s[0]} result",
          7.2, False, "left", "555555")
    label(p, 22, 746, 551, 30,
          "Design basis: supplied Context Diagram, DFD and Functional Decomposition Diagram. "
          "All objects remain editable in Microsoft Word.", 7, False, "left", "666666")


def draw_screen(p, number, s):
    base.CURRENT_PAGE[0] = number
    chrome(p, number, s)
    if s[8] == "dashboard":
        dashboard_content(p)
    elif s[8] in ("list", "report", "logs"):
        list_content(p, s[8])
    else:
        input_area(p, s)
    output_and_popup(p, s)
    notes(p, number, s)


def build():
    base.SHAPE_RECTS.clear()
    doc = Document()
    base.setup_section(doc.sections[0])
    doc.styles["Normal"].font.name = FONT
    doc.styles["Normal"].font.size = Pt(10)
    cover(doc)
    index_page(doc)
    guide_page(doc)
    for number, screen in enumerate(SCREENS, 1):
        p = base.anchor_paragraph(doc)
        draw_screen(p, number, screen)
        if number < len(SCREENS):
            doc.add_page_break()
    for section in doc.sections:
        base.setup_section(section)
    props = doc.core_properties
    props.title = "e-Library System - 21 Editable Word Shape Interfaces"
    props.subject = "Interface design based on the supplied CD, DFD and FDD"
    props.author = "Muhammad Haziq, Muhammad Danish and Seri Maisarah"
    props.keywords = "e-Library, Word shapes, editable interfaces, CD, DFD, FDD"
    doc.save(OUTPUT)
    print(f"CREATED={OUTPUT.name}")
    print(f"INTERFACES={len(SCREENS)}")
    print(f"SHAPES={len(base.SHAPE_RECTS)}")
    print(f"SIZE_BYTES={OUTPUT.stat().st_size}")


if __name__ == "__main__":
    build()
