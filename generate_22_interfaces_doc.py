"""Build a 22-interface design document for the Student Learning Hub System.

Format is intentionally different from the earlier 20-interface version:
  - one large full-window screen mockup per page (no 2x2 mini-panel grid)
  - numbered callout markers placed directly on the mockup
  - an element annotation grid below the mockup (Input / Process / Message / Output)
  - horizontal top navigation instead of a left sidebar
  - popup dialog floats on top of the same screen instead of a separate panel
  - no page numbers anywhere in the document

Everything is drawn with native, editable Microsoft Word shapes (DrawingML
wps:wsp). No pictures are embedded.
"""

from itertools import count
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "Student_Learning_Hub_22_Interface_Blueprint.docx"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

EMU_PER_PT = 12700
PAGE_W = 595.3
PAGE_H = 841.9
FONT = "Calibri"

SHAPE_IDS = count(1)
SHAPE_RECTS = []
CURRENT_PAGE = [0]

# ---------------------------------------------------------------- page geometry
BAND_X, BAND_Y, BAND_W, BAND_H = 24.0, 24.0, 547.0, 34.0
FRAME_X, FRAME_Y, FRAME_W, FRAME_H = 24.0, 68.0, 547.0, 436.0
TITLEBAR_H = 22.0
NAV_Y, NAV_H = 90.0, 32.0
BODY_X, BODY_Y, BODY_W, BODY_H = 40.0, 128.0, 515.0, 330.0
# The screen body is an L-shaped area: full-width interactive region on top,
# output panel at the lower left, and the popup dialog at the lower right.
UPPER_BOTTOM = 304.0
OUT_X, OUT_Y, OUT_W, OUT_H = 40.0, 312.0, 248.0, 144.0
MARGIN_MARK = 5.0    # input / process / output callouts sit in the left margin
RIGHT_MARK = 574.0   # action-button callouts sit in the right margin
PROC_X, PROC_Y, PROC_W, PROC_H = 32.0, 464.0, 531.0, 32.0
MODAL_X, MODAL_Y, MODAL_W, MODAL_H = 297.0, 316.0, 258.0, 140.0
ANNO_TITLE_Y = 510.0
HDR_Y, HDR_H = 530.0, 22.0
ROW_Y, ROW_H = 552.0, 30.0
NOTE_Y = 768.0
COL_W = [26.0, 122.0, 74.0, 325.0]

INK = "222222"
EDGE = "555555"
SOFT = "888888"
FAINT = "BBBBBB"
GREY_FILL = "F2F2F2"
BAND_FILL = "ECECEC"

NAV_ITEMS = ["Home", "Courses", "Materials", "Tasks", "Quizzes", "Forum"]

# ------------------------------------------------------------------ screen data
# name, module, layout, inputs, button, process, popup, popup kind,
# output, validation rule, next screen
SCREENS = [
    ("Login", "Access Control", "form",
     ["Student email", "Password", "Remember me"], "LOG IN",
     "System matches the email and password against the student account table.",
     "Login successful. Welcome back, Student.", "ok",
     "Student dashboard is loaded with the current semester summary.",
     "Both fields are compulsory and the password must be at least 8 characters.",
     "Student Dashboard"),
    ("Student Registration", "Access Control", "form",
     ["Full name", "Matric number", "Email", "Password"], "CREATE ACCOUNT",
     "System checks for duplicate matric number and stores the new record.",
     "Account created. Please log in to continue.", "ok",
     "Confirmation summary of the new account is displayed.",
     "Matric number must be unique and email must contain a valid domain.",
     "Login"),
    ("Forgot Password", "Access Control", "form",
     ["Registered email", "Matric number"], "SEND RESET LINK",
     "System generates a reset token and sends it to the registered email.",
     "A reset link has been sent to your email.", "ok",
     "Instruction panel with the reset link expiry time is displayed.",
     "Email must already exist in the student account table.",
     "Login"),
    ("Student Dashboard", "Main Hub", "dashboard",
     ["Select a summary tile", "Select a shortcut"], "OPEN",
     "System collects course, task and progress counters for the logged-in student.",
     "You have 2 assignments due this week.", "warning",
     "Tiles show enrolled courses, pending tasks, average marks and progress chart.",
     "Only data that belongs to the logged-in student is displayed.",
     "Course Catalogue"),
    ("Student Profile", "Student Record", "profile",
     ["Phone number", "Programme", "Short biography"], "SAVE CHANGES",
     "System validates the fields and updates the student record.",
     "Profile has been updated successfully.", "ok",
     "Updated profile details are shown together with the profile photo frame.",
     "Phone number must contain digits only, between 10 and 11 characters.",
     "Account Settings"),
    ("Account Settings", "Student Record", "settings",
     ["Email alert toggle", "Language", "New password"], "APPLY SETTINGS",
     "System saves the preference values and re-applies them to the session.",
     "Save these settings and sign out other devices?", "confirm",
     "Settings list is refreshed and shows the last updated time.",
     "New password cannot be the same as the previous password.",
     "Student Dashboard"),
    ("Course Catalogue", "Courses", "grid",
     ["Keyword", "Faculty filter", "Semester filter"], "SEARCH",
     "System filters the course table using the keyword and selected filters.",
     "Course has been added to your watch list.", "ok",
     "Course cards with code, credit hours and lecturer name are listed.",
     "Keyword must be at least 3 characters before searching.",
     "Course Details"),
    ("Course Details", "Courses", "detail",
     ["Select a tab", "Select a topic row"], "OPEN TOPIC",
     "System reads the selected course outline and lists its topics.",
     "Topic opened in the module viewer.", "ok",
     "Course synopsis, lecturer information and topic list are displayed.",
     "Only courses that are open for the current semester can be opened.",
     "Course Enrolment"),
    ("Course Enrolment", "Courses", "form",
     ["Semester", "Section", "Accept enrolment terms"], "ENROL NOW",
     "System checks credit limit, timetable clash and remaining seats.",
     "Confirm enrolment for this course and section?", "confirm",
     "Enrolment slip with status SUCCESSFUL and seat number is displayed.",
     "Total credit hours after enrolment cannot exceed 21.",
     "Class Timetable"),
    ("Learning Module Viewer", "Learning", "module",
     ["Select a lesson", "Mark lesson as done"], "MARK COMPLETE",
     "System records the lesson progress and recalculates the completion rate.",
     "Progress saved. Module completion is now 60 percent.", "ok",
     "Lesson content pane and the module progress bar are updated.",
     "A lesson can only be marked complete after the content is opened.",
     "Learning Materials"),
    ("Learning Materials", "Learning", "table",
     ["File type filter", "Search title"], "DOWNLOAD",
     "System prepares the selected file and records the download activity.",
     "File downloaded successfully to your device.", "ok",
     "Material table lists title, type, size and upload date.",
     "Only materials of the courses the student is enrolled in are listed.",
     "Assignment List"),
    ("Class Timetable", "Learning", "calendar",
     ["Select week", "Select a time slot"], "VIEW SLOT",
     "System maps the enrolled sections onto the weekly time grid.",
     "This slot clashes with another class.", "warning",
     "Weekly grid shows day, time, course code and venue for each class.",
     "A time slot can only hold one active class for the same student.",
     "Student Dashboard"),
    ("Assignment List", "Assessment", "list",
     ["Course filter", "Status filter"], "APPLY FILTER",
     "System sorts the assignment records by the nearest due date.",
     "One assignment is due today.", "warning",
     "Assignment rows show title, course, due date and submission status.",
     "Assignments that are already closed are shown as read only.",
     "Assignment Details"),
    ("Assignment Details", "Assessment", "detail",
     ["Select attachment", "Select rubric row"], "OPEN BRIEF",
     "System retrieves the assignment brief, rubric and attachment list.",
     "Attachment has been opened in a new tab.", "ok",
     "Brief, full mark, due date and marking rubric are displayed.",
     "The brief can only be opened after the release date.",
     "Assignment Submission"),
    ("Assignment Submission", "Assessment", "upload",
     ["Choose file (PDF or DOCX)", "Remark to lecturer"], "SUBMIT WORK",
     "System validates the file type and size, then uploads and time stamps it.",
     "Submit this file now? Submission cannot be changed later.", "confirm",
     "Submission receipt with file name, size and submitted time is displayed.",
     "File must be PDF or DOCX and not larger than 10 MB.",
     "Assignment Results"),
    ("Assignment Results", "Assessment", "result",
     ["Select a graded assignment"], "VIEW FEEDBACK",
     "System reads the mark, rubric score and lecturer feedback record.",
     "Marks for this assignment have been published.", "ok",
     "Score ring, rubric breakdown bars and lecturer comment are displayed.",
     "Only assignments with the status GRADED can be opened.",
     "Quiz List"),
    ("Quiz List", "Assessment", "list",
     ["Course filter", "Select available quiz"], "START QUIZ",
     "System checks the quiz window and remaining attempts for the student.",
     "The quiz timer of 15 minutes will start now.", "confirm",
     "Quiz rows show title, number of questions, duration and attempt count.",
     "A quiz can only be started inside its open and close date range.",
     "Quiz Attempt"),
    ("Quiz Attempt", "Assessment", "quiz",
     ["Select answer A to D", "Flag question"], "SUBMIT ANSWERS",
     "System stores each selected answer and keeps the countdown running.",
     "2 questions have not been answered yet.", "warning",
     "Question pane, option list and question navigator are updated.",
     "All questions must be answered before the paper can be submitted.",
     "Quiz Results"),
    ("Quiz Results", "Assessment", "result",
     ["Select review answers"], "REVIEW ANSWERS",
     "System marks the answer sheet and calculates the percentage score.",
     "Quiz completed. Your score is 80 percent.", "ok",
     "Score ring, correct and wrong count and per-question review list appear.",
     "The review list is only available after the quiz is submitted.",
     "Discussion Forum"),
    ("Discussion Forum", "Communication", "thread",
     ["Topic title", "Message body", "Attach file"], "POST TOPIC",
     "System saves the new topic and notifies the members of the course.",
     "Your topic has been posted to the forum.", "ok",
     "Thread list with topic title, author, reply count and last activity.",
     "Topic title cannot be empty and must not exceed 120 characters.",
     "Notification Centre"),
    ("Notification Centre", "Communication", "notifications",
     ["Category filter", "Select notification"], "MARK AS READ",
     "System updates the read flag and refreshes the unread counter.",
     "All notifications have been marked as read.", "ok",
     "Notification rows with icon tag, message, source and received time.",
     "Only notifications addressed to the logged-in student are listed.",
     "Log Out Confirmation"),
    ("Log Out Confirmation", "Access Control", "confirm",
     ["Select Yes or No"], "LOG OUT",
     "System closes the session, clears the token and returns to the login page.",
     "Are you sure you want to log out of the hub?", "confirm",
     "Login page is displayed with a session-closed notice.",
     "Unsaved work must be saved before the session is closed.",
     "Login"),
]


# ------------------------------------------------------------------- primitives
def tag(namespace, local):
    return f"{{{namespace}}}{local}"


def emu(value):
    return str(int(round(value * EMU_PER_PT)))


def _paragraph(container, value, size, bold, color, align, spacing_line=200):
    p = etree.SubElement(container, tag(W, "p"))
    ppr = etree.SubElement(p, tag(W, "pPr"))
    jc = etree.SubElement(ppr, tag(W, "jc"))
    jc.set(tag(W, "val"), align)
    spacing = etree.SubElement(ppr, tag(W, "spacing"))
    spacing.set(tag(W, "before"), "0")
    spacing.set(tag(W, "after"), "0")
    spacing.set(tag(W, "line"), str(spacing_line))
    spacing.set(tag(W, "lineRule"), "auto")
    run = etree.SubElement(p, tag(W, "r"))
    rpr = etree.SubElement(run, tag(W, "rPr"))
    fonts = etree.SubElement(rpr, tag(W, "rFonts"))
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(tag(W, attr), FONT)
    if bold:
        etree.SubElement(rpr, tag(W, "b"))
    color_el = etree.SubElement(rpr, tag(W, "color"))
    color_el.set(tag(W, "val"), color)
    half = str(int(round(size * 2)))
    etree.SubElement(rpr, tag(W, "sz")).set(tag(W, "val"), half)
    etree.SubElement(rpr, tag(W, "szCs")).set(tag(W, "val"), half)
    text_el = etree.SubElement(run, tag(W, "t"))
    text_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_el.text = value


def add_shape(paragraph, x, y, width, height, text="", geometry="rect", fill="FFFFFF",
              line=INK, line_width=1.0, size=8, bold=False, color=INK, align="center",
              dashed=False, arrow=False, flip_h=False, flip_v=False, name="Shape"):
    """Anchor one independently editable native Word shape to the page."""
    assert x >= 0 and y >= 0, f"negative origin: {name} at {x},{y}"
    assert x + width <= PAGE_W + 0.01, f"{name} overflows page width at x={x} w={width}"
    assert y + height <= PAGE_H + 0.01, f"{name} overflows page height at y={y} h={height}"
    SHAPE_RECTS.append({
        "page": CURRENT_PAGE[0], "x": x, "y": y, "w": width, "h": height, "text": text,
        "name": name, "geom": geometry, "fill": fill, "line": line, "size": size,
        "bold": bold, "align": align, "dashed": dashed, "color": color,
    })

    shape_id = next(SHAPE_IDS)
    run = paragraph.add_run()
    drawing = etree.SubElement(run._r, tag(W, "drawing"))
    anchor = etree.SubElement(drawing, tag(WP, "anchor"), {
        "distT": "0", "distB": "0", "distL": "0", "distR": "0", "simplePos": "0",
        "relativeHeight": str(251658240 + shape_id), "behindDoc": "0", "locked": "0",
        "layoutInCell": "1", "allowOverlap": "1",
    })
    etree.SubElement(anchor, tag(WP, "simplePos"), {"x": "0", "y": "0"})
    pos_h = etree.SubElement(anchor, tag(WP, "positionH"), {"relativeFrom": "page"})
    etree.SubElement(pos_h, tag(WP, "posOffset")).text = emu(x)
    pos_v = etree.SubElement(anchor, tag(WP, "positionV"), {"relativeFrom": "page"})
    etree.SubElement(pos_v, tag(WP, "posOffset")).text = emu(y)
    etree.SubElement(anchor, tag(WP, "extent"), {"cx": emu(width), "cy": emu(height)})
    etree.SubElement(anchor, tag(WP, "effectExtent"), {"l": "0", "t": "0", "r": "0", "b": "0"})
    etree.SubElement(anchor, tag(WP, "wrapNone"))
    etree.SubElement(anchor, tag(WP, "docPr"), {"id": str(shape_id), "name": f"{name} {shape_id}"})
    nv = etree.SubElement(anchor, tag(WP, "cNvGraphicFramePr"))
    etree.SubElement(nv, tag(A, "graphicFrameLocks"), {"noChangeAspect": "0"})

    graphic = etree.SubElement(anchor, tag(A, "graphic"))
    data = etree.SubElement(graphic, tag(A, "graphicData"), {
        "uri": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"})
    wsp = etree.SubElement(data, tag(WPS, "wsp"))
    cnv = etree.SubElement(wsp, tag(WPS, "cNvSpPr"))
    if text:
        cnv.set("txBox", "1")
    sppr = etree.SubElement(wsp, tag(WPS, "spPr"))
    xfrm_attrs = {}
    if flip_h:
        xfrm_attrs["flipH"] = "1"
    if flip_v:
        xfrm_attrs["flipV"] = "1"
    xfrm = etree.SubElement(sppr, tag(A, "xfrm"), xfrm_attrs)
    etree.SubElement(xfrm, tag(A, "off"), {"x": "0", "y": "0"})
    etree.SubElement(xfrm, tag(A, "ext"), {"cx": emu(width), "cy": emu(height)})
    geom = etree.SubElement(sppr, tag(A, "prstGeom"), {"prst": geometry})
    etree.SubElement(geom, tag(A, "avLst"))

    if fill is None:
        etree.SubElement(sppr, tag(A, "noFill"))
    else:
        solid = etree.SubElement(sppr, tag(A, "solidFill"))
        etree.SubElement(solid, tag(A, "srgbClr"), {"val": fill})

    ln = etree.SubElement(sppr, tag(A, "ln"), {"w": str(int(line_width * EMU_PER_PT))})
    if line is None:
        etree.SubElement(ln, tag(A, "noFill"))
    else:
        line_fill = etree.SubElement(ln, tag(A, "solidFill"))
        etree.SubElement(line_fill, tag(A, "srgbClr"), {"val": line})
        if dashed:
            etree.SubElement(ln, tag(A, "prstDash"), {"val": "dash"})
        if arrow:
            etree.SubElement(ln, tag(A, "tailEnd"), {"type": "triangle", "w": "med", "len": "med"})

    if text:
        txbx = etree.SubElement(wsp, tag(WPS, "txbx"))
        content = etree.SubElement(txbx, tag(W, "txbxContent"))
        for value in text.split("\n"):
            _paragraph(content, value, size, bold, color, align)
        body = etree.SubElement(wsp, tag(WPS, "bodyPr"), {
            "rot": "0", "spcFirstLastPara": "0", "vertOverflow": "overflow",
            "horzOverflow": "overflow", "vert": "horz", "wrap": "square",
            "lIns": emu(2.5), "tIns": emu(1.5), "rIns": emu(2.5), "bIns": emu(1.5),
            "numCol": "1", "spcCol": "0", "rtlCol": "0", "fromWordArt": "0",
            "anchor": "ctr", "anchorCtr": "0", "forceAA": "0", "compatLnSpc": "1"})
        etree.SubElement(body, tag(A, "spAutoFit"))
    else:
        body = etree.SubElement(wsp, tag(WPS, "bodyPr"), {"rot": "0"})
        etree.SubElement(body, tag(A, "spAutoFit"))
    return shape_id


def box(p, x, y, w, h, text="", fill="FFFFFF", line=EDGE, size=8, bold=False,
        align="center", rounded=False, dashed=False, color=INK, line_width=1.0):
    return add_shape(p, x, y, w, h, text, "roundRect" if rounded else "rect", fill, line,
                     line_width, size, bold, color, align, dashed,
                     name="Rounded Rectangle" if rounded else "Rectangle")


def label(p, x, y, w, h, text, size=8, bold=False, align="left", color=INK):
    return add_shape(p, x, y, w, h, text, "rect", None, None, 1.0, size, bold, color,
                     align, name="Text Box")


def circle(p, x, y, d, text="", fill="FFFFFF", line=EDGE, size=7, bold=True, color=INK):
    return add_shape(p, x, y, d, d, text, "ellipse", fill, line, 1.0, size, bold, color,
                     "center", name="Oval")


def rule(p, x, y, w, h=0.8, dashed=False, arrow=False, flip_h=False, color=SOFT, width=0.9):
    return add_shape(p, x, y, max(w, 0.8), max(h, 0.8), "", "line", None, color, width,
                     dashed=dashed, arrow=arrow, flip_h=flip_h,
                     name="Arrow" if arrow else "Line")


def marker(p, x, y, number):
    """Numbered callout marker used to link the mockup to the annotation grid."""
    circle(p, x, y, 17, str(number), "FFFFFF", "333333", 8, True)


# ------------------------------------------------------------- mockup chrome
def window_chrome(p, screen_name, module):
    box(p, FRAME_X, FRAME_Y, FRAME_W, FRAME_H, "", "FFFFFF", EDGE, line_width=1.3)
    box(p, FRAME_X, FRAME_Y, FRAME_W, TITLEBAR_H, "", GREY_FILL, EDGE)
    label(p, FRAME_X + 8, FRAME_Y + 4, 300, 15,
          f"Student Learning Hub  -  {screen_name}", 7.5, False, "left", "444444")
    label(p, FRAME_X + FRAME_W - 170, FRAME_Y + 4, 162, 15,
          f"Module: {module}", 7.5, False, "right", "666666")

    box(p, FRAME_X, NAV_Y, FRAME_W, NAV_H, "", "FFFFFF", FAINT)
    nav_x = FRAME_X + 10
    for index, item in enumerate(NAV_ITEMS):
        active = index == 0
        box(p, nav_x, NAV_Y + 7, 62, 18, item, GREY_FILL if active else "FFFFFF",
            EDGE if active else SOFT, 7, active, "center", rounded=True)
        nav_x += 66
    box(p, FRAME_X + FRAME_W - 140, NAV_Y + 7, 92, 18, "Search hub", "FFFFFF", SOFT,
        7, False, "left", rounded=True, color="777777")
    circle(p, FRAME_X + FRAME_W - 40, NAV_Y + 6, 20, "SA", GREY_FILL, SOFT, 7, True)


def section_caption(p, x, y, w, text):
    label(p, x, y, w, 14, text, 7.5, True, "left", "444444")
    rule(p, x, y + 14, w, color=FAINT)


def process_strip(p, text):
    box(p, PROC_X, PROC_Y, PROC_W, PROC_H, "", "FAFAFA", SOFT, dashed=True)
    box(p, PROC_X + 8, PROC_Y + 6, 92, 20, "PROCESSING", GREY_FILL, EDGE, 7, True,
        rounded=True)
    label(p, PROC_X + 108, PROC_Y + 6, PROC_W - 118, 20, text, 7.5, False, "left", "333333")
    marker(p, MARGIN_MARK, PROC_Y + 6, 3)


def popup(p, message, kind):
    """Dialog that floats above the screen body, inside the reserved lower-right zone."""
    box(p, MODAL_X, MODAL_Y, MODAL_W, MODAL_H, "", "FFFFFF", "222222", line_width=1.6)
    box(p, MODAL_X, MODAL_Y, MODAL_W, 22, "", GREY_FILL, "222222")
    heading_text = {"ok": "Information", "warning": "Warning", "confirm": "Please confirm"}[kind]
    label(p, MODAL_X + 8, MODAL_Y + 4, 150, 14, heading_text, 8, True, "left")
    label(p, MODAL_X + MODAL_W - 24, MODAL_Y + 4, 16, 14, "X", 8, False, "center", "555555")
    circle(p, MODAL_X + 12, MODAL_Y + 34, 28, "!" if kind != "ok" else "i", "FFFFFF",
           EDGE, 11, True)
    label(p, MODAL_X + 46, MODAL_Y + 28, MODAL_W - 58, 46, message, 8, False, "left")
    rule(p, MODAL_X + 10, MODAL_Y + 84, MODAL_W - 20, color=FAINT)
    if kind == "confirm":
        box(p, MODAL_X + MODAL_W - 156, MODAL_Y + 100, 72, 26, "No, cancel", "FFFFFF",
            EDGE, 7.5, False, rounded=True)
        box(p, MODAL_X + MODAL_W - 78, MODAL_Y + 100, 70, 26, "Yes, proceed", GREY_FILL,
            "333333", 7.5, True, rounded=True)
    else:
        box(p, MODAL_X + MODAL_W - 78, MODAL_Y + 100, 70, 26, "OK", GREY_FILL, "333333",
            8, True, rounded=True)
    marker(p, MODAL_X - 9, MODAL_Y - 9, 4)


# ---------------------------------------------------------------- body layouts
# Rule for every layout: interactive content stays in the upper full-width strip
# (y up to UPPER_BOTTOM). Anything drawn lower must stay left of x = 288 so the
# popup dialog in the lower-right never covers a control or a label.
COL2 = 300.0
NARROW = 246.0


def out_panel(p, title):
    """Output panel pinned to the lower-left of every screen body."""
    box(p, OUT_X, OUT_Y, OUT_W, OUT_H, "", "FFFFFF", EDGE, line_width=1.3)
    box(p, OUT_X, OUT_Y, OUT_W, 20, "", GREY_FILL, EDGE)
    label(p, OUT_X + 7, OUT_Y + 3, OUT_W - 14, 14, title, 7.5, True, "left")
    marker(p, MARGIN_MARK, OUT_Y - 6, 5)
    return OUT_X + 8, OUT_Y + 26, OUT_W - 16


def out_lines(p, ix, iy, iw, items):
    for index, (key, value) in enumerate(items):
        ry = iy + index * 22
        label(p, ix, ry, 92, 15, key, 7, False, "left", "777777")
        label(p, ix + 96, ry, iw - 96, 15, value, 7.5, True, "left")
        rule(p, ix, ry + 16, iw, color=FAINT)


def out_rows(p, ix, iy, iw, count=4, prefix="Row"):
    for index in range(count):
        box(p, ix, iy + index * 24, iw, 21, f"{prefix} {index + 1}    ____________",
            "FFFFFF", FAINT, 7, False, "left")


def out_bars(p, ix, iy, iw, items):
    for index, (name, pct) in enumerate(items):
        ry = iy + index * 24
        label(p, ix, ry, 78, 15, name, 7, False, "left")
        box(p, ix + 82, ry + 2, iw - 116, 11, "", "FFFFFF", SOFT)
        box(p, ix + 82, ry + 2, (iw - 116) * pct / 100.0, 11, "", GREY_FILL, EDGE)
        label(p, ix + iw - 30, ry, 30, 15, f"{pct}%", 7, False, "right", "555555")


def body_form(p, s):
    fields, action = s[3], s[4]
    section_caption(p, BODY_X, BODY_Y, 250, "Entry form")
    for index, field in enumerate(fields[:4]):
        fx = BODY_X + (index // 2) * 260
        fy = 150 + (index % 2) * 44
        tick = field.lower().startswith(("remember", "accept", "agree", "select"))
        if tick:
            label(p, fx, fy, 200, 13, "Option", 7.5, False, "left", "333333")
            box(p, fx, fy + 16, 14, 14, "", "FFFFFF", "555555")
            label(p, fx + 20, fy + 15, NARROW - 24, 16, field, 7.5, False, "left")
        else:
            label(p, fx, fy, 160, 13, f"{field} *", 7.5, False, "left", "333333")
            box(p, fx, fy + 14, NARROW, 24, "", "FFFFFF", SOFT, 7.5, False, "left")
            rule(p, fx + 8, fy + 20, 0.9, 12, color="999999")
    box(p, BODY_X, 240, 140, 28, action, GREY_FILL, "333333", 8.5, True, rounded=True)
    label(p, BODY_X, 274, NARROW, 26,
          "Fields marked with a star cannot be left empty.", 7, False, "left", "666666")
    box(p, COL2, 240, NARROW, 58, "", "FAFAFA", FAINT, dashed=True)
    label(p, COL2 + 8, 244, 230, 50,
          "Wrong entry is shown in small red text right under the field,\nand the button stays inactive until every star field is filled.",
          7, False, "left", "555555")
    marker(p, MARGIN_MARK, 145, 1)
    marker(p, BODY_X + 148, 246, 2)

    ix, iy, iw = out_panel(p, "Output: entry result")
    out_lines(p, ix, iy, iw, [("Status", "Accepted"), ("Record", "SLH-00231"),
                              ("Saved at", "9.14 pm"), ("Next", s[10])])


def body_dashboard(p, s):
    section_caption(p, BODY_X, BODY_Y, 250, "Summary tiles")
    tiles = [("Enrolled courses", "5"), ("Pending tasks", "2"),
             ("Average mark", "78"), ("Unread notices", "4")]
    for index, (title, value) in enumerate(tiles):
        tx = BODY_X + index * 130
        box(p, tx, 150, 120, 58, "", "FFFFFF", SOFT, rounded=True)
        label(p, tx + 8, 154, 104, 13, title, 7, False, "left", "666666")
        label(p, tx + 8, 170, 104, 32, value, 15, True, "left")
    section_caption(p, BODY_X, 216, 250, "Progress by course")
    for index, (name, pct) in enumerate([("CSC101", 40), ("CSC205", 65), ("MAT120", 85)]):
        ry = 238 + index * 22
        label(p, BODY_X, ry, 82, 15, name, 7.5, False, "left")
        box(p, BODY_X + 88, ry + 2, 170, 12, "", "FFFFFF", SOFT)
        box(p, BODY_X + 88, ry + 2, 170 * pct / 100.0, 12, "", GREY_FILL, EDGE)
        label(p, BODY_X + 264, ry, 34, 15, f"{pct}%", 7, False, "left", "555555")
    section_caption(p, 364, 216, 191, "Shortcuts")
    for index, name in enumerate(["Submit a task", "Open next quiz"]):
        box(p, 364, 238 + index * 30, 191, 24, name, "FFFFFF", SOFT, 7.5, False, "left",
            rounded=True)
    marker(p, MARGIN_MARK, 142, 1)
    marker(p, 344, 240, 2)

    ix, iy, iw = out_panel(p, "Output: dashboard summary")
    out_lines(p, ix, iy, iw, [("Courses", "5 active"), ("Tasks due", "2 this week"),
                              ("Average mark", "78 of 100"), ("Refreshed", "9.14 pm")])


def body_profile(p, s):
    section_caption(p, BODY_X, BODY_Y, 200, "Profile card")
    box(p, BODY_X, 150, 180, 148, "", "FAFAFA", SOFT)
    circle(p, BODY_X + 58, 158, 64, "Photo\nframe", "FFFFFF", SOFT, 7, False)
    label(p, BODY_X + 8, 228, 164, 15, "Student Aiman", 9, True, "center")
    label(p, BODY_X + 8, 244, 164, 13, "Matric 2024100234", 7, False, "center", "666666")
    box(p, BODY_X + 30, 266, 120, 24, "Change photo", "FFFFFF", EDGE, 7.5, False,
        rounded=True)
    section_caption(p, 236, BODY_Y, 319, "Editable details")
    for index, field in enumerate(s[3][:3]):
        fy = 150 + index * 44
        label(p, 236, fy, 160, 13, field, 7.5, False, "left", "333333")
        box(p, 236, fy + 14, 300, 24, "", "FFFFFF", SOFT, 7.5, False, "left")
    box(p, 236, 280, 130, 24, s[4], GREY_FILL, "333333", 8, True, rounded=True)
    box(p, 374, 280, 90, 24, "Cancel", "FFFFFF", EDGE, 8, False, rounded=True)
    marker(p, RIGHT_MARK, 150, 1)
    marker(p, 472, 284, 2)

    ix, iy, iw = out_panel(p, "Output: saved profile")
    out_lines(p, ix, iy, iw, [("Name", "Student Aiman"), ("Phone", "011 2345 678"),
                              ("Programme", "Diploma in IT"), ("Updated", "9.14 pm")])


def body_settings(p, s):
    section_caption(p, BODY_X, BODY_Y, 250, "Preference list")
    rows = [("Email alert", "Sends a mail for every new task", "toggle"),
            ("Push notification", "Shows an alert on the phone", "toggle"),
            ("Interface language", "Language used across the hub", "select")]
    for index, (name, note, kind) in enumerate(rows):
        ry = 150 + index * 44
        box(p, BODY_X, ry, 340, 38, "", "FFFFFF", FAINT)
        label(p, BODY_X + 10, ry + 4, 200, 15, name, 8, True, "left")
        label(p, BODY_X + 10, ry + 19, 220, 14, note, 6.5, False, "left", "777777")
        if kind == "toggle":
            box(p, BODY_X + 268, ry + 10, 46, 18, "", GREY_FILL, EDGE, rounded=True)
            circle(p, BODY_X + 296, ry + 11, 16, "", "FFFFFF", "333333")
        else:
            box(p, BODY_X + 250, ry + 9, 90, 20, "English  v", "FFFFFF", SOFT, 7, False,
                "left")
    label(p, BODY_X, 282, 90, 14, "New password", 7.5, False, "left", "333333")
    box(p, BODY_X + 96, 280, 130, 20, "* * * * * * * *", "FFFFFF", SOFT, 7, False, "left")
    box(p, 396, 150, 159, 126, "", "FAFAFA", FAINT, dashed=True)
    label(p, 404, 156, 145, 14, "Last change", 7.5, True, "left")
    label(p, 404, 174, 145, 96,
          "12 May, 9.14 pm\nby this account\nfrom campus network\n\nSigning out other\ndevices is optional.",
          7, False, "left", "555555")
    box(p, 396, 282, 159, 22, s[4], GREY_FILL, "333333", 8, True, rounded=True)
    marker(p, MARGIN_MARK, 144, 1)
    marker(p, 376, 276, 2)

    ix, iy, iw = out_panel(p, "Output: applied settings")
    out_lines(p, ix, iy, iw, [("Email alert", "On"), ("Push alert", "Off"),
                              ("Language", "English"), ("Password", "Changed")])


def body_grid(p, s):
    section_caption(p, BODY_X, BODY_Y, 515, "Search bar and filters")
    box(p, BODY_X, 150, 200, 22, "Type a course name or code", "FFFFFF", SOFT, 7.5, False,
        "left", color="777777")
    box(p, 248, 150, 90, 22, "Faculty  v", "FFFFFF", SOFT, 7.5, False, "left")
    box(p, 344, 150, 90, 22, "Semester  v", "FFFFFF", SOFT, 7.5, False, "left")
    box(p, 440, 150, 115, 22, s[4], GREY_FILL, "333333", 8, True, rounded=True)
    for index, code in enumerate(["CSC101", "CSC205", "MAT120"]):
        cx = BODY_X + index * 174
        box(p, cx, 182, 164, 116, "", "FFFFFF", SOFT, rounded=True)
        box(p, cx + 8, 190, 54, 18, code, GREY_FILL, EDGE, 7, True, rounded=True)
        label(p, cx + 8, 212, 148, 28, "Course title over\ntwo short lines", 7.5, False, "left")
        label(p, cx + 8, 242, 120, 13, "3 credit hours", 6.5, False, "left", "777777")
        label(p, cx + 8, 256, 84, 13, "Seats left: 12", 6.5, False, "left", "777777")
        box(p, cx + 96, 268, 60, 22, "Open", "FFFFFF", EDGE, 7, False, rounded=True)
    marker(p, MARGIN_MARK, 144, 1)
    marker(p, RIGHT_MARK, 150, 2)

    ix, iy, iw = out_panel(p, "Output: matching courses")
    label(p, ix, iy, iw, 14, "3 of 48 courses match the keyword", 7.5, True, "left")
    out_rows(p, ix, iy + 20, iw, 3, "Course")


def body_detail(p, s):
    section_caption(p, BODY_X, BODY_Y, 340, "Detail pane")
    for index, name in enumerate(["Overview", "Topics", "Files", "Rubric"]):
        box(p, BODY_X + index * 78, 148, 74, 22, name,
            GREY_FILL if index == 1 else "FFFFFF", EDGE if index == 1 else SOFT, 7,
            index == 1, rounded=True)
    box(p, BODY_X, 178, 336, 58, "", "FAFAFA", FAINT)
    label(p, BODY_X + 8, 182, 320, 50,
          "Description block keyed in by the lecturer.\n________________________________________\n________________________________________",
          7, False, "left", "555555")
    for index in range(2):
        ry = 242 + index * 30
        box(p, BODY_X, ry, 336, 26, "", "FFFFFF", FAINT)
        circle(p, BODY_X + 7, ry + 5, 16, str(index + 1), "FFFFFF", SOFT, 6.5, False)
        label(p, BODY_X + 30, ry + 5, 200, 16, f"Topic row {index + 1}", 7.5, False, "left")
        label(p, BODY_X + 246, ry + 5, 84, 16, "Not opened", 6.5, False, "left", "777777")
    section_caption(p, 388, BODY_Y, 167, "Info panel")
    box(p, 388, 148, 167, 110, "", "FFFFFF", SOFT)
    label(p, 396, 154, 150, 98,
          "Lecturer\n____________\n\nCredit hours\n____________\n\nSeats or full mark\n____________",
          7, False, "left", "555555")
    box(p, 388, 266, 167, 26, s[4], GREY_FILL, "333333", 8.5, True, rounded=True)
    marker(p, MARGIN_MARK, 142, 1)
    marker(p, RIGHT_MARK, 268, 2)

    ix, iy, iw = out_panel(p, "Output: opened detail")
    out_lines(p, ix, iy, iw, [("Title", "Topic 2 brief"), ("Released", "10 May"),
                              ("Full mark", "100"), ("Status", "Open")])


def body_module(p, s):
    section_caption(p, BODY_X, BODY_Y, 140, "Lesson list")
    for index in range(3):
        ly = 150 + index * 30
        box(p, BODY_X, ly, 132, 26, "", GREY_FILL if index == 1 else "FFFFFF", SOFT)
        label(p, BODY_X + 8, ly + 5, 92, 16, f"Lesson {index + 1}", 7.5, index == 1, "left")
        label(p, BODY_X + 100, ly + 5, 28, 16, "done" if index < 1 else "", 6, False,
              "left", "777777")
    label(p, BODY_X, 240, 80, 13, "Progress", 7.5, True, "left")
    box(p, BODY_X, 256, 120, 12, "", "FFFFFF", SOFT)
    box(p, BODY_X, 256, 72, 12, "", GREY_FILL, EDGE)
    label(p, BODY_X + 126, 254, 20, 15, "60%", 7, False, "left", "555555")
    box(p, BODY_X, 276, 120, 26, s[4], GREY_FILL, "333333", 8.5, True, rounded=True)
    section_caption(p, 190, BODY_Y, 365, "Content viewer")
    box(p, 190, 148, 365, 150, "", "FFFFFF", SOFT)
    label(p, 200, 154, 340, 16, "Lesson 2: Data and information", 8, True, "left")
    for index in range(5):
        rule(p, 200, 180 + index * 16, 330 - index * 40, color=FAINT)
    box(p, 200, 264, 120, 26, "Play video", "FFFFFF", EDGE, 7.5, False, rounded=True)
    box(p, 330, 264, 120, 26, "Open slide", "FFFFFF", EDGE, 7.5, False, rounded=True)
    marker(p, MARGIN_MARK, 144, 1)
    marker(p, 166, 278, 2)

    ix, iy, iw = out_panel(p, "Output: progress record")
    out_lines(p, ix, iy, iw, [("Lesson", "2 of 5 done"), ("Module", "60 percent"),
                              ("Time spent", "42 minutes"), ("Next", s[10])])


def body_table(p, s):
    section_caption(p, BODY_X, BODY_Y, 515, "Filter row")
    box(p, BODY_X, 148, 140, 22, "Search title", "FFFFFF", SOFT, 7.5, False, "left",
        color="777777")
    box(p, 188, 148, 100, 22, "File type  v", "FFFFFF", SOFT, 7.5, False, "left")
    box(p, 296, 148, 100, 22, "Course  v", "FFFFFF", SOFT, 7.5, False, "left")
    box(p, 440, 148, 115, 22, s[4], GREY_FILL, "333333", 8, True, rounded=True)
    heads = [("Title", 190), ("Type", 66), ("Size", 66), ("Uploaded", 86), ("Action", 95)]
    hx = BODY_X
    for name, width in heads:
        box(p, hx, 178, width, 22, name, GREY_FILL, EDGE, 7.5, True)
        hx += width
    for row in range(3):
        ry = 200 + row * 26
        rx = BODY_X
        for col, (_, width) in enumerate(heads):
            if col == 4:
                box(p, rx, ry, width, 26, "", "FFFFFF", FAINT)
                box(p, rx + 16, ry + 3, 62, 20, "Get file", "FFFFFF", SOFT, 7, False,
                    rounded=True)
            else:
                filler = f"Material {row + 1}" if col == 0 else "______"
                box(p, rx, ry, width, 26, filler, "FFFFFF", FAINT, 7, False,
                    "left" if col == 0 else "center")
            rx += width
    marker(p, MARGIN_MARK, 142, 1)
    marker(p, RIGHT_MARK, 150, 2)

    ix, iy, iw = out_panel(p, "Output: download receipt")
    out_lines(p, ix, iy, iw, [("File", "week2_notes.pdf"), ("Size", "1.8 MB"),
                              ("Saved to", "Device folder"), ("Logged", "Yes")])


def body_calendar(p, s):
    section_caption(p, BODY_X, BODY_Y, 515, "Week selector")
    box(p, BODY_X, 148, 32, 22, "<", "FFFFFF", SOFT, 8, True)
    box(p, 76, 148, 150, 22, "Week 7  -  12 to 16 May", "FFFFFF", SOFT, 7.5, False)
    box(p, 230, 148, 32, 22, ">", "FFFFFF", SOFT, 8, True)
    box(p, 440, 148, 115, 22, s[4], GREY_FILL, "333333", 8, True, rounded=True)
    box(p, BODY_X, 178, 54, 20, "Time", GREY_FILL, EDGE, 7, True)
    for index, day in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri"]):
        box(p, 94 + index * 90, 178, 90, 20, day, GREY_FILL, EDGE, 7, True)
    filled = {(0, 0): "CSC101  DK2", (0, 3): "MAT120  BK5",
              (1, 1): "CSC205  Lab3", (2, 2): "ENG150  DK1", (2, 4): "STA210  BK2"}
    for row, time_text in enumerate(["8 am", "10 am", "12 pm"]):
        ry = 198 + row * 34
        box(p, BODY_X, ry, 54, 34, time_text, "FFFFFF", FAINT, 7, False)
        for col in range(5):
            text = filled.get((row, col), "")
            box(p, 94 + col * 90, ry, 90, 34, text, "FFFFFF" if not text else GREY_FILL,
                FAINT if not text else EDGE, 6.5, bool(text))
    marker(p, MARGIN_MARK, 142, 1)
    marker(p, RIGHT_MARK, 150, 2)

    ix, iy, iw = out_panel(p, "Output: selected slot")
    out_lines(p, ix, iy, iw, [("Course", "CSC101"), ("Day and time", "Monday 8 am"),
                              ("Venue", "DK2, Block A"), ("Lecturer", "Madam Murni")])


def body_list(p, s):
    section_caption(p, BODY_X, BODY_Y, 515, "Filter and sort")
    box(p, BODY_X, 148, 120, 22, "Course  v", "FFFFFF", SOFT, 7.5, False, "left")
    box(p, 168, 148, 120, 22, "Status  v", "FFFFFF", SOFT, 7.5, False, "left")
    box(p, 296, 148, 110, 22, "Due date  v", "FFFFFF", SOFT, 7.5, False, "left")
    box(p, 440, 148, 115, 22, s[4], GREY_FILL, "333333", 8, True, rounded=True)
    for row in range(2):
        ry = 180 + row * 60
        box(p, BODY_X, ry, 515, 54, "", "FFFFFF", FAINT)
        box(p, BODY_X, ry, 5, 54, "", EDGE, None)
        label(p, BODY_X + 14, ry + 6, 260, 16, f"Item title number {row + 1}", 8, True, "left")
        label(p, BODY_X + 14, ry + 24, 260, 15, "Course code  -  short note line", 6.5,
              False, "left", "777777")
        label(p, BODY_X + 288, ry + 18, 84, 16, "Due 20 May", 7, False, "left", "555555")
        box(p, 420, ry + 16, 60, 22, "Open", GREY_FILL, EDGE, 7, False, rounded=True)
        box(p, 488, ry + 16, 60, 22, "Pending", "FFFFFF", SOFT, 6.5, False, rounded=True)
    marker(p, MARGIN_MARK, 142, 1)
    marker(p, RIGHT_MARK, 150, 2)

    ix, iy, iw = out_panel(p, "Output: filtered rows")
    label(p, ix, iy, iw, 14, "6 rows found, sorted by nearest due date", 7.5, True, "left")
    out_rows(p, ix, iy + 20, iw, 3, "Row")


def body_upload(p, s):
    section_caption(p, BODY_X, BODY_Y, 320, "Upload area")
    box(p, BODY_X, 148, 320, 86, "", "FAFAFA", EDGE, dashed=True)
    label(p, BODY_X + 20, 158, 280, 38,
          "Drag the file here or use the button below.\nAccepted: PDF and DOCX, up to 10 MB.",
          7.5, False, "center", "555555")
    box(p, 140, 202, 120, 24, "Choose file", "FFFFFF", "333333", 8, True, rounded=True)
    box(p, BODY_X, 242, 320, 38, "", "FFFFFF", SOFT)
    box(p, BODY_X + 8, 250, 26, 22, "PDF", GREY_FILL, EDGE, 6, True)
    label(p, BODY_X + 42, 246, 180, 15, "assignment_final.pdf", 7.5, False, "left")
    label(p, BODY_X + 42, 261, 180, 13, "2.4 MB  -  ready to send", 6.5, False, "left",
          "777777")
    box(p, 296, 252, 56, 20, "Remove", "FFFFFF", SOFT, 6.5, False, rounded=True)
    label(p, BODY_X, 286, 80, 14, "Upload bar", 7.5, True, "left")
    box(p, 124, 286, 150, 12, "", "FFFFFF", SOFT)
    box(p, 124, 286, 104, 12, "", GREY_FILL, EDGE)
    section_caption(p, 376, BODY_Y, 179, "Remark and submit")
    box(p, 376, 148, 179, 86, "", "FFFFFF", SOFT)
    label(p, 382, 152, 166, 14, "Remark to lecturer", 6.5, False, "left", "777777")
    for index in range(3):
        rule(p, 382, 176 + index * 16, 160 - index * 30, color=FAINT)
    box(p, 376, 242, 179, 26, s[4], GREY_FILL, "333333", 8.5, True, rounded=True)
    label(p, 376, 274, 179, 26, "Attempt 1 of 2. Draft saved at 8.05 pm.", 7, False,
          "left", "666666")
    marker(p, MARGIN_MARK, 142, 1)
    marker(p, 356, 236, 2)

    ix, iy, iw = out_panel(p, "Output: submission receipt")
    out_lines(p, ix, iy, iw, [("File", "assignment_final.pdf"), ("Size", "2.4 MB"),
                              ("Time in", "9.14 pm"), ("Status", "Received")])


def body_result(p, s):
    section_caption(p, BODY_X, BODY_Y, 280, "Score summary")
    circle(p, BODY_X, 150, 104, "80%\nscore", "FFFFFF", "333333", 13, True)
    box(p, 160, 150, 150, 46, "", "FFFFFF", SOFT)
    label(p, 168, 154, 134, 14, "Correct answers", 7, False, "left", "777777")
    label(p, 168, 168, 134, 24, "8 of 10", 12, True, "left")
    box(p, 160, 204, 150, 46, "", "FFFFFF", SOFT)
    label(p, 168, 208, 134, 14, "Grade band", 7, False, "left", "777777")
    label(p, 168, 222, 134, 24, "A minus", 12, True, "left")
    section_caption(p, 330, BODY_Y, 225, "Rubric breakdown")
    for index, part in enumerate(["Content", "Structure", "Reference", "Language"]):
        ry = 150 + index * 22
        label(p, 330, ry, 72, 15, part, 7, False, "left")
        box(p, 406, ry + 2, 100, 11, "", "FFFFFF", SOFT)
        box(p, 406, ry + 2, 40 + index * 15, 11, "", GREY_FILL, EDGE)
        label(p, 512, ry, 40, 15, f"{15 + index * 5}", 7, False, "left", "555555")
    box(p, 330, 244, 225, 54, "", "FAFAFA", FAINT)
    label(p, 338, 248, 210, 46,
          "Lecturer comment block.\n_______________________________\n_______________________________",
          7, False, "left", "555555")
    box(p, 160, 262, 150, 26, s[4], GREY_FILL, "333333", 8.5, True, rounded=True)
    marker(p, MARGIN_MARK, 144, 1)
    marker(p, 138, 256, 2)

    ix, iy, iw = out_panel(p, "Output: score record")
    out_lines(p, ix, iy, iw, [("Score", "80 of 100"), ("Correct", "8 of 10"),
                              ("Grade", "A minus"), ("Published", "12 May")])


def body_quiz(p, s):
    label(p, BODY_X, BODY_Y, 220, 16, "Question 3 of 10", 9, True, "left")
    box(p, 400, 126, 155, 22, "Time left  09 : 42", GREY_FILL, "333333", 8, True,
        rounded=True)
    box(p, BODY_X, 152, 340, 46, "", "FAFAFA", SOFT)
    label(p, BODY_X + 10, 158, 322, 34,
          "Question text sits here over two lines so the wording of\nthe item stays clear to the student.",
          7.5, False, "left")
    for index, letter in enumerate(["A", "B", "C", "D"]):
        oy = 204 + index * 24
        box(p, BODY_X, oy, 340, 21, "", "FFFFFF" if index != 1 else GREY_FILL, SOFT)
        circle(p, BODY_X + 6, oy + 3, 15, letter, "FFFFFF", EDGE, 6.5, True)
        label(p, BODY_X + 28, oy + 3, 300, 15, f"Answer option {letter}", 7.5, index == 1,
              "left")
    section_caption(p, 396, 152, 159, "Question navigator")
    for index in range(10):
        col, row = index % 5, index // 5
        box(p, 396 + col * 32, 174 + row * 30, 28, 26, str(index + 1),
            GREY_FILL if index < 3 else "FFFFFF", EDGE if index < 3 else SOFT, 7,
            index < 3)
    label(p, 396, 238, 159, 26, "A shaded number means that\nitem is already answered.",
          6.5, False, "left", "666666")
    box(p, 396, 270, 159, 26, s[4], GREY_FILL, "333333", 8, True, rounded=True)
    marker(p, MARGIN_MARK, 204, 1)
    marker(p, 376, 264, 2)

    ix, iy, iw = out_panel(p, "Output: answer sheet")
    out_lines(p, ix, iy, iw, [("Answered", "8 of 10"), ("Flagged", "1 item"),
                              ("Time used", "05 : 18"), ("Saved", "Every click")])


def body_thread(p, s):
    section_caption(p, BODY_X, BODY_Y, NARROW, "New topic composer")
    box(p, BODY_X, 150, NARROW, 22, "Topic title", "FFFFFF", SOFT, 7.5, False, "left",
        color="777777")
    box(p, BODY_X, 178, NARROW, 74, "", "FFFFFF", SOFT)
    label(p, BODY_X + 8, 182, 230, 14, "Message body", 6.5, False, "left", "777777")
    for index in range(3):
        rule(p, BODY_X + 8, 206 + index * 14, 220 - index * 50, color=FAINT)
    box(p, BODY_X, 258, 100, 24, "Attach file", "FFFFFF", EDGE, 7.5, False, rounded=True)
    box(p, 166, 258, 120, 24, s[4], GREY_FILL, "333333", 8, True, rounded=True)
    section_caption(p, COL2, BODY_Y, 255, "Thread list")
    for index in range(3):
        ty = 150 + index * 50
        box(p, COL2, ty, 255, 44, "", "FFFFFF", FAINT)
        label(p, COL2 + 8, ty + 4, 180, 15, f"Thread title {index + 1}", 7.5, True, "left")
        label(p, COL2 + 8, ty + 19, 180, 14, "by Student  -  2 hours ago", 6.5, False,
              "left", "777777")
        label(p, COL2 + 190, ty + 4, 60, 14, f"{index + 2} replies", 6.5, False, "left",
              "555555")
        box(p, COL2 + 188, ty + 20, 60, 20, "Reply", "FFFFFF", SOFT, 6.5, False,
            rounded=True)
    marker(p, MARGIN_MARK, 144, 1)
    marker(p, 144, 260, 2)

    ix, iy, iw = out_panel(p, "Output: posted topic")
    out_lines(p, ix, iy, iw, [("Topic", "Question on task 2"), ("Posted", "9.14 pm"),
                              ("Course", "CSC101"), ("Members told", "38")])


def body_notifications(p, s):
    section_caption(p, BODY_X, BODY_Y, 515, "Category tabs and action")
    for index, name in enumerate(["All", "Task", "Quiz", "Forum", "Notice"]):
        box(p, BODY_X + index * 74, 148, 70, 22, name,
            GREY_FILL if index == 0 else "FFFFFF", EDGE if index == 0 else SOFT, 7,
            index == 0, rounded=True)
    box(p, 440, 146, 115, 26, s[4], GREY_FILL, "333333", 8, True, rounded=True)
    for index, tag_text in enumerate(["TASK", "QUIZ", "NOTE"]):
        ry = 180 + index * 40
        unread = index < 2
        box(p, BODY_X, ry, 515, 36, "", GREY_FILL if unread else "FFFFFF", FAINT)
        box(p, BODY_X + 8, ry + 7, 44, 22, tag_text, "FFFFFF", EDGE, 6.5, True,
            rounded=True)
        label(p, BODY_X + 60, ry + 3, 280, 15, f"Notification message line {index + 1}",
              7.5, unread, "left")
        label(p, BODY_X + 60, ry + 18, 280, 14, "from CSC101  -  system generated", 6.5,
              False, "left", "777777")
        label(p, 384, ry + 10, 66, 15, "10 min ago", 6.5, False, "left", "555555")
        box(p, 456, ry + 7, 90, 22, "Open", "FFFFFF", SOFT, 6.5, False, rounded=True)
    marker(p, MARGIN_MARK, 142, 1)
    marker(p, RIGHT_MARK, 150, 2)

    ix, iy, iw = out_panel(p, "Output: notification state")
    out_lines(p, ix, iy, iw, [("Total", "12 notifications"), ("Unread before", "4"),
                              ("Unread now", "0"), ("Refreshed", "9.14 pm")])


def body_confirm(p, s):
    section_caption(p, BODY_X, BODY_Y, 515, "Session panel behind the dialog")
    box(p, BODY_X, 150, 250, 74, "", "FFFFFF", SOFT)
    circle(p, BODY_X + 10, 160, 44, "SA", GREY_FILL, EDGE, 10, True)
    label(p, BODY_X + 66, 162, 170, 16, "Student Aiman", 8.5, True, "left")
    label(p, BODY_X + 66, 180, 176, 34, "Signed in since 8.02 pm\nDevice: campus laptop",
          7, False, "left", "666666")
    box(p, COL2, 150, 255, 74, "", "FAFAFA", FAINT, dashed=True)
    label(p, COL2 + 10, 156, 235, 62,
          "Items still open in this session:\n- 1 assignment draft\n- 1 forum reply not posted",
          7, False, "left", "555555")
    box(p, BODY_X, 236, 150, 26, s[4], GREY_FILL, "333333", 8.5, True, rounded=True)
    box(p, 200, 236, 150, 26, "Stay signed in", "FFFFFF", EDGE, 8.5, False, rounded=True)
    label(p, BODY_X, 272, NARROW, 28,
          "The dialog blocks the page until one button is picked.", 7, False, "left",
          "666666")
    marker(p, MARGIN_MARK, 228, 1)
    marker(p, 356, 238, 2)

    ix, iy, iw = out_panel(p, "Output: session closed")
    out_lines(p, ix, iy, iw, [("Session", "Ended"), ("Token", "Cleared"),
                              ("Signed out at", "9.14 pm"), ("Next", s[10])])


LAYOUTS = {
    "form": body_form, "dashboard": body_dashboard, "profile": body_profile,
    "settings": body_settings, "grid": body_grid, "detail": body_detail,
    "module": body_module, "table": body_table, "calendar": body_calendar,
    "list": body_list, "upload": body_upload, "result": body_result,
    "quiz": body_quiz, "thread": body_thread, "notifications": body_notifications,
    "confirm": body_confirm,
}


# ------------------------------------------------------------- annotation grid
def annotation_grid(p, number, s):
    name, module, layout, inputs, action, process, message, kind = s[:8]
    output, validation, next_screen = s[8], s[9], s[10]
    kind_word = {"ok": "Information popup", "warning": "Warning popup",
                 "confirm": "Confirmation popup"}[kind]

    label(p, BAND_X, ANNO_TITLE_Y, 400, 15,
          "Element notes (numbers match the callouts on the mockup)", 8, True, "left")
    heads = ["No.", "Element on screen", "Category", "What happens"]
    hx = BAND_X
    for text, width in zip(heads, COL_W):
        box(p, hx, HDR_Y, width, HDR_H, text, BAND_FILL, EDGE, 7.5, True,
            "center" if text == "No." else "left")
        hx += width

    rows = [
        ("1", "Input field or control", "Input", "; ".join(inputs)),
        ("2", f"Button: {action}", "Input", "Student clicks it to send the entered data to the system."),
        ("3", "Processing strip", "Process", process),
        ("4", "Popup dialog", "Message", f"{kind_word}: \"{message}\""),
        ("5", "Result display area", "Output", output),
        ("-", "Validation rule", "Rule", validation),
        ("-", "Screen after this one", "Flow", f"Goes to: {next_screen}"),
    ]
    for index, row in enumerate(rows):
        ry = ROW_Y + index * ROW_H
        rx = BAND_X
        for col, (text, width) in enumerate(zip(row, COL_W)):
            fill = "FFFFFF" if index % 2 == 0 else "FAFAFA"
            box(p, rx, ry, width, ROW_H, text, fill, FAINT, 7.5,
                col == 1 and row[0] != "-", "center" if col == 0 else "left")
            rx += width

    label(p, BAND_X, NOTE_Y, 547, 16,
          f"Screen flow  {number:02d}  Input  >  Process  >  Popup message  >  Output  >  {next_screen}",
          7.5, False, "left", "555555")


def title_band(p, number, s):
    box(p, BAND_X, BAND_Y, BAND_W, BAND_H, "", BAND_FILL, EDGE, line_width=1.2)
    box(p, BAND_X + 8, BAND_Y + 7, 54, 20, f"IF-{number:02d}", "FFFFFF", "333333", 8.5,
        True, rounded=True)
    label(p, BAND_X + 70, BAND_Y + 8, 268, 18, s[0], 12, True, "left")
    label(p, BAND_X + BAND_W - 200, BAND_Y + 9, 192, 16,
          f"{s[1]}  |  {s[2].replace('_', ' ').title()} layout", 8, False, "right", "444444")


def draw_page(p, number, s):
    CURRENT_PAGE[0] = number
    title_band(p, number, s)
    window_chrome(p, s[0], s[1])
    LAYOUTS[s[2]](p, s)
    popup(p, s[6], s[7])
    process_strip(p, s[5])
    annotation_grid(p, number, s)


# ------------------------------------------------------------------ docx setup
def setup_section(section):
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    for attr in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(section, attr, Inches(0.3))
    section.header_distance = Inches(0.1)
    section.footer_distance = Inches(0.1)


def shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def cell_margins(cell, top=60, start=90, bottom=60, end=90):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for tag_name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{tag_name}"))
        if node is None:
            node = OxmlElement(f"w:{tag_name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def text_line(doc, text, size=11, bold=False, align="left", space_before=0, space_after=4,
              color=(34, 34, 34)):
    p = doc.add_paragraph()
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "right": WD_ALIGN_PARAGRAPH.RIGHT}[align]
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color)
    return p


def anchor_paragraph(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    return p


# ----------------------------------------------------------------- front pages
def cover_page(doc):
    text_line(doc, "STUDENT LEARNING HUB SYSTEM", 20, True, "center", 96, 6)
    text_line(doc, "Interface Design Blueprint", 15, False, "center", 0, 4)
    text_line(doc, "22 screens drawn with Microsoft Word shapes", 11, False, "center", 0, 24,
              (90, 90, 90))
    text_line(doc, "Each page shows one full screen mockup, numbered callouts and an "
                   "element note grid that covers input, process, popup message and output.",
              10, False, "center", 0, 28, (90, 90, 90))

    table = doc.add_table(rows=6, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    labels = ["Course / Class", "Group name", "Member 1", "Member 2", "Member 3",
              "Submission date"]
    for row, name in zip(table.rows, labels):
        row.cells[0].text = name
        row.cells[1].text = ""
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[0].paragraphs[0].runs[0].font.name = FONT
        shading(row.cells[0], "EFEFEF")
        for cell in row.cells:
            cell_margins(cell, 100, 130, 100, 130)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    text_line(doc, "Drawn only with rectangle, rounded rectangle, oval, line, arrow and "
                   "text box in Microsoft Word. No picture is inserted, so every part can "
                   "still be moved or edited.", 9, False, "center", 22, 0, (100, 100, 100))
    doc.add_page_break()


def index_page(doc):
    text_line(doc, "Interface Index", 16, True, "left", 0, 2)
    text_line(doc, "22 interfaces grouped by module", 10, False, "left", 0, 10, (100, 100, 100))
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    headers = ["Code", "Interface name", "Module", "Main popup message type"]
    kinds = {"ok": "Information", "warning": "Warning", "confirm": "Confirmation"}
    for index, value in enumerate(headers):
        cell = table.rows[0].cells[index]
        cell.text = value
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.name = FONT
        cell.paragraphs[0].runs[0].font.size = Pt(9.5)
        shading(cell, "EFEFEF")
    for number, screen in enumerate(SCREENS, 1):
        cells = table.add_row().cells
        values = (f"IF-{number:02d}", screen[0], screen[1], kinds[screen[7]])
        for index, value in enumerate(values):
            cells[index].text = value
            cells[index].paragraphs[0].runs[0].font.name = FONT
            cells[index].paragraphs[0].runs[0].font.size = Pt(9.5)
            cell_margins(cells[index], 30, 80, 30, 80)
    doc.add_page_break()


def guide_page(doc):
    text_line(doc, "How to Read Each Interface Page", 16, True, "left", 0, 2)
    text_line(doc, "The same five callout numbers are used on all 22 pages.", 10, False,
              "left", 0, 10, (100, 100, 100))
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    for index, value in enumerate(["Callout", "Meaning", "Where it sits on the page"]):
        cell = table.rows[0].cells[index]
        cell.text = value
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.name = FONT
        cell.paragraphs[0].runs[0].font.size = Pt(9.5)
        shading(cell, "EFEFEF")
    guide_rows = [
        ("1", "Input field or control that the student fills in or selects",
         "Upper part of the screen body, full width"),
        ("2", "Action button that submits the input to the system",
         "Beside or right under the input controls"),
        ("3", "Processing strip that states what the system does with the input",
         "Dashed strip at the bottom of the screen frame"),
        ("4", "Popup message shown after the process finishes or fails",
         "Dialog floating over the lower right of the screen"),
        ("5", "Output area that displays the result to the student",
         "Output panel at the lower left of the screen"),
        ("-", "Validation rule and next screen; these have no callout marker",
         "Last two rows of the element note grid"),
    ]
    for values in guide_rows:
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cells[index].text = value
            cells[index].paragraphs[0].runs[0].font.name = FONT
            cells[index].paragraphs[0].runs[0].font.size = Pt(9.5)
            cell_margins(cells[index], 40, 80, 40, 80)
    text_line(doc, "Layout of every interface page: a code band at the top, then one full "
                   "window mockup made of a title bar, a horizontal top navigation with six "
                   "menu items, a search box and an account initial. Inside the mockup, the "
                   "interactive controls fill the upper strip, the output panel sits at the "
                   "lower left, the popup dialog floats at the lower right, and a dashed "
                   "processing strip runs along the bottom. The element note grid under the "
                   "mockup explains all five callouts.", 9.5, False, "left", 14, 0,
              (90, 90, 90))
    doc.add_page_break()


# ------------------------------------------------------------------------ build
def build():
    doc = Document()
    setup_section(doc.sections[0])
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)

    cover_page(doc)
    index_page(doc)
    guide_page(doc)

    for number, screen in enumerate(SCREENS, 1):
        page = anchor_paragraph(doc)
        draw_page(page, number, screen)
        if number < len(SCREENS):
            doc.add_page_break()

    for section in doc.sections:
        setup_section(section)  # no footer or page-number field is added anywhere

    props = doc.core_properties
    props.title = "Student Learning Hub System - 22 Interface Blueprint"
    props.subject = "Full screen mockups with numbered callouts and element notes"
    props.author = "Student Learning Hub Group"
    props.keywords = "interface design, Word shapes, editable, input, process, output, popup"
    doc.save(OUTPUT)

    print(f"CREATED={OUTPUT.name}")
    print(f"INTERFACES={len(SCREENS)}")
    print(f"SHAPES={len(SHAPE_RECTS)}")
    print(f"LAYOUT_KINDS={len({s[2] for s in SCREENS})}")
    print(f"SIZE_BYTES={OUTPUT.stat().st_size}")


if __name__ == "__main__":
    build()
