# -*- coding: utf-8 -*-
import openpyxl
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# ---------- رنگ‌ها (هرجا از string استفاده می‌کنیم نه object) ----------
DARK_GREEN = "1F5E24"
MEDIUM_GREEN = "2E7D32"
LIGHT_YELLOW = "FFFFF9C4"
LIGHT_GRAY = "F5F5F5"
GRAY_BORDER = "9E9E9E"
INPUT_BORDER = "A18CD1"
ALTERNATE = "E8F5E9"

# ---------- فونت‌ها ----------
TITLE_FONT = {"name": "B Nazanin", "size": 16, "bold": True, "color": "000000"}
HEADER_FONT = {"name": "B Nazanin", "size": 11, "bold": True, "color": "FFFFFF"}
LABEL_FONT = {"name": "B Nazanin", "size": 10, "bold": True, "color": DARK_GREEN}
INPUT_FONT = {"name": "B Nazanin", "size": 10, "color": "000000"}
BODY_FONT = {"name": "B Nazanin", "size": 9, "color": "000000"}
GREEN_FONT = {"name": "B Nazanin", "size": 11, "bold": True, "color": DARK_GREEN}
RED_FONT = {"name": "B Nazanin", "size": 11, "bold": True, "color": "C00000"}

TOMAN = "#,##0"
DATE_FMT = "yyyy/mm/dd"

wb = openpyxl.Workbook()

def set_font(cell, props):
    cell.font = Font(**props)

def make_border(color):
    return Border(
        left=Side("thin", color=color),
        right=Side("thin", color=color),
        top=Side("thin", color=color),
        bottom=Side("thin", color=color),
    )

def hcell(ws, col, text, size=11, bg=DARK_GREEN, bold=True, color="FFFFFF", align="center", wrap=False):
    c = ws.cell(row=1, column=col, value=text)
    set_font(c, HEADER_FONT)
    c.fill = PatternFill("solid", fgColor=bg)
    c.border = make_border(GRAY_BORDER)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)

def set_style(cell, text=None, fill="FFFFFF", font=None, border=True, align="left", num_format=None):
    if text is not None:
        cell.value = text
    if font is not None:
        set_font(cell, font)
    if fill != "":
        cell.fill = PatternFill("solid", fgColor=fill)
    if border:
        cell.border = make_border(GRAY_BORDER)
    if align:
        cell.alignment = Alignment(horizontal=align, vertical="center")
    if num_format:
        cell.number_format = num_format

def title(ws, coord, text):
    cell = ws[coord]
    cell.value = text
    set_font(cell, TITLE_FONT)
    cell.fill = PatternFill("solid", fgColor=DARK_GREEN)
    cell.border = make_border(GRAY_BORDER)
    cell.alignment = Alignment(horizontal="center", vertical="center")

def section(ws, coord, text):
    cell = ws[coord]
    cell.value = text
    set_font(cell, {"name": "B Nazanin", "size": 12, "bold": True, "color": "FFFFFF"})
    cell.fill = PatternFill("solid", fgColor=MEDIUM_GREEN)
    cell.border = make_border(GRAY_BORDER)
    cell.alignment = Alignment(horizontal="left", vertical="center")

def input_row(ws, row, cells, formats=None, sizes=None):
    formats = formats or []
    sizes = sizes or []
    for j, val in enumerate(cells):
        col = j + 1
        c = ws.cell(row=row, column=col, value=val)
        fmt = formats[j] if j < len(formats) else ""
        is_num = isinstance(val, (int, float))
        is_formula = isinstance(val, str) and val.startswith("=")
        if is_formula:
            # سلول فرمول - رنگ سبز
            set_style(c, font=GREEN_FONT, fill=ALTERNATE, border=True, align="left", num_format=TOMAN)
        elif fmt == "date" or fmt == "textgen" or fmt == "text":
            # متن / تاریخ شمسی به صورت متن - بدون فرمت تاریخ تا اکسل تعمیر نکند
            set_style(c, font=INPUT_FONT, fill=LIGHT_YELLOW, border=True,
                      align="center" if fmt == "date" else "left", num_format=None)
        elif fmt == "money" or (is_num and fmt != "plain"):
            # مبلغ پولی
            set_style(c, font=INPUT_FONT, fill=LIGHT_YELLOW, border=True, align="right", num_format=TOMAN)
        else:
            # عدد ساده (ردیف، تعداد)
            set_style(c, font=INPUT_FONT, fill=LIGHT_YELLOW, border=True, align="center", num_format=None)

def data_row(ws, row, values, formulas, formats, colors):
    for j, val in enumerate(values):
        col = j + 1
        c = ws.cell(row=row, column=col, value=val)
        set_style(c, font=LABEL_FONT, fill="FFFFFF", border=True, align="right")
    for j, formula in enumerate(formulas):
        col = j + 1
        if formula:
            c = ws.cell(row=row, column=col, value=formula)
            set_style(c, font=GREEN_FONT if colors[j] == "green" else (RED_FONT if colors[j] == "red" else BODY_FONT),
                      fill=ALTERNATE if (j % 2 == 1) else "FFFFFF", border=True, align="right", num_format=formats[j])

def sheet(title, zoom=100):
    if title == "داشبورد" and "Sheet" in wb.sheetnames:
        ws = wb["Sheet"]
        ws.title = title
    else:
        ws = wb.create_sheet(title)
    ws.sheet_view.zoomScale = zoom
    ws.sheet_view.showGridLines = True
    return ws

def col_width(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

# ============================================================
# SHEET 1: داشبورد
# ============================================================
ws = sheet("داشبورد", 120)
col_width(ws, {"A": 52, "B": 18, "C": 45, "D": 12})

title(ws, "A1", "تراوت مارکت - داشبورد مالی")
ws["C1"] = "امروز:"
ws["C1"].font = Font(name="B Nazanin", size=9, bold=True, color="FFFFFF")
ws["C1"].fill = PatternFill("solid", fgColor=DARK_GREEN)
ws["C1"].border = make_border(GRAY_BORDER)
ws["D1"] = "=TEXT(TODAY(),\"yyyy/mm/dd\")"
ws["D1"].font = Font(name="B Nazanin", size=9, color="FFFFFF")
ws["D1"].fill = PatternFill("solid", fgColor=DARK_GREEN)
ws["D1"].alignment = Alignment(horizontal="center", vertical="center")
ws["D1"].border = make_border(GRAY_BORDER)

# عنوان از A1:B1 را برای جدول جابجا می‌کنیم
ws["C2"] = "واحد: تومان"
ws["C2"].font = Font(name="B Nazanin", size=9, italic=True, color="666666")
ws["C2"].border = make_border(GRAY_BORDER)

# هدر جدول در row 3
for col, txt in [(1, "شاخص"), (2, "مقدار")]:
    cell = ws.cell(row=3, column=col, value=txt)
    cell.font = Font(name="B Nazanin", size=11, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=MEDIUM_GREEN)
    cell.border = make_border(GRAY_BORDER)
    cell.alignment = Alignment(horizontal="center", vertical="center")

rows = [
    ("مجموع فروش کل (تومان)", "=SUMPRODUCT(--('ورودی‌ها'!$D:$D=\"فروش\"),'ورودی‌ها'!$E:$E)", "green"),
    ("مجموع هزینه‌ها (تومان)", "=SUMPRODUCT(--('ورودی‌ها'!$D:$D=\"هزینه\"),'ورودی‌ها'!$E:$E)", "red"),
    ("سود خالص (تومان)", "=B4-B5", "green"),
    ("موجودی نقد بانک (تومان)", "=SUM('حساب‌ها'!$D:$D)", "green"),
    ("کل مطالبات (تومان)", "=SUMIF('بدهی_مطالبات'!$B:$B,\"مطالبات\",'بدهی_مطالبات'!$E:$E)", "green"),
    ("کل بدهی به تأمین‌کنندگان (تومان)", "=SUMIF('بدهی_مطالبات'!$B:$B,\"بدهی\",'بدهی_مطالبات'!$E:$E)", "red"),
    ("تعداد سفارشات ثبت‌شده", "=COUNTA('سفارشات'!$A:$A)-1", "green"),
    ("تعداد کالاها", "=COUNTA('محصولات'!$A:$A)-1", "green"),
]
r = 4
for label, formula, color in rows:
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=1).font = Font(name="B Nazanin", size=10, bold=True, color=DARK_GREEN)
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor="FFFFFF")
    ws.cell(row=r, column=1).border = make_border(GRAY_BORDER)
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="right", vertical="center")
    c2 = ws.cell(row=r, column=2, value=formula)
    c2.font = Font(name="B Nazanin", size=10, bold=True,
                   color=DARK_GREEN if color == "green" else "C00000")
    c2.fill = PatternFill("solid", fgColor=ALTERNATE if r % 2 == 0 else "FFFFFF")
    c2.border = make_border(GRAY_BORDER)
    c2.alignment = Alignment(horizontal="right", vertical="center")
    c2.number_format = TOMAN
    r += 1

# وضعیت مالی
ws.cell(row=13, column=1, value="وضعیت مالی کل")
ws.cell(row=13, column=1).font = Font(name="B Nazanin", size=11, bold=True, color=DARK_GREEN)
ws.cell(row=13, column=1).fill = PatternFill("solid", fgColor="DDEBF7")
ws.cell(row=13, column=1).border = make_border(GRAY_BORDER)
ws.cell(row=13, column=1).alignment = Alignment(horizontal="right", vertical="center")
c = ws.cell(row=13, column=2, value='=IF(B6>0,"✅ سودده","❌ زیان‌ده")')
c.font = Font(name="B Nazanin", size=12, bold=True, color=DARK_GREEN)
c.fill = PatternFill("solid", fgColor="DDEBF7")
c.border = make_border(GRAY_BORDER)
c.alignment = Alignment(horizontal="center", vertical="center")

ws["A16"] = "راهنما: برای گزارش ماهانه، تمام تراکنش‌ها را در تب \"ورودی‌ها\" ثبت کنید."
ws["A16"].font = Font(name="B Nazanin", size=9, italic=True, color="666666")

# ============================================================
# SHEET 2: محصولات
# ============================================================
ws = sheet("محصولات", 110)
col_width(ws, {"A": 7, "B": 12, "C": 28, "D": 16, "E": 18, "F": 12})
for i, t in enumerate(["ردیف", "کد کالا", "نام محصول", "تعداد در انبار", "تاریخ ورود", "وضعیت"]):
    hcell(ws, i+1, t)
input_row(ws, 2, [1, "K001", "قهوه اسپرسو", 50, "1405/01/10", "فعال"],
          formats=["text", "text", "text", "text", "date", "text"], sizes=[9]*6)
input_row(ws, 3, [2, "K002", "چای سیاه", 30, "1405/01/12", "فعال"],
          formats=["text", "text", "text", "text", "date", "text"], sizes=[9]*6)

# ============================================================
# SHEET 3: تأمین‌کنندگان
# ============================================================
ws = sheet("تأمین‌کنندگان", 110)
col_width(ws, {"A": 7, "B": 32, "C": 18, "D": 14})
for i, t in enumerate(["ردیف", "نام تأمین‌کننده", "شماره تماس", "وضعیت"]):
    hcell(ws, i+1, t)
input_row(ws, 2, [1, "تأمین‌کننده قهوه البرز", "09123456789", "فعال"],
          formats=["text", "text", "text", "text"])

# ============================================================
# SHEET 4: سفارشات
# ============================================================
ws = sheet("سفارشات", 100)
col_width(ws, {"A": 7, "B": 14, "C": 12, "D": 26, "E": 12, "F": 16, "G": 16})
for i, t in enumerate(["ردیف", "تاریخ", "کد کالا", "نام محصول", "مقدار فروش", "قیمت واحد (تومان)", "درآمد (تومان)"]):
    hcell(ws, i+1, t)
input_row(ws, 2, [1, "1405/02/01", "K001", "قهوه اسپرسو", 10, 120000, "=E2*F2"],
          formats=["text", "date", "text", "text", "text", "text", "formula"], sizes=[9]*7)
input_row(ws, 3, [2, "1405/02/02", "K002", "چای سیاه", 5, 45000, "=E3*F3"],
          formats=["text", "date", "text", "text", "text", "text", "formula"], sizes=[9]*7)

# ============================================================
# SHEET 5: دریافت و پرداخت
# ============================================================
ws = sheet("دریافت_و_پرداخت", 100)
col_width(ws, {"A": 7, "B": 14, "C": 30, "D": 14, "E": 18, "F": 18, "G": 18})
for i, t in enumerate(["ردیف", "تاریخ", "شرح", "نوع", "دریافتی (تومان)", "پرداختی (تومان)", "مانده (تومان)"]):
    hcell(ws, i+1, t)
input_row(ws, 2, [1, "1405/02/01", "دریافت وجه فروش", "دریافتی", 1200000, "", "=E2"],
          formats=["text", "date", "text", "text", "text", "text", "formula"], sizes=[9]*7)
input_row(ws, 3, [2, "1405/02/03", "پرداخت اجاره مغازه", "پرداختی", "", 3000000, "=-F3"],
          formats=["text", "date", "text", "text", "text", "text", "formula"], sizes=[9]*7)

# ============================================================
# SHEET 6: کالای امانی
# ============================================================
ws = sheet("کالای_امانی", 100)
col_width(ws, {"A": 7, "B": 15, "C": 24, "D": 32, "E": 14, "F": 15, "G": 16})
for i, t in enumerate(["ردیف", "تاریخ ورود", "نام کالا", "شرح", "مقدار ورود", "تاریخ خروج", "وضعیت"]):
    hcell(ws, i+1, t)
input_row(ws, 2, [1, "1405/01/05", "کالای سوپرمارکتی الف", "محصولات امانی برای فروشگاه دیگر", 20, "", '=IF(F2="","در انبار","خروج داده شده")'],
          formats=["text", "date", "text", "text", "text", "date", "formula"], sizes=[9]*7)
input_row(ws, 3, [2, "1405/01/20", "لوازم خانگی ب", "کالای امانی برای تعمیر", 5, "1405/02/10", '=IF(F3="","در انبار","خروج داده شده")'],
          formats=["text", "date", "text", "text", "text", "date", "formula"], sizes=[9]*7)

# ============================================================
# SHEET 7: بدهی و مطالبات
# ============================================================
ws = sheet("بدهی_مطالبات", 100)
col_width(ws, {"A": 7, "B": 10, "C": 30, "D": 14, "E": 18, "F": 20, "G": 18})
for i, t in enumerate(["ردیف", "نوع", "شرح/نام طرف", "تاریخ", "مبلغ (تومان)", "پرداخت/دریافت (تومان)", "باقیمانده (تومان)"]):
    hcell(ws, i+1, t)
input_row(ws, 2, [1, "بدهی", "تأمین‌کننده قهوه - خرید امانی", "1405/01/10", 1500000, 0, "=E2-F2"],
          formats=["text", "text", "text", "date", "text", "text", "formula"], sizes=[9]*7)
input_row(ws, 3, [2, "مطالبات", "عمده‌فروشی - حساب باز", "1405/02/01", 850000, 0, "=E3-F3"],
          formats=["text", "text", "text", "date", "text", "text", "formula"], sizes=[9]*7)

# ============================================================
# SHEET 8: کمیسیون تأمین‌کننده
# ============================================================
ws = sheet("کمیسیون_تأمین‌کننده", 100)
col_width(ws, {"A": 7, "B": 14, "C": 12, "D": 24, "E": 12, "F": 16, "G": 18, "H": 22})
for i, t in enumerate(["ردیف", "تاریخ", "کد کالا", "نام محصول", "مقدار فروش", "قیمت واحد (تومان)", "درآمد کل (تومان)", "سهم تأمین‌کننده 15% (تومان)"]):
    hcell(ws, i+1, t)
input_row(ws, 2, [1, "1405/02/01", "K001", "قهوه اسپرسو", 10, 120000, "=E2*F2", "=G2*0.15"],
          formats=["text", "date", "text", "text", "text", "text", "formula", "formula"], sizes=[9]*8)
input_row(ws, 3, [2, "1405/02/02", "K002", "چای سیاه", 5, 45000, "=E3*F3", "=G3*0.15"],
          formats=["text", "date", "text", "text", "text", "text", "formula", "formula"], sizes=[9]*8)

# ============================================================
# SHEET 9: حساب‌ها
# ============================================================
ws = sheet("حساب‌ها", 110)
col_width(ws, {"A": 30, "B": 16, "C": 18, "D": 18})
for i, t in enumerate(["شرح حساب", "نوع", "موجودی اولیه (تومان)", "موجودی فعلی (تومان)"]):
    hcell(ws, i+1, t)
input_row(ws, 2, ["بانک", "دارایی", 50000000, "=C2+SUM('دریافت_و_پرداخت'!E:E)-SUM('دریافت_و_پرداخت'!F:F)"],
          formats=["text", "text", "text", "formula"])
input_row(ws, 3, ["صندوق", "دارایی", 10000000, "=C3+SUM('دریافت_و_پرداخت'!E:E)-SUM('دریافت_و_پرداخت'!F:F)"],
          formats=["text", "text", "text", "formula"])

# ============================================================
# SHEET 10: خلاصه
# ============================================================
ws = sheet("خلاصه", 110)
col_width(ws, {"A": 55, "B": 22})
title(ws, "A1", "خلاصه وضعیت مالی تراوت مارکت")
hcell(ws, 1, "شاخص", bg=MEDIUM_GREEN)
hcell(ws, 2, "مقدار (تومان)", bg=MEDIUM_GREEN)
sum_rows = [
    ("کل فروش", "=SUMPRODUCT(--('ورودی‌ها'!$D:$D=\"فروش\"),'ورودی‌ها'!$E:$E)"),
    ("کل هزینه", "=SUMPRODUCT(--('ورودی‌ها'!$D:$D=\"هزینه\"),'ورودی‌ها'!$E:$E)"),
    ("سود ناخالص", "=B4-B5"),
    ("مجموع دریافتی", "=SUM('دریافت_و_پرداخت'!E:E)"),
    ("مجموع پرداختی", "=SUM('دریافت_و_پرداخت'!F:F)"),
    ("موجودی نقد نهایی", "=B7-B8"),
    ("کل مطالبات", "=SUMIF('بدهی_مطالبات'!$B:$B,\"مطالبات\",'بدهی_مطالبات'!$E:$E)"),
    ("کل بدهی تأمین‌کنندگان", "=SUMIF('بدهی_مطالبات'!$B:$B,\"بدهی\",'بدهی_مطالبات'!$E:$E)"),
]
for i, (label, formula) in enumerate(sum_rows):
    r = 4 + i
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=1).font = Font(name="B Nazanin", size=10, bold=True, color=DARK_GREEN)
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=ALTERNATE if i % 2 == 0 else "FFFFFF")
    ws.cell(row=r, column=1).border = make_border(GRAY_BORDER)
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="right", vertical="center")
    c2 = ws.cell(row=r, column=2, value=formula)
    c2.font = Font(name="B Nazanin", size=10, bold=True, color=DARK_GREEN)
    c2.fill = PatternFill("solid", fgColor=ALTERNATE if i % 2 == 0 else "FFFFFF")
    c2.border = make_border(GRAY_BORDER)
    c2.alignment = Alignment(horizontal="right", vertical="center")
    c2.number_format = TOMAN

# ============================================================
# SHEET 11: ورودی‌ها
# ============================================================
ws = sheet("ورودی‌ها", 100)
col_width(ws, {"A": 7, "B": 14, "C": 36, "D": 18, "E": 18})
for i, t in enumerate(["ردیف", "تاریخ", "شرح", "نوع (فروش/هزینه)", "مبلغ (تومان)"]):
    hcell(ws, i+1, t)
input_row(ws, 2, [1, "1405/02/01", "فروش روزانه - فروشگاه اصلی", "فروش", 5000000],
          formats=["text", "date", "text", "text", "text"])
input_row(ws, 3, [2, "1405/02/02", "خرید مجدد قهوه", "هزینه", 1500000],
          formats=["text", "date", "text", "text", "text"])
input_row(ws, 4, [3, "1405/02/05", "قبض برق و آب", "هزینه", 450000],
          formats=["text", "date", "text", "text", "text"])

# ============================================================
# SHEET 12: تنظیمات
# ============================================================
ws = sheet("تنظیمات", 110)
col_width(ws, {"A": 40, "B": 40})
title(ws, "A1", "تنظیمات فایل")
for i, (label, value) in enumerate([
    ("نرخ کمیسیون پیش‌فرض (درصد)", 15),
    ("واحد پول", "تومان"),
    ("نام فروشگاه", "تراوت مارکت"),
    ("سال مالی", "1405"),
]):
    r = 3 + i
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=1).font = Font(name="B Nazanin", size=10, bold=True, color=DARK_GREEN)
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor="FFFFFF")
    ws.cell(row=r, column=1).border = make_border(GRAY_BORDER)
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=r, column=2, value=value)
    ws.cell(row=r, column=2).font = Font(name="B Nazanin", size=10, color="000000")
    ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=LIGHT_YELLOW)
    ws.cell(row=r, column=2).border = make_border(INPUT_BORDER)
    ws.cell(row=r, column=2).alignment = Alignment(horizontal="left", vertical="center")

# ============================================================
# SHEET 13: راهنما
# ============================================================
ws = sheet("راهنما", 100)
col_width(ws, {"A": 120})
title(ws, "A1", "راهنمای استفاده از فایل حسابداری تراوت مارکت")

help_rows = [
    ("نحوه کار با فایل", "1. این فایل اکسل برای مدیریت مالی تراوت مارکت طراحی شده است.\n2. واحد پول: تومان. تاریخ‌ها به صورت شمسی وارد کنید (مثلاً 1405/02/06).\n3. نیازی به دانش حسابداری تخصصی ندارید - کافی است داده‌ها را در جای مناسب وارد کنید.\n4. همه محاسبات (سود، کمیسیون، مانده) به صورت خودکار انجام می‌شود.\n5. برای گزارش ماهانه، کل داده‌های ماه را ثبت کنید و به داشبورد مراجعه کنید."),
    ("توضیح تب‌ها", "- داشبورد: نمای کلی وضعیت مالی\n- محصولات: لیست محصولات با کد اختصاصی\n- تأمین‌کنندگان: لیست تأمین‌کنندگان\n- سفارشات: ثبت سفارشات و درآمد\n- دریافت_و_پرداخت: تراکنش‌های نقدی\n- کالای_امانی: کالاهای امانی با تاریخ\n- بدهی_مطالبات: بدهی و مطالبات\n- کمیسیون_تأمین‌کننده: محاسبه 15%\n- حساب‌ها: موجودی بانکی و صندوق\n- خلاصه: گزارش وضعیت مالی\n- ورودی‌ها: ثبت روزانه تراکنش‌ها\n- تنظیمات: اطلاعات پایه\n- راهنما: این فایل"),
    ("نحوه ثبت سفارش", "به تب سفارشات بروید: تاریخ، کد کالا، نام محصول، مقدار، قیمت واحد را وارد کنید - درآمد محاسبه می‌شود."),
    ("نحوه ثبت کالای امانی", "به تب کالای_امانی بروید: تاریخ ورود، نام کالا، شرح، مقدار را وارد کنید. هنگام خروج، تاریخ خروج را بنویسید."),
    ("نحوه ثبت بدهی/مطالبات", "به تب بدهی_مطالبات بروید: نوع، شرح، تاریخ، مبلغ را وارد کنید. هنگام پرداخت، مبلغ پرداختی را بنویسید - باقیمانده محاسبه می‌شود."),
    ("نحوه ثبت تراکنش روزانه", "به تب ورودی‌ها بروید: تاریخ، شرح، نوع (فروش/هزینه)، مبلغ را وارد کنید - همه در داشبورد استفاده می‌شود."),
    ("نحوه گزارش ماهانه", "کل فعالیت‌های ماه را ثبت کنید. به داشبورد یا خلاصه بروید - اعداد خودکار محاسبه می‌شوند."),
]

r = 3
for title, text in help_rows:
    cell = ws.cell(row=r, column=1, value=title)
    set_font(cell, {"name": "B Nazanin", "size": 12, "bold": True, "color": "FFFFFF"})
    cell.fill = PatternFill("solid", fgColor=MEDIUM_GREEN)
    cell.border = make_border(GRAY_BORDER)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    r += 1
    for line in text.split("\n"):
        cell = ws.cell(row=r, column=1, value=line)
        set_font(cell, {"name": "B Nazanin", "size": 10, "color": "000000"})
        cell.fill = PatternFill("solid", fgColor="FFFFFF")
        cell.border = make_border(GRAY_BORDER)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        r += 1
    r += 1

# ============================================================
# ذخیره
# ============================================================
wb.save("Tavat_Market_Finance.xlsx")
print("فایل با موفقیت ساخته شد!")
print("تعداد تب‌ها:", len(wb.sheetnames))
for s in wb.sheetnames:
    print(" -", s)
