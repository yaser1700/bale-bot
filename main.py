import os
import time
import threading
import tempfile
import requests

from flask import Flask
from openpyxl import load_workbook

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import arabic_reshaper
from bidi.algorithm import get_display


# =========================================================
# تنظیمات
# =========================================================

TOKEN = os.getenv("BALE_TOKEN")

BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}"

app = Flask(__name__)


# =========================================================
# فایل‌های Excel
# =========================================================

FILES = {

    "📦 سوکت عباسی":
        "لیست_قیمت_سوکت عباسی.xlsx",

    "🔌 کابل تکنو سبزوار":
        "لیست_قیمت_کابل تکنو سبزوار.xlsx",

    "⚡ وایر عباسی":
        "لیست_قیمت_وایر عباسی.xlsx",

    "🔩 مهره و سنسور":
        "لیست_قیمت_مهره و سنسور.xlsx",

    "💡 قطعات برقی خودرو":
        "لیست_قیمت_قطعات برقی خودرو.xlsx",

    "🧩 خارجات و پلیمرجات":
        "لیست_قیمت_خارجات و پلیمرجات.xlsx",

    "🔌 کابل خودرو سبزوار":
        "لیست_قیمت_کابل خودرو سبزوار.xlsx",

    "⚙️ شیلنگ خودرو":
        "لیست_قیمت_شیلنگ خودرو.xlsx",

    "⚙️ جلوبندی":
        "لیست_قیمت_جلوبندی.xlsx",
}


# =========================================================
# فونت PDF
# =========================================================

FONT_PATHS = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
]

FONT_PATH = None

for p in FONT_PATHS:
    if os.path.exists(p):
        FONT_PATH = p
        break

if FONT_PATH:
    pdfmetrics.registerFont(
        TTFont("PersianFont", FONT_PATH)
    )
else:
    print("WARNING: Persian font not found")


# =========================================================
# تبدیل فارسی برای PDF
# =========================================================

def rtl(text):

    if text is None:
        return ""

    text = str(text)

    try:
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    except Exception:
        return text


# =========================================================
# تمیز کردن مقدار Excel
# =========================================================

def clean(value):

    if value is None:
        return ""

    return str(value).strip()


# =========================================================
# خواندن Excel
# =========================================================

def read_excel(path):

    if not os.path.exists(path):
        return None, "❌ فایل قیمت در سرور پیدا نشد."

    try:

        workbook = load_workbook(
            path,
            read_only=True,
            data_only=True
        )

        sheets = []

        for sheet in workbook.worksheets:

            rows = []

            for row in sheet.iter_rows(values_only=True):

                values = [
                    clean(x)
                    for x in row
                ]

                if any(values):
                    rows.append(values)

            if rows:
                sheets.append(
                    (
                        sheet.title,
                        rows
                    )
                )

        workbook.close()

        if not sheets:
            return None, "❌ این فایل Excel خالی است."

        return sheets, None

    except Exception as e:

        print("Excel error:", e)

        return None, "❌ خطا در خواندن فایل Excel."


# =========================================================
# ساخت PDF از Excel
# =========================================================

def excel_to_pdf(excel_path, pdf_path, title):

    sheets, error = read_excel(excel_path)

    if error:
        return False, error

    try:

        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=landscape(A4),
            rightMargin=20,
            leftMargin=20,
            topMargin=20,
            bottomMargin=20
        )

        styles = getSampleStyleSheet()

        if FONT_PATH:

            title_style = ParagraphStyle(
                "PersianTitle",
                parent=styles["Title"],
                fontName="PersianFont",
                fontSize=18,
                leading=24,
                alignment=TA_CENTER,
                spaceAfter=15
            )

            cell_style = ParagraphStyle(
                "PersianCell",
                parent=styles["Normal"],
                fontName="PersianFont",
                fontSize=8,
                leading=11,
                alignment=TA_RIGHT
            )

        else:

            title_style = ParagraphStyle(
                "Title2",
                parent=styles["Title"],
                fontSize=18,
                alignment=TA_CENTER
            )

            cell_style = ParagraphStyle(
                "Cell2",
                parent=styles["Normal"],
                fontSize=8,
                alignment=TA_RIGHT
            )

        story = []

        # عنوان
        story.append(
            Paragraph(
                rtl("لیست قیمت - " + title),
                title_style
            )
        )

        story.append(Spacer(1, 10))

        for sheet_index, (sheet_name, rows) in enumerate(sheets):

            if sheet_index > 0:
                story.append(PageBreak())

            story.append(
                Paragraph(
                    rtl("برگه: " + sheet_name),
                    cell_style
                )
            )

            story.append(Spacer(1, 8))

            # بیشترین تعداد ستون
            max_cols = max(
                len(row)
                for row in rows
            )

            # محدود کردن تعداد ستون‌ها
            max_cols = min(max_cols, 12)

            table_data = []

            for row in rows:

                row = row[:max_cols]

                while len(row) < max_cols:
                    row.append("")

                converted = []

                for value in row:

                    converted.append(
                        Paragraph(
                            rtl(value),
                            cell_style
                        )
                    )

                table_data.append(converted)

            if not table_data:
                continue

            page_width = landscape(A4)[0] - 40

            col_width = page_width / max_cols

            table = Table(
                table_data,
                colWidths=[col_width] * max_cols,
                repeatRows=1
            )

            table.setStyle(
                TableStyle([

                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.5,
                        colors.grey
                    ),

                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.lightgrey
                    ),

                    (
                        "VALIGN",
                        (0, 0),
                        (-1, -1),
                        "MIDDLE"
                    ),

                    (
                        "ALIGN",
                        (0, 0),
                        (-1, -1),
                        "RIGHT"
                    ),

                    (
                        "LEFTPADDING",
                        (0, 0),
                        (-1, -1),
                        4
                    ),

                    (
                        "RIGHTPADDING",
                        (0, 0),
                        (-1, -1),
                        4
                    ),

                    (
                        "TOPPADDING",
                        (0, 0),
                        (-1, -1),
                        4
                    ),

                    (
                        "BOTTOMPADDING",
                        (0, 0),
                        (-1, -1),
                        4
                    ),

                ])
            )

            story.append(table)

        doc.build(story)

        return True, None

    except Exception as e:

        print("PDF error:", e)

        return False, "❌ خطا در ساخت فایل PDF."


# =========================================================
# ارسال پیام متنی
# =========================================================

def send_message(chat_id, text, keyboard=None):

    data = {
        "chat_id": chat_id,
        "text": text
    }

    if keyboard:

        data["reply_markup"] = {
            "keyboard": keyboard,
            "resize_keyboard": True
        }

    try:

        response = requests.post(
            f"{BASE_URL}/sendMessage",
            json=data,
            timeout=30
        )

        print(
            "sendMessage:",
            response.status_code,
            response.text[:500]
        )

        return response.json()

    except Exception as e:

        print("Send message error:", e)

        return None


# =========================================================
# ارسال PDF
# =========================================================

def send_document(chat_id, pdf_path, caption=""):

    try:

        with open(pdf_path, "rb") as file:

            files = {
                "document": (
                    os.path.basename(pdf_path),
                    file,
                    "application/pdf"
                )
            }

            data = {
                "chat_id": str(chat_id),
                "caption": caption
            }

            response = requests.post(
                f"{BASE_URL}/sendDocument",
                data=data,
                files=files,
                timeout=180
            )

        print(
            "sendDocument:",
            response.status_code,
            response.text[:1000]
        )

        result = response.json()

        return result

    except Exception as e:

        print("Document error:", e)

        return {
            "ok": False,
            "description": str(e)
        }


# =========================================================
# منوی اصلی
# =========================================================

def main_menu(chat_id):

    keyboard = [

        [
            "📄 دریافت لیست قیمت"
        ],

        [
            "📦 سوکت عباسی",
            "🔌 کابل تکنو سبزوار"
        ],

        [
            "⚡ وایر عباسی",
            "🔩 مهره و سنسور"
        ],

        [
            "💡 قطعات برقی خودرو",
            "🧩 خارجات و پلیمرجات"
        ],

        [
            "🔌 کابل خودرو سبزوار",
            "⚙️ شیلنگ خودرو"
        ],

        [
            "⚙️ جلوبندی"
        ]
    ]

    send_message(
        chat_id,
        "سلام 👋\n\n"
        "به ربات تولید و پخش قطعات خودرو عباسی خوش آمدید.\n\n"
        "لطفاً گروه مورد نظر را انتخاب کنید:",
        keyboard
    )


# =========================================================
# پردازش پیام
# =========================================================

def process_message(message):

    chat = message.get("chat") or {}

    chat_id = chat.get("id")

    text = clean(
        message.get("text")
    )

    if not chat_id:
        return

    print(
        "MESSAGE:",
        chat_id,
        repr(text)
    )

    # -----------------------------------------------------
    # شروع
    # -----------------------------------------------------

    if text == "/start":

        main_menu(chat_id)

        return

    # -----------------------------------------------------
    # منوی اصلی
    # -----------------------------------------------------

    if text in (
        "🔙 منوی اصلی",
        "منوی اصلی"
    ):

        main_menu(chat_id)

        return

    # -----------------------------------------------------
    # دریافت لیست قیمت
    # -----------------------------------------------------

    if text == "📄 دریافت لیست قیمت":

        main_menu(chat_id)

        return

    # -----------------------------------------------------
    # بررسی انتخاب گروه
    # -----------------------------------------------------

    filename = FILES.get(text)

    if filename:

        send_message(
            chat_id,
            "⏳ لطفاً کمی صبر کنید...\n\n"
            "در حال آماده‌سازی لیست قیمت PDF است."
        )

        # مسیر فایل Excel
        excel_path = os.path.join(
            os.path.dirname(__file__),
            filename
        )

        if not os.path.exists(excel_path):

            send_message(
                chat_id,
                "❌ فایل مربوط به این گروه در سرور پیدا نشد.\n\n"
                f"نام فایل مورد انتظار:\n{filename}"
            )

            return

        # فایل موقت PDF
        temp_dir = tempfile.gettempdir()

        safe_name = (
            text
            .replace("/", "_")
            .replace("\\", "_")
            .replace(" ", "_")
        )

        pdf_path = os.path.join(
            temp_dir,
            f"{safe_name}_لیست_قیمت.pdf"
        )

        # ساخت PDF
        success, error = excel_to_pdf(
            excel_path,
            pdf_path,
            text
        )

        if not success:

            send_message(
                chat_id,
                error
            )

            return

        # ارسال PDF
        result = send_document(
            chat_id,
            pdf_path,
            f"📄 لیست قیمت {text}"
        )

        # حذف فایل موقت
        try:

            if os.path.exists(pdf_path):
                os.remove(pdf_path)

        except Exception:
            pass

        if result and result.get("ok"):

            send_message(
                chat_id,
                "✅ لیست قیمت با موفقیت ارسال شد.\n\n"
                "برای انتخاب گروه دیگر /start را بفرستید."
            )

        else:

            description = ""

            if isinstance(result, dict):
                description = result.get(
                    "description",
                    ""
                )

            send_message(
                chat_id,
                "❌ ارسال PDF انجام نشد.\n\n"
                + description
            )

        return

    # -----------------------------------------------------
    # پیام ناشناخته
    # -----------------------------------------------------

    send_message(
        chat_id,
        "❓ گزینه مورد نظر پیدا نشد.\n\n"
        "برای نمایش منو /start را بفرستید."
    )


# =========================================================
# دریافت آپدیت‌های بله
# =========================================================

def bot_loop():

    offset = 0

    print("Bale bot started...")

    while True:

        try:

            response = requests.get(
                f"{BASE_URL}/getUpdates",
                params={
                    "offset": offset,
                    "timeout": 30
                },
                timeout=45
            )

            data = response.json()

            if not data.get("ok"):

                print(
                    "API error:",
                    data
                )

                time.sleep(5)

                continue

            updates = data.get(
                "result",
                []
            )

            for update in updates:

                offset = (
                    update.get(
                        "update_id",
                        offset
                    ) + 1
                )

                message = update.get(
                    "message"
                )

                if message:

                    try:

                        process_message(
                            message
                        )

                    except Exception as e:

                        print(
                            "Message error:",
                            e
                        )

        except Exception as e:

            print(
                "Bot error:",
                e
            )

            time.sleep(5)


# =========================================================
# صفحه اصلی Render
# =========================================================

@app.route("/")
def home():

    return "Bale bot is running!"


# =========================================================
# اجرای برنامه
# =========================================================

if __name__ == "__main__":

    threading.Thread(
        target=bot_loop,
        daemon=True
    ).start()

    port = int(
        os.environ.get(
            "PORT",
            10000
        )
    )

    app.run(
        host="0.0.0.0",
        port=port
    )
