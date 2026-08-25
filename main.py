import os
import time
import threading
import requests
from flask import Flask
from openpyxl import load_workbook


# =========================
# تنظیمات ربات
# =========================

TOKEN = os.getenv("BALE_TOKEN")

if not TOKEN:
    print("ERROR: BALE_TOKEN is not set!")

BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}"

app = Flask(__name__)


# =========================
# فایل‌های قیمت
# =========================

FILES = {
    "سوکت عباسی 📦": "لیست_قیمت_سوکت عباسی.xlsx",
    "کابل تکنو سبزوار 🔌": "لیست_قیمت_کابل تکنو سبزوار.xlsx",
    "وایر عباسی ⚡": "لیست_قیمت_وایر عباسی.xlsx",
    "مهره و سنسور 🔩": "لیست_قیمت_مهره و سنسور.xlsx",
    "قطعات برقی خودرو 💡": "لیست_قیمت_قطعات برقی خودرو.xlsx",
    "خارجات و پلیمرجات 🧩": "لیست_قیمت_خارجات و پلیمرجات.xlsx",
    "کابل خودرو سبزوار 🔌": "لیست_قیمت_کابل خودرو سبزوار.xlsx",
    "شیلنگ خودرو ⚙️": "لیست_قیمت_شیلنگ خودرو.xlsx",
    "جلوبندی ⚙️": "لیست_قیمت_جلوبندی.xlsx"
}


# =========================
# ارسال پیام
# =========================

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

        print("Send:", response.status_code, response.text)

    except Exception as e:
        print("Send error:", e)


# =========================
# منوی اصلی
# =========================

def main_menu(chat_id):

    keyboard = [
        ["📄 دریافت لیست قیمت"],
        ["📦 سوکت عباسی", "🔌 کابل تکنو سبزوار"],
        ["⚡ وایر عباسی", "🔩 مهره و سنسور"],
        ["💡 قطعات برقی خودرو", "🧩 خارجات و پلیمرجات"],
        ["🔌 کابل خودرو سبزوار", "⚙️ شیلنگ خودرو"],
        ["⚙️ جلوبندی"]
    ]

    send_message(
        chat_id,
        "لطفاً گروه مورد نظر را انتخاب کنید:",
        keyboard
    )


# =========================
# تمیز کردن متن
# =========================

def clean(value):

    if value is None:
        return ""

    text = str(value).strip()

    # یکسان‌سازی فاصله‌ها
    text = " ".join(text.split())

    return text


# =========================
# پیدا کردن فایل
# =========================

def find_filename(text):

    text = clean(text)

    # اول تطبیق دقیق
    for button_name, filename in FILES.items():

        if clean(button_name) == text:
            return filename

    # اگر ایموجی یا فاصله متفاوت بود
    for button_name, filename in FILES.items():

        button_without_emoji = (
            button_name
            .replace("📦", "")
            .replace("🔌", "")
            .replace("⚡", "")
            .replace("🔩", "")
            .replace("💡", "")
            .replace("🧩", "")
            .replace("⚙️", "")
            .replace("⚙", "")
            .strip()
        )

        text_without_emoji = (
            text
            .replace("📦", "")
            .replace("🔌", "")
            .replace("⚡", "")
            .replace("🔩", "")
            .replace("💡", "")
            .replace("🧩", "")
            .replace("⚙️", "")
            .replace("⚙", "")
            .strip()
        )

        if button_without_emoji == text_without_emoji:
            return filename

    return None


# =========================
# خواندن فایل Excel
# =========================

def read_excel(path):

    if not os.path.exists(path):

        return (
            "❌ فایل قیمت در سرور پیدا نشد.\n\n"
            f"نام فایل مورد انتظار:\n{os.path.basename(path)}"
        )

    try:

        workbook = load_workbook(
            path,
            read_only=True,
            data_only=True
        )

        lines = []

        for sheet in workbook.worksheets:

            lines.append(
                f"📋 {sheet.title}"
            )

            for row in sheet.iter_rows(values_only=True):

                values = [
                    clean(x)
                    for x in row
                ]

                if not any(values):
                    continue

                lines.append(
                    " | ".join(values)
                )

            lines.append("")

        workbook.close()

        if not lines:

            return "❌ این فایل خالی است."

        return "\n".join(lines)

    except Exception as e:

        print("Excel error:", e)

        return (
            "❌ خطا در خواندن فایل قیمت.\n"
            "لطفاً فایل Excel را بررسی کنید."
        )


# =========================
# ارسال متن‌های طولانی
# =========================

def send_long(chat_id, text):

    max_length = 3500

    if len(text) <= max_length:

        send_message(
            chat_id,
            text
        )

        return

    while text:

        part = text[:max_length]

        if len(text) > max_length:

            position = part.rfind("\n")

            if position > 500:
                part = part[:position]

        send_message(
            chat_id,
            part
        )

        text = text[len(part):]

        time.sleep(0.3)


# =========================
# پردازش پیام
# =========================

def process_message(message):

    chat = message.get("chat") or {}

    chat_id = chat.get("id")

    text = clean(
        message.get("text")
    )

    if not chat_id:
        return


    # =====================
    # شروع
    # =====================

    if text == "/start":

        send_message(
            chat_id,
            "سلام 👋\n\n"
            "به گروه تولیدی و بازرگانی عباسی خوش آمدید.\n\n"
            "برای دریافت اطلاعات محصولات، "
            "گروه مورد نظر را انتخاب کنید."
        )

        time.sleep(0.3)

        main_menu(chat_id)

        return


    # =====================
    # منوی اصلی
    # =====================

    if text in (
        "🔙 منوی اصلی",
        "منوی اصلی",
        "🏠 منوی اصلی"
    ):

        main_menu(chat_id)

        return


    # =====================
    # دریافت لیست قیمت
    # =====================

    if text in (
        "📄 دریافت لیست قیمت",
        "دریافت لیست قیمت"
    ):

        send_message(
            chat_id,
            "📋 لطفاً گروه مورد نظر را انتخاب کنید:",
        )

        main_menu(chat_id)

        return


    # =====================
    # پیدا کردن فایل
    # =====================

    filename = find_filename(text)


    if filename:

        send_message(
            chat_id,
            "⏳ در حال خواندن لیست قیمت..."
        )

        path = os.path.join(
            os.path.dirname(__file__),
            filename
        )

        print("Reading file:")
        print(path)

        content = read_excel(path)

        send_long(
            chat_id,
            content
        )

        send_message(
            chat_id,
            "✅ پایان لیست قیمت\n\n"
            "برای انتخاب گروه دیگر، /start را بفرستید."
        )

        return


    # =====================
    # پیام ناشناخته
    # =====================

    send_message(
        chat_id,
        "❗ لطفاً یکی از گزینه‌های منو را انتخاب کنید."
    )

    main_menu(chat_id)


# =========================
# دریافت پیام‌های بله
# =========================

def bot_loop():

    offset = 0

    print("================================")
    print("Bale bot started...")
    print("================================")

    while True:

        try:

            response = requests.get(
                f"{BASE_URL}/getUpdates",
                params={
                    "offset": offset,
                    "timeout": 30
                },
                timeout=40
            )

            data = response.json()

            if not data.get("ok"):

                print(
                    "API error:",
                    data
                )

                time.sleep(5)

                continue


            for update in data.get(
                "result",
                []
            ):

                offset = (
                    update.get(
                        "update_id",
                        offset
                    )
                    + 1
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


# =========================
# صفحه اصلی Render
# =========================

@app.route("/")
def home():

    return "Bale bot is running!"


# =========================
# اجرای برنامه
# =========================

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
