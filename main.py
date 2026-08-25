import os
import time
import threading
import requests
from flask import Flask
from openpyxl import load_workbook

TOKEN = os.getenv("BALE_TOKEN")
BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}" if TOKEN else ""

app = Flask(__name__)

FILES = {
    "📦 سوکت عباسی": "لیست_قیمت_سوکت عباسی.xlsx",
    "🔌 کابل تکنو سبزوار": "لیست_قیمت_کابل تکنو سبزوار.xlsx",
    "⚡ وایر عباسی": "لیست_قیمت_وایر عباسی.xlsx",
    "🔩 مهره و سنسور": "لیست_قیمت_مهره و سنسور.xlsx",
    "💡 قطعات برقی خودرو": "لیست_قیمت_قطعات برقی خودرو.xlsx",
    "🧩 خارجات و پلیمریجات": "لیست_قیمت_خارجات و پلیمریجات.xlsx",
    "🔌 کابل خودرو سبزوار": "لیست_قیمت_کابل خودرو سبزوار.xlsx",
    "⚙️ شیلنگ خودرو": "لیست_قیمت_شیلنگ خودرو.xlsx",
    "⚙️ جلوبندی": "لیست_قیمت_جلوبندی.xlsx",
}


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
        requests.post(
            f"{BASE_URL}/sendMessage",
            json=data,
            timeout=20
        )
    except Exception as e:
        print("Send error:", e)


def main_menu(chat_id):
    keyboard = [
        ["📄 دریافت لیست قیمت"],
        ["📦 سوکت عباسی", "🔌 کابل تکنو سبزوار"],
        ["⚡ وایر عباسی", "🔩 مهره و سنسور"],
        ["💡 قطعات برقی خودرو", "🧩 خارجات و پلیمریجات"],
        ["🔌 کابل خودرو سبزوار", "⚙️ شیلنگ خودرو"],
        ["⚙️ جلوبندی"]
    ]

    send_message(
        chat_id,
        "لطفاً گروه مورد نظر را انتخاب کنید:",
        keyboard
    )


def clean(value):
    if value is None:
        return ""
    return str(value).strip()


def read_excel(path):
    if not os.path.exists(path):
        return "❌ فایل قیمت در سرور پیدا نشد."

    try:
        workbook = load_workbook(
            path,
            read_only=True,
            data_only=True
        )

        lines = []

        for sheet in workbook.worksheets:
            lines.append(f"📋 {sheet.title}")

            for row in sheet.iter_rows(values_only=True):
                values = [clean(x) for x in row]

                if not any(values):
                    continue

                lines.append(" | ".join(values))

        workbook.close()

        if not lines:
            return "❌ این فایل خالی است."

        return "\n".join(lines)

    except Exception as e:
        print("Excel error:", e)
        return "❌ خطا در خواندن فایل قیمت."


def send_long(chat_id, text):
    max_length = 3500

    if len(text) <= max_length:
        send_message(chat_id, text)
        return

    while text:
        part = text[:max_length]

        if len(text) > max_length:
            position = part.rfind("\n")

            if position > 500:
                part = part[:position]

        send_message(chat_id, part)

        text = text[len(part):]

        time.sleep(0.3)


def process_message(message):
    chat = message.get("chat") or {}
    chat_id = chat.get("id")

    text = clean(message.get("text"))

    if not chat_id:
        return

    if text == "/start":
        send_message(
            chat_id,
            "سلام 👋\n\n"
            "به ربات گروه تولیدی و بازرگانی عباسی خوش آمدید.\n\n"
            "برای دریافت اطلاعات محصولات، "
            "گزینه مورد نظر را انتخاب کنید."
        )

        time.sleep(0.3)

        main_menu(chat_id)
        return

    if text in ("🔙 منوی اصلی", "منوی اصلی"):
        main_menu(chat_id)
        return

    if text == "📄 دریافت لیست قیمت":
        send_message(
            chat_id,
            "📋 لطفاً گروه مورد نظر را انتخاب کنید:"
        )

        main_menu(chat_id)
        return

    filename = FILES.get(text)

    if filename:
        send_message(
            chat_id,
            "⏳ در حال خواندن لیست قیمت..."
        )

        path = os.path.join(
            os.path.dirname(__file__),
            filename
        )

        content = read_excel(path)

        send_long(chat_id, content)

        send_message(
            chat_id,
            "✅ پایان لیست قیمت\n\n"
            "برای انتخاب گروه دیگر، /start را بفرستید."
        )

        return


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
                timeout=40
            )

            data = response.json()

            if not data.get("ok"):
                print("API error:", data)
                time.sleep(5)
                continue

            for update in data.get("result", []):

                offset = update.get(
                    "update_id",
                    offset
                ) + 1

                message = update.get("message")

                if message:
                    process_message(message)

        except Exception as e:
            print("Bot error:", e)
            time.sleep(5)


@app.route("/")
def home():
    return "Bale bot is running!"


if __name__ == "__main__":

    threading.Thread(
        target=bot_loop,
        daemon=True
    ).start()

    port = int(
        os.environ.get("PORT", 10000)
    )

    app.run(
        host="0.0.0.0",
        port=port
    )
