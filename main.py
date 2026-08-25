import os
import time
import threading
import requests
from flask import Flask

TOKEN = os.getenv("BALE_TOKEN")
BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}"

app = Flask(__name__)


# -------------------------
# صفحه اصلی Render
# -------------------------
@app.route("/")
def home():
    return "Bale bot is running!"


# -------------------------
# ارسال پیام
# -------------------------
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


# -------------------------
# منوی اصلی
# -------------------------
def main_menu(chat_id):
    keyboard = [
        ["📄 دریافت لیست قیمت"],
        ["📦 سوکت عباسی", "🔌 کابل تکنو سبزوار"],
        ["⚡ وایر عباسی", "🔩 مهره و سنسور"],
        ["🧩 خارجات و پلیمری جات", "💡 قطعات برقی خودرو"],
        ["🔌 کابل خودرو سبزوار", "🛞 شیلنگ خودرو"],
        ["🚗 جلوبندی"]
    ]

    send_message(
        chat_id,
        "لطفاً یکی از گزینه‌های زیر را انتخاب کنید:",
        keyboard
    )


# -------------------------
# گروه‌های کالا
# -------------------------
groups = [
    "📦 سوکت عباسی",
    "🔌 کابل تکنو سبزوار",
    "⚡ وایر عباسی",
    "🔩 مهره و سنسور",
    "🧩 خارجات و پلیمری جات",
    "💡 قطعات برقی خودرو",
    "🔌 کابل خودرو سبزوار",
    "🛞 شیلنگ خودرو",
    "🚗 جلوبندی"
]


# -------------------------
# پردازش پیام‌ها
# -------------------------
def process_message(message):

    chat = message.get("chat", {})
    chat_id = chat.get("id")

    text = message.get("text", "").strip()

    if not chat_id:
        return

    # شروع ربات
    if text == "/start":
        send_message(
            chat_id,
            "سلام 👋\n\n"
            "به ربات گروه تولیدی و بازرگانی عباسی خوش آمدید.\n\n"
            "برای دریافت اطلاعات محصولات، گزینه مورد نظر را انتخاب کنید."
        )

        time.sleep(0.5)
        main_menu(chat_id)
        return

    # نمایش منوی قیمت
    if text == "📄 دریافت لیست قیمت":
        send_message(
            chat_id,
            "📄 دریافت لیست قیمت\n\n"
            "لطفاً گروه مورد نظر را انتخاب کنید:"
        )

        keyboard = [
            ["📦 سوکت عباسی", "🔌 کابل تکنو سبزوار"],
            ["⚡ وایر عباسی", "🔩 مهره و سنسور"],
            ["🧩 خارجات و پلیمری جات", "💡 قطعات برقی خودرو"],
            ["🔌 کابل خودرو سبزوار", "🛞 شیلنگ خودرو"],
            ["🚗 جلوبندی"],
            ["🔙 بازگشت به منوی اصلی"]
        ]

        send_message(
            chat_id,
            "گروه محصولات:",
            keyboard
        )
        return

    # انتخاب گروه
    if text in groups:
        send_message(
            chat_id,
            f"📦 {text}\n\n"
            "لیست قیمت این گروه به‌زودی از فایل PDF ارسال خواهد شد.\n\n"
            "فعلاً فایل قیمت این گروه را به ربات اضافه نکرده‌ایم."
        )
        return

    # بازگشت
    if text == "🔙 بازگشت به منوی اصلی":
        main_menu(chat_id)
        return

    # پیام معمولی
    send_message(
        chat_id,
        "پیام شما دریافت شد ✅\n\n"
        "لطفاً از منوی ربات گزینه مورد نظر را انتخاب کنید."
    )


# -------------------------
# دریافت پیام‌های بله
# -------------------------
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

            updates = data.get("result", [])

            for update in updates:

                offset = update.get("update_id", offset) + 1

                message = update.get("message")

                if message:
                    process_message(message)

        except Exception as e:

            print("Bot error:", e)
            time.sleep(5)


# -------------------------
# اجرای ربات
# -------------------------
threading.Thread(
    target=bot_loop,
    daemon=True
).start()


if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(
        host="0.0.0.0",
        port=port
    )
