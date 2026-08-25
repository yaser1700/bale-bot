import os
import time
import threading
import requests
from flask import Flask

TOKEN = os.getenv("BALE_TOKEN")
BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}"

app = Flask(__name__)

PDFS = {
    "📦 سوکت عباسی": "سوکت عباسی.pdf",
    "🔌 کابل تکنو سبزوار": "کابل تکنو سبزوار.pdf",
    "⚡ وایر عباسی": "وایر عباسی.pdf",
    "🔩 مهره و سنسور": "مهره و سنسور.pdf",
    "💡 قطعات برقی خودرو": "قطعات برقی خودرو.pdf",
    "🧩 خارجات و پلیمریجات": "خارجات و پلیمریجات.pdf",
    "🔌 کابل خودرو سبزوار": "کابل خودرو سبزوار.pdf",
    "⚙️ شیلنگ خودرو": "شیلنگ خودرو.pdf",
    "⚙️ جلوبندی": "جلوبندی.pdf",
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
        r = requests.post(
            f"{BASE_URL}/sendMessage",
            json=data,
            timeout=30
        )

        print("sendMessage:", r.status_code)

        return r

    except Exception as e:
        print("sendMessage error:", e)
        return None


def send_pdf(chat_id, pdf_path, caption):
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

            r = requests.post(
                f"{BASE_URL}/sendDocument",
                data=data,
                files=files,
                timeout=180
            )

        print("sendDocument:", r.status_code)
        print(r.text[:1000])

        return r

    except Exception as e:
        print("PDF error:", e)
        return None


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
        "📋 گروه مورد نظر را انتخاب کنید:",
        keyboard
    )


def process_message(message):

    chat = message.get("chat") or {}

    chat_id = chat.get("id")

    text = str(
        message.get("text") or ""
    ).strip()

    if not chat_id:
        return


    # /start
    if text == "/start":

        send_message(
            chat_id,
            "سلام 👋\n\n"
            "به ربات تولیدی و بازرگانی عباسی خوش آمدید."
        )

        time.sleep(0.3)

        main_menu(chat_id)

        return


    # منوی اصلی
    if text in (
        "📄 دریافت لیست قیمت",
        "منوی اصلی",
        "🔙 منوی اصلی"
    ):

        main_menu(chat_id)

        return


    # پیدا کردن PDF
    filename = PDFS.get(text)

    if filename:

        base_dir = os.path.dirname(
            os.path.abspath(__file__)
        )

        pdf_path = os.path.join(
            base_dir,
            filename
        )

        print("Selected:", text)
        print("PDF path:", pdf_path)
        print("Exists:", os.path.exists(pdf_path))


        if not os.path.exists(pdf_path):

            send_message(
                chat_id,
                "❌ فایل PDF این گروه در سرور پیدا نشد.\n\n"
                f"نام فایل:\n{filename}"
            )

            return


        send_message(
            chat_id,
            "⏳ در حال ارسال لیست قیمت..."
        )


        result = send_pdf(
            chat_id,
            pdf_path,
            f"📄 لیست قیمت {text}"
        )


        if result is not None:

            try:

                data = result.json()

                if data.get("ok"):

                    send_message(
                        chat_id,
                        "✅ لیست قیمت ارسال شد."
                    )

                    return

                else:

                    print(
                        "Bale error:",
                        data
                    )

            except Exception as e:

                print(
                    "JSON error:",
                    e
                )


        send_message(
            chat_id,
            "❌ ارسال PDF انجام نشد."
        )

        return


    # گزینه ناشناخته
    send_message(
        chat_id,
        "❓ لطفاً یکی از گزینه‌های منو را انتخاب کنید."
    )

    main_menu(chat_id)


def bot_loop():

    offset = 0

    print("================================")
    print("BALE PDF BOT STARTED")
    print("================================")


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


            print(
                "getUpdates:",
                response.status_code
            )


            if response.status_code != 200:

                print(
                    response.text
                )

                time.sleep(5)

                continue


            data = response.json()


            if not data.get("ok"):

                print(
                    "API ERROR:",
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
                    ) + 1
                )


                message = update.get(
                    "message"
                )


                if message:

                    print(
                        "NEW MESSAGE:",
                        message
                    )

                    try:

                        process_message(
                            message
                        )

                    except Exception as e:

                        print(
                            "PROCESS ERROR:",
                            e
                        )


        except Exception as e:

            print(
                "BOT LOOP ERROR:",
                e
            )

            time.sleep(5)


@app.route("/")
def home():

    return "Bale PDF Bot is running."


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
