import os
import time
import threading
import requests
from flask import Flask

TOKEN = os.getenv("BALE_TOKEN")
BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}"

app = Flask(__name__)


@app.route("/")
def home():
    return "Bale bot is running!"


def send_message(chat_id, text):
    try:
        requests.post(
            f"{BASE_URL}/sendMessage",
            json={
                "chat_id": chat_id,
                "text": text
            },
            timeout=20
        )
    except Exception as e:
        print("Send error:", e)


def bot_loop():
    offset = 0

    while True:
        try:
            response = requests.get(
                f"{BASE_URL}/getUpdates",
                params={
                    "offset": offset,
                    "timeout": 25
                },
                timeout=30
            )

            data = response.json()

            for update in data.get("result", []):
                offset = update["update_id"] + 1

                message = update.get("message")
                if not message:
                    continue

                chat = message.get("chat", {})
                chat_id = chat.get("id")

                text = message.get("text", "")

                if text == "/start":
                    send_message(
                        chat_id,
                        "سلام 👋\n"
                        "به ربات گروه تولیدی بازرگانی عباسی خوش آمدید.\n\n"
                        "برای دریافت اطلاعات محصولات، پیام خود را ارسال کنید."
                    )

                elif text:
                    send_message(
                        chat_id,
                        "پیام شما دریافت شد ✅\n"
                        "به زودی پاسخ داده می‌شود."
                    )

        except Exception as e:
            print("Bot error:", e)
            time.sleep(5)


threading.Thread(target=bot_loop, daemon=True).start()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
