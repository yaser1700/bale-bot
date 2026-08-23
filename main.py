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
    requests.post(
        f"{BASE_URL}/sendMessage",
        json={
            "chat_id": chat_id,
            "text": text
        },
        timeout=20
    )

def bot_loop():
    offset = 0

    while True:
        try:
            response = requests.get(
                f"{BASE_URL}/getUpdates",
                params={
                    "offset": offset,
                    "timeout": 30
                },
                timeout=40
            ).json()

            for update in response.get("result", []):
                offset = update["update_id"] + 1

                message = update.get("message", {})
                chat_id = message.get("chat", {}).get("id")
                text = message.get("text", "")

                if chat_id and text == "/start":
                    send_message(
                        chat_id,
                        "گروه تولیدی بازرگانی عباسی\n"
                        "تولید و پخش انواع قطعات خودرو\n\n"
                        "https://tecnoyadakabbasi.ir"
                    )

        except Exception:
            time.sleep(5)

if __name__ == "__main__":
    threading.Thread(target=bot_loop, daemon=True).start()

    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
