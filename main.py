import os
import time
import requests

TOKEN = os.getenv("BALE_TOKEN")
BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}"

def send_message(chat_id, text):
    requests.post(
        f"{BASE_URL}/sendMessage",
        json={"chat_id": chat_id, "text": text}
    )

offset = 0

while True:
    try:
        response = requests.get(
            f"{BASE_URL}/getUpdates",
            params={"offset": offset, "timeout": 30}
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
                    "برای مشاهده سایت:\n"
                    "https://tecnoyadakabbasi.ir"
                )

    except Exception:
        time.sleep(5)
