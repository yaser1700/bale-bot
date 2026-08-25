import os
import time
import threading
import requests
from flask import Flask
from openpyxl import load_workbook


# =========================================================
# تنظیمات ربات
# =========================================================

TOKEN = os.getenv("BALE_TOKEN")

if not TOKEN:
    print("ERROR: BALE_TOKEN is not set!")

BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}"

app = Flask(__name__)


# =========================================================
# فایل‌های قیمت
# =========================================================
# نام دکمه -> کلمات مورد استفاده برای پیدا کردن فایل
# اگر اسم فایل کمی متفاوت باشد هم تلاش می‌کند آن را پیدا کند.
# =========================================================

FILES = {
    "📦 سوکت عباسی": [
        "سوکت عباسی"
    ],

    "🔌 کابل تکنو سبزوار": [
        "کابل تکنو سبزوار"
    ],

    "⚡ وایر عباسی": [
        "وایر عباسی"
    ],

    "🔩 مهره و سنسور": [
        "مهره و سنسور"
    ],

    "💡 قطعات برقی خودرو": [
        "قطعات برقی خودرو"
    ],

    "🧩 خارجات و پلیمر جات": [
        "خارجات",
        "پلیمر"
    ],

    "🔌 کابل خودرو سبزوار": [
        "کابل خودرو سبزوار"
    ],

    "⚙️ شیلنگ خودرو": [
        "شیلنگ خودرو"
    ],

    "⚙️ جلوبندی": [
        "جلوبندی"
    ]
}


# =========================================================
# پیدا کردن فایل
# =========================================================

def find_file(keywords):

    folder = os.path.dirname(os.path.abspath(__file__))

    try:
        files = os.listdir(folder)
    except Exception as e:
        print("Folder error:", e)
        return None

    # فقط فایل‌های Excel
    excel_files = [
        f for f in files
        if f.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm"))
    ]

    # جستجوی دقیق‌تر
    for filename in excel_files:

        name = filename.replace("_", " ").replace("-", " ")

        ok = True

        for keyword in keywords:
            if keyword not in name:
                ok = False
                break

        if ok:
            return os.path.join(folder, filename)

    # جستجوی ساده‌تر
    for filename in excel_files:

        name = filename.replace("_", " ").replace("-", " ")

        for keyword in keywords:

            if keyword in name:
                return os.path.join(folder, filename)

    return None


# =========================================================
# ارسال پیام
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

        print("sendMessage:", response.status_code)
        print(response.text)

        return response

    except Exception as e:

        print("Send message error:", e)

        return None


# =========================================================
# ارسال فایل
# =========================================================

def send_document(chat_id, file_path, caption=""):

    try:

        with open(file_path, "rb") as file:

            data = {
                "chat_id": str(chat_id),
                "caption": caption
            }

            files = {
                "document": (
                    os.path.basename(file_path),
                    file,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            }

            response = requests.post(
                f"{BASE_URL}/sendDocument",
                data=data,
                files=files,
                timeout=120
            )

        print("sendDocument:", response.status_code)
        print(response.text)

        return response

    except Exception as e:

        print("Send document error:", e)

        return None


# =========================================================
# منوی اصلی
# =========================================================

def main_menu(chat_id):

    keyboard = [

        ["📄 دریافت لیست قیمت"],

        ["📦 سوکت عباسی", "🔌 کابل تکنو سبزوار"],

        ["⚡ وایر عباسی", "🔩 مهره و سنسور"],

        ["💡 قطعات برقی خودرو", "🧩 خارجات و پلیمر جات"],

        ["🔌 کابل خودرو سبزوار", "⚙️ شیلنگ خودرو"],

        ["⚙️ جلوبندی"]
    ]

    send_message(
        chat_id,
        "📋 لطفاً گروه مورد نظر را انتخاب کنید:",
        keyboard
    )


# =========================================================
# خواندن Excel
# =========================================================

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

            lines.append(
                f"📋 {sheet.title}"
            )

            for row in sheet.iter_rows(values_only=True):

                values = []

                for value in row:

                    if value is None:
                        values.append("")
                    else:
                        values.append(str(value).strip())

                if not any(values):
                    continue

                line = " | ".join(values)

                lines.append(line)

        workbook.close()

        if not lines:

            return "❌ این فایل خالی است."

        return "\n".join(lines)

    except Exception as e:

        print("Excel error:", e)

        return "❌ خطا در خواندن فایل قیمت."


# =========================================================
# ارسال متن‌های طولانی
# =========================================================

def send_long(chat_id, text):

    max_length = 3500

    if not text:

        send_message(
            chat_id,
            "❌ اطلاعاتی برای نمایش وجود ندارد."
        )

        return

    if len(text) <= max_length:

        send_message(chat_id, text)

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


# =========================================================
# پردازش پیام
# =========================================================

def process_message(message):

    chat = message.get("chat") or {}

    chat_id = chat.get("id")

    text = message.get("text")

    if text is None:
        text = ""

    text = str(text).strip()

    if not chat_id:

        return


    # -----------------------------------------------------
    # شروع
    # -----------------------------------------------------

    if text == "/start":

        send_message(
            chat_id,
            "سلام 👋\n\n"
            "به ربات تولیدی و بازرگانی عباسی خوش آمدید.\n\n"
            "برای دریافت اطلاعات محصولات، گزینه مورد نظر را انتخاب کنید."
        )

        time.sleep(0.3)

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

        send_message(
            chat_id,
            "📋 لطفاً گروه مورد نظر را انتخاب کنید:"
        )

        main_menu(chat_id)

        return


    # -----------------------------------------------------
    # بررسی انتخاب محصول
    # -----------------------------------------------------

    keywords = FILES.get(text)

    if keywords:

        send_message(
            chat_id,
            "⏳ در حال آماده‌سازی لیست قیمت..."
        )

        # پیدا کردن فایل
        path = find_file(keywords)

        if not path:

            send_message(
                chat_id,
                "❌ فایل قیمت این گروه در سرور پیدا نشد.\n\n"
                "لطفاً نام فایل Excel را در GitHub بررسی کنید."
            )

            main_menu(chat_id)

            return


        print("FILE FOUND:")
        print(path)


        # -------------------------------------------------
        # ارسال خود فایل Excel
        # -------------------------------------------------

        result = send_document(
            chat_id,
            path,
            f"📋 لیست قیمت {text}"
        )


        # اگر ارسال فایل موفق نبود
        if result is None or result.status_code != 200:

            send_message(
                chat_id,
                "⚠️ ارسال فایل با مشکل مواجه شد.\n"
                "در حال ارسال اطلاعات داخل فایل..."
            )

            # خواندن Excel و ارسال متن
            content = read_excel(path)

            send_long(
                chat_id,
                content
            )


        else:

            try:

                result_json = result.json()

                if not result_json.get("ok"):

                    print(
                        "Bale API error:",
                        result_json
                    )

                    content = read_excel(path)

                    send_long(
                        chat_id,
                        content
                    )

            except Exception as e:

                print("JSON error:", e)


        time.sleep(0.5)

        send_message(
            chat_id,
            "✅ پایان لیست قیمت\n\n"
            "برای انتخاب گروه دیگر /start را بفرستید."
        )

        return


    # -----------------------------------------------------
    # پیام ناشناخته
    # -----------------------------------------------------

    send_message(
        chat_id,
        "❓ گزینه مورد نظر را از منوی زیر انتخاب کنید:"
    )

    main_menu(chat_id)


# =========================================================
# دریافت پیام‌های بله
# =========================================================

def bot_loop():

    offset = 0

    print("======================================")
    print("Bale bot started...")
    print("======================================")


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
                    "HTTP error:",
                    response.text
                )

                time.sleep(5)

                continue


            data = response.json()


            if not data.get("ok"):

                print(
                    "API error:",
                    data
                )

                time.sleep(5)

                continue


            updates = data.get("result", [])


            for update in updates:

                try:

                    update_id = update.get(
                        "update_id"
                    )

                    if update_id is not None:

                        offset = update_id + 1


                    message = update.get("message")


                    if message:

                        print(
                            "NEW MESSAGE:",
                            message
                        )

                        process_message(
                            message
                        )


                except Exception as e:

                    print(
                        "Message processing error:",
                        e
                    )


        except Exception as e:

            print(
                "Bot loop error:",
                e
            )

            time.sleep(5)


# =========================================================
# صفحه اصلی Render
# =========================================================

@app.route("/")
def home():

    return "Bale Bot is running."


# =========================================================
# Health Check
# =========================================================

@app.route("/health")
def health():

    return "OK"


# =========================================================
# اجرای ربات
# =========================================================

if __name__ == "__main__":

    thread = threading.Thread(
        target=bot_loop,
        daemon=True
    )

    thread.start()

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
