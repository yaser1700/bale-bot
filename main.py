import os
import time
import threading
import requests
from flask import Flask
from openpyxl import load_workbook

TOKEN = os.getenv("BALE_TOKEN")
BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}"

app = Flask(__name__)

# =========================
# PDF ها
# =========================

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

        print("sendMessage:", response.status_code)

        return response

    except Exception as e:

        print("sendMessage ERROR:", e)

        return None


# =========================
# ارسال PDF
# =========================

def send_pdf(chat_id, file_path, caption):

    try:

        with open(file_path, "rb") as file:

            response = requests.post(
                f"{BASE_URL}/sendDocument",

                data={
                    "chat_id": str(chat_id),
                    "caption": caption
                },

                files={
                    "document": (
                        os.path.basename(file_path),
                        file,
                        "application/pdf"
                    )
                },

                timeout=180
            )

        print(
            "sendDocument:",
            response.status_code
        )

        print(
            response.text[:500]
        )

        return response

    except Exception as e:

        print(
            "sendPDF ERROR:",
            e
        )

        return None


# =========================
# منوی اصلی
# =========================

def main_menu(chat_id):

    keyboard = [

        ["📄 دریافت لیست قیمت"],

        ["🔎 جستجوی کالا"],

        ["📦 سوکت عباسی",
         "🔌 کابل تکنو سبزوار"],

        ["⚡ وایر عباسی",
         "🔩 مهره و سنسور"],

        ["💡 قطعات برقی خودرو",
         "🧩 خارجات و پلیمریجات"],

        ["🔌 کابل خودرو سبزوار",
         "⚙️ شیلنگ خودرو"],

        ["⚙️ جلوبندی"]
    ]

    send_message(
        chat_id,
        "📋 گزینه مورد نظر را انتخاب کنید:",
        keyboard
    )


# =========================
# نرمال‌سازی فارسی
# =========================

def normalize(text):

    text = str(text or "")

    text = text.strip().lower()

    text = text.replace("ي", "ی")
    text = text.replace("ى", "ی")
    text = text.replace("ك", "ک")

    text = text.replace(
        "\u200c",
        " "
    )

    return text


# =========================
# فرمت قیمت
# =========================

def format_price(value):

    if value is None:
        return ""

    text = str(value).strip()

    try:

        number = float(
            text.replace(",", "")
        )

        if number.is_integer():

            return f"{int(number):,}"

    except Exception:

        pass

    return text


# =========================
# جستجوی Excel
# =========================

def search_excel(query):

    base_dir = os.path.dirname(
        os.path.abspath(__file__)
    )

    query = normalize(query)

    results = []

    # همه Excel های موجود در پروژه
    excel_files = []

    for filename in os.listdir(base_dir):

        if filename.lower().endswith(
            ".xlsx"
        ):

            excel_files.append(filename)


    for filename in excel_files:

        file_path = os.path.join(
            base_dir,
            filename
        )

        try:

            workbook = load_workbook(
                file_path,
                read_only=True,
                data_only=True
            )


            for sheet in workbook.worksheets:

                for row in sheet.iter_rows(
                    values_only=True
                ):

                    values = list(row)

                    if len(values) < 2:
                        continue


                    # ستون اول = کد
                    code = ""

                    if values[0] is not None:

                        code = str(
                            values[0]
                        ).strip()


                    # ستون دوم = نام
                    name = ""

                    if values[1] is not None:

                        name = str(
                            values[1]
                        ).strip()


                    # ستون سوم = قیمت
                    price = ""

                    if len(values) >= 3:

                        if values[2] is not None:

                            price = format_price(
                                values[2]
                            )


                    if not code and not name:

                        continue


                    code_search = normalize(
                        code
                    )

                    name_search = normalize(
                        name
                    )


                    # جستجو در کد یا نام
                    if (
                        query in code_search
                        or
                        query in name_search
                    ):

                        group = filename

                        if group.startswith(
                            "لیست_قیمت_"
                        ):

                            group = group[
                                len("لیست_قیمت_"):
                            ]


                        if group.endswith(
                            ".xlsx"
                        ):

                            group = group[
                                :-5
                            ]


                        results.append({

                            "code": code,

                            "name": name,

                            "price": price,

                            "group": group
                        })


                        # حداکثر ۲۰ نتیجه
                        if len(results) >= 20:

                            workbook.close()

                            return results


            workbook.close()


        except Exception as e:

            print(
                "Excel ERROR:",
                filename,
                e
            )


    return results


# =========================
# ارسال نتایج جستجو
# =========================

def send_search_results(
    chat_id,
    results
):

    if not results:

        send_message(
            chat_id,
            "❌ کالایی با این کد یا نام پیدا نشد."
        )

        return


    message = "🔎 نتایج جستجو:\n\n"


    for item in results:

        message += (

            f"📦 کد کالا: "
            f"{item['code']}\n"

            f"📝 نام کالا: "
            f"{item['name']}\n"

            f"💰 قیمت: "
            f"{item['price']} ریال\n"

            f"📁 گروه: "
            f"{item['group']}\n"

            "────────────────\n"
        )


    # تقسیم پیام‌های طولانی
    while message:

        part = message[:3500]


        if len(message) > 3500:

            position = part.rfind(
                "\n"
            )

            if position > 500:

                part = part[:position]


        send_message(
            chat_id,
            part
        )


        message = message[
            len(part):
        ]


# =========================
# پردازش پیام
# =========================

def process_message(message):

    chat = message.get(
        "chat"
    ) or {}

    chat_id = chat.get(
        "id"
    )

    text = str(
        message.get("text") or ""
    ).strip()


    if not chat_id:

        return


    # =====================
    # START
    # =====================

    if text == "/start":

        send_message(
            chat_id,
            "سلام 👋\n\n"
            "به ربات تولیدی و بازرگانی عباسی خوش آمدید."
        )

        time.sleep(0.3)

        main_menu(
            chat_id
        )

        return


    # =====================
    # منوی اصلی
    # =====================

    if text in (

        "📄 دریافت لیست قیمت",

        "منوی اصلی",

        "🔙 منوی اصلی"

    ):

        main_menu(
            chat_id
        )

        return


    # =====================
    # جستجو
    # =====================

    if text == "🔎 جستجوی کالا":

        send_message(
            chat_id,

            "🔎 کد کالا یا بخشی از نام کالا را بفرستید.\n\n"

            "مثال:\n"

            "100158\n\n"

            "یا:\n"

            "کابل دنا"
        )

        return


    # =====================
    # ارسال PDF
    # =====================

    if text in PDFS:

        base_dir = os.path.dirname(
            os.path.abspath(__file__)
        )


        pdf_name = PDFS[text]


        pdf_path = os.path.join(
            base_dir,
            pdf_name
        )


        print(
            "PDF:",
            pdf_path
        )


        if not os.path.exists(
            pdf_path
        ):

            send_message(
                chat_id,

                "❌ فایل PDF این گروه پیدا نشد."
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


        try:

            if (
                result is not None
                and
                result.json().get("ok")
            ):

                send_message(
                    chat_id,
                    "✅ لیست قیمت ارسال شد."
                )

                return

        except Exception:

            pass


        send_message(
            chat_id,
            "❌ ارسال PDF انجام نشد."
        )

        return


    # =====================
    # جستجوی خودکار
    # =====================

    if text:

        send_message(
            chat_id,
            "🔎 در حال جستجو..."
        )


        results = search_excel(
            text
        )


        send_search_results(
            chat_id,
            results
        )


        return


# =========================
# حلقه ربات
# =========================

def bot_loop():

    offset = 0

    print(
        "================================"
    )

    print(
        "BALE PDF + SEARCH BOT STARTED"
    )

    print(
        "================================"
    )


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


            if response.status_code != 200:

                print(
                    "HTTP ERROR:",
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


# =========================
# Render
# =========================

@app.route("/")
def home():

    return (
        "Bale PDF + Search Bot "
        "is running."
    )


# =========================
# اجرای ربات
# =========================

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
