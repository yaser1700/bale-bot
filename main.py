import os
import re
import time
import threading
import requests
from flask import Flask
from openpyxl import load_workbook

TOKEN = os.getenv("BALE_TOKEN")
BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}" if TOKEN else ""

app = Flask(__name__)

# =========================
# PDF FILES
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
# SEND MESSAGE
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

        print(
            "sendMessage:",
            response.status_code
        )

        return response

    except Exception as e:

        print(
            "sendMessage ERROR:",
            e
        )

        return None


# =========================
# SEND PDF
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
            "PDF ERROR:",
            e
        )

        return None


# =========================
# MAIN MENU
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
# NORMALIZE PERSIAN
# =========================

def normalize(text):

    text = str(text or "")

    text = text.strip().lower()

    replacements = {

        "ي": "ی",
        "ى": "ی",
        "ك": "ک",
        "ۀ": "ه",
        "ة": "ه",

        "\u200c": " ",
        "\u200f": "",
        "\u200e": ""

    }

    for old, new in replacements.items():

        text = text.replace(
            old,
            new
        )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# =========================
# COMPACT TEXT
# =========================

def compact(text):

    text = normalize(text)

    return re.sub(
        r"[\s\-_./]+",
        "",
        text
    )


# =========================
# FORMAT PRICE
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
# SEARCH ALL EXCEL FILES
# =========================

def search_excel(query):

    base_dir = os.path.dirname(
        os.path.abspath(__file__)
    )

    query_normal = normalize(
        query
    )

    query_compact = compact(
        query
    )

    exact_results = []

    partial_results = []


    # پیدا کردن تمام Excel ها
    excel_files = [

        filename

        for filename in os.listdir(
            base_dir
        )

        if filename.lower().endswith(
            ".xlsx"
        )
    ]


    print(
        "Excel files:",
        excel_files
    )


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


            for worksheet in workbook.worksheets:

                for row in worksheet.iter_rows(
                    values_only=True
                ):

                    values = list(row)


                    if len(values) < 2:

                        continue


                    # -------------------------
                    # CODE
                    # -------------------------

                    code = ""

                    if values[0] is not None:

                        code = str(
                            values[0]
                        ).strip()


                    # -------------------------
                    # NAME
                    # -------------------------

                    name = ""

                    if values[1] is not None:

                        name = str(
                            values[1]
                        ).strip()


                    # -------------------------
                    # PRICE
                    # -------------------------

                    price = ""

                    if len(values) >= 3:

                        if values[2] is not None:

                            price = format_price(
                                values[2]
                            )


                    if not code and not name:

                        continue


                    code_normal = normalize(
                        code
                    )

                    name_normal = normalize(
                        name
                    )

                    code_compact = compact(
                        code
                    )

                    name_compact = compact(
                        name
                    )


                    # -------------------------
                    # EXACT MATCH
                    # -------------------------

                    exact = (

                        query_normal == code_normal

                        or

                        query_normal == name_normal

                        or

                        (
                            query_compact
                            and
                            query_compact == code_compact
                        )

                        or

                        (
                            query_compact
                            and
                            query_compact == name_compact
                        )

                    )


                    # -------------------------
                    # PARTIAL MATCH
                    # -------------------------

                    partial = (

                        query_normal in code_normal

                        or

                        query_normal in name_normal

                        or

                        (
                            query_compact
                            and
                            query_compact in code_compact
                        )

                        or

                        (
                            query_compact
                            and
                            query_compact in name_compact
                        )

                    )


                    if not partial:

                        continue


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

                        group = group[:-5]


                    item = {

                        "code": code,

                        "name": name,

                        "price": price,

                        "group": group

                    }


                    if exact:

                        exact_results.append(
                            item
                        )

                    else:

                        partial_results.append(
                            item
                        )


            workbook.close()


        except Exception as e:

            print(
                "Excel ERROR:",
                filename,
                e
            )


    # -------------------------
    # REMOVE DUPLICATES
    # -------------------------

    all_results = []

    seen = set()


    for item in (
        exact_results +
        partial_results
    ):

        key = (

            normalize(
                item["code"]
            ),

            normalize(
                item["name"]
            ),

            item["price"],

            item["group"]

        )


        if key in seen:

            continue


        seen.add(key)

        all_results.append(
            item
        )


    return all_results


# =========================
# SEND SEARCH RESULTS
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


    send_message(

        chat_id,

        f"🔎 تعداد نتایج پیدا شده: "
        f"{len(results)}"
    )


    message = ""


    for item in results:

        block = (

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


        # تقسیم پیام در صورت طولانی شدن

        if (
            len(message)
            +
            len(block)
            >
            3500
        ):

            if message:

                send_message(
                    chat_id,
                    message
                )

            message = block

        else:

            message += block


    if message:

        send_message(
            chat_id,
            message
        )


# =========================
# PROCESS MESSAGE
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


        time.sleep(
            0.3
        )


        main_menu(
            chat_id
        )


        return


    # =====================
    # MAIN MENU
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
    # SEARCH BUTTON
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
    # PDF
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
            "PDF PATH:",
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

                result.json().get(
                    "ok"
                )

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
    # SEARCH
    # =====================

    if text:

        send_message(

            chat_id,

            "🔎 در حال جستجوی کامل..."
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
# BOT LOOP
# =========================

def bot_loop():

    offset = 0


    print(
        "================================"
    )

    print(
        "BALE PDF + FULL SEARCH BOT"
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


                time.sleep(
                    5
                )


                continue


            data = response.json()


            if not data.get("ok"):

                print(
                    "API ERROR:",
                    data
                )


                time.sleep(
                    5
                )


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


            time.sleep(
                5
            )


# =========================
# RENDER
# =========================

@app.route("/")
def home():

    return (
        "Bale PDF + Full Search Bot "
        "is running."
    )


# =========================
# RUN
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
