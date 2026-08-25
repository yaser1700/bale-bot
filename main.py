import os
import re
import time
import threading
import requests
from flask import Flask
from openpyxl import load_workbook

TOKEN = os.getenv("BALE_TOKEN")

if not TOKEN:
    print("ERROR: BALE_TOKEN is not set")

BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}" if TOKEN else ""

app = Flask(__name__)


# =========================================================
# PDF GROUPS
# =========================================================

PDF_GROUPS = {
    "📦 سوکت عباسی": "سوکت عباسی",
    "🔌 کابل تکنو سبزوار": "کابل تکنو سبزوار",
    "⚡ وایر عباسی": "وایر عباسی",
    "🔩 مهره و سنسور": "مهره و سنسور",
    "💡 قطعات برقی خودرو": "قطعات برقی خودرو",
    "🧩 خارجات و پلیمریجات": "خارجات و پلیمریجات",
    "🔌 کابل خودرو سبزوار": "کابل خودرو سبزوار",
    "⚙️ شیلنگ خودرو": "شیلنگ خودرو",
    "⚙️ جلوبندی": "جلوبندی",
}


# =========================================================
# NORMALIZE
# =========================================================

def normalize(text):

    text = str(text or "")

    # فارسی و عربی به انگلیسی
    text = text.translate(
        str.maketrans(
            "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
            "01234567890123456789"
        )
    )

    replacements = {
        "ي": "ی",
        "ى": "ی",
        "ك": "ک",
        "ۀ": "ه",
        "ة": "ه",
        "\u200c": " ",
        "\u200f": "",
        "\u200e": "",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = text.lower()

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def compact(text):

    text = normalize(text)

    return re.sub(
        r"[^0-9a-zآ-ی]+",
        "",
        text
    )


# =========================================================
# WORD BY WORD SEARCH
# =========================================================

def search_words_match(query, text):

    query = normalize(query)
    text = normalize(text)

    query_words = [
        word
        for word in query.split()
        if word
    ]

    text_words = [
        word
        for word in text.split()
        if word
    ]

    if not query_words:
        return False

    # هر کلمه جستجو باید در یکی از کلمات متن وجود داشته باشد
    for query_word in query_words:

        found = False

        for text_word in text_words:

            if query_word in text_word:
                found = True
                break

        if not found:
            return False

    return True


# =========================================================
# SEND MESSAGE
# =========================================================

def send_message(
    chat_id,
    text,
    keyboard=None
):

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
            "MESSAGE:",
            response.status_code
        )

        return response

    except Exception as e:

        print(
            "SEND MESSAGE ERROR:",
            e
        )

        return None


# =========================================================
# MAIN MENU
# =========================================================

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


# =========================================================
# FIND PDF
# =========================================================

def find_pdf(group_name):

    base_dir = os.path.dirname(
        os.path.abspath(__file__)
    )

    wanted = compact(
        group_name
    )

    print(
        "LOOKING FOR PDF:",
        group_name
    )

    try:

        files = os.listdir(
            base_dir
        )

    except Exception as e:

        print(
            "LIST FILE ERROR:",
            e
        )

        return None


    for filename in files:

        if not filename.lower().endswith(
            ".pdf"
        ):
            continue

        name = os.path.splitext(
            filename
        )[0]

        normalized_name = compact(
            name
        )

        # تطبیق کامل یا بخشی
        if (
            wanted in normalized_name
            or
            normalized_name in wanted
        ):

            path = os.path.join(
                base_dir,
                filename
            )

            print(
                "PDF FOUND:",
                path
            )

            return path


    print(
        "PDF NOT FOUND:",
        group_name
    )

    print(
        "AVAILABLE PDF FILES:"
    )

    for filename in files:

        if filename.lower().endswith(
            ".pdf"
        ):

            print(
                filename
            )

    return None


# =========================================================
# SEND PDF
# =========================================================

def send_pdf(
    chat_id,
    pdf_path,
    caption
):

    try:

        with open(
            pdf_path,
            "rb"
        ) as file:

            response = requests.post(

                f"{BASE_URL}/sendDocument",

                data={
                    "chat_id": str(chat_id),
                    "caption": caption
                },

                files={
                    "document": (
                        os.path.basename(
                            pdf_path
                        ),
                        file,
                        "application/pdf"
                    )
                },

                timeout=180
            )


        print(
            "PDF STATUS:",
            response.status_code
        )

        print(
            "PDF RESPONSE:",
            response.text[:500]
        )

        return response


    except Exception as e:

        print(
            "PDF ERROR:",
            e
        )

        return None


# =========================================================
# FORMAT PRICE
# =========================================================

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


# =========================================================
# FIND PRODUCT CODE
# =========================================================

def find_code(values):

    # ابتدا ستون اول
    if len(values) > 0:

        value = values[0]

        if value is not None:

            text = str(
                value
            ).strip()

            match = re.search(
                r"\d{4,}",
                text
            )

            if match:

                return match.group(
                    0
                )


    # سپس تمام سلول‌ها
    for value in values:

        if value is None:
            continue

        text = str(value)

        match = re.search(
            r"\d{4,}",
            text
        )

        if match:

            return match.group(
                0
            )

    return ""


# =========================================================
# SEARCH ALL EXCEL FILES
# =========================================================

def search_excel(query):

    base_dir = os.path.dirname(
        os.path.abspath(__file__)
    )

    query = normalize(
        query
    )

    print(
        "================================"
    )

    print(
        "SEARCH QUERY:",
        query
    )

    print(
        "================================"
    )


    results = []

    seen = set()


    # پیدا کردن تمام Excel ها
    excel_files = []

    for root, dirs, files in os.walk(
        base_dir
    ):

        for filename in files:

            if not filename.lower().endswith(
                ".xlsx"
            ):
                continue

            if filename.startswith(
                "~$"
            ):
                continue

            excel_files.append(
                os.path.join(
                    root,
                    filename
                )
            )


    print(
        "EXCEL FILE COUNT:",
        len(excel_files)
    )


    # =====================================================
    # هر Excel
    # =====================================================

    for file_path in excel_files:

        filename = os.path.basename(
            file_path
        )

        print(
            "SEARCHING FILE:",
            filename
        )


        try:

            workbook = load_workbook(

                file_path,

                read_only=True,

                data_only=True
            )


            # =================================================
            # هر Sheet
            # =================================================

            for worksheet in workbook.worksheets:

                print(
                    "SHEET:",
                    worksheet.title
                )


                # =================================================
                # هر Row
                # =================================================

                for row in worksheet.iter_rows(
                    values_only=True
                ):

                    values = list(row)

                    if not values:
                        continue


                    # -----------------------------
                    # تمام سلول‌های ردیف
                    # -----------------------------

                    cells = []

                    for value in values:

                        if value is None:
                            continue

                        value_text = str(
                            value
                        ).strip()

                        if value_text:

                            cells.append(
                                value_text
                            )


                    if not cells:
                        continue


                    # متن کامل ردیف
                    full_text = " ".join(
                        cells
                    )


                    # -----------------------------
                    # تطبیق کلمه به کلمه
                    # -----------------------------

                    if not search_words_match(
                        query,
                        full_text
                    ):

                        continue


                    # -----------------------------
                    # کد کالا
                    # -----------------------------

                    code = find_code(
                        values
                    )


                    # -----------------------------
                    # نام کالا
                    # -----------------------------

                    name = ""

                    if len(values) >= 2:

                        if values[1] is not None:

                            name = str(
                                values[1]
                            ).strip()


                    if not name:

                        name = full_text


                    # -----------------------------
                    # قیمت
                    # -----------------------------

                    price = ""

                    if len(values) >= 3:

                        if values[2] is not None:

                            price = format_price(
                                values[2]
                            )


                    # -----------------------------
                    # گروه
                    # -----------------------------

                    group = filename


                    if group.startswith(
                        "لیست_قیمت_"
                    ):

                        group = group[
                            len("لیست_قیمت_"):
                        ]


                    if group.lower().endswith(
                        ".xlsx"
                    ):

                        group = group[:-5]


                    # -----------------------------
                    # جلوگیری از تکرار
                    # -----------------------------

                    key = (

                        normalize(code),

                        normalize(name),

                        price,

                        normalize(group)

                    )


                    if key in seen:

                        continue


                    seen.add(
                        key
                    )


                    # -----------------------------
                    # ذخیره نتیجه
                    # -----------------------------

                    results.append({

                        "code": code,

                        "name": name,

                        "price": price,

                        "group": group

                    })


            workbook.close()


        except Exception as e:

            print(
                "EXCEL ERROR:",
                filename,
                e
            )


    print(
        "================================"
    )

    print(
        "RESULT COUNT:",
        len(results)
    )

    print(
        "================================"
    )


    return results


# =========================================================
# SEND SEARCH RESULTS
# =========================================================

def send_results(
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

        "🔎 تعداد نتایج پیدا شده: "
        +
        str(
            len(results)
        )
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


        # محدودیت حجم پیام
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


# =========================================================
# PROCESS MESSAGE
# =========================================================

def process_message(message):

    chat = (
        message.get(
            "chat"
        )
        or {}
    )


    chat_id = chat.get(
        "id"
    )


    text = str(
        message.get(
            "text"
        )
        or ""
    ).strip()


    if not chat_id:

        return


    print(
        "================================"
    )

    print(
        "NEW MESSAGE:",
        text
    )

    print(
        "================================"
    )


    # =====================================================
    # START
    # =====================================================

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


    # =====================================================
    # MAIN MENU
    # =====================================================

    if text in (

        "📄 دریافت لیست قیمت",

        "منوی اصلی",

        "🔙 منوی اصلی"

    ):

        main_menu(
            chat_id
        )

        return


    # =====================================================
    # SEARCH BUTTON
    # =====================================================

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


    # =====================================================
    # PDF BUTTON
    # =====================================================

    if text in PDF_GROUPS:

        group_name = PDF_GROUPS[
            text
        ]


        send_message(

            chat_id,

            "⏳ در حال پیدا کردن فایل PDF..."
        )


        pdf_path = find_pdf(
            group_name
        )


        if not pdf_path:

            send_message(

                chat_id,

                "❌ فایل PDF این گروه در مخزن پیدا نشد.\n\n"
                "گروه: "
                +
                group_name
            )

            return


        send_message(

            chat_id,

            "📄 فایل پیدا شد؛ در حال ارسال..."
        )


        response = send_pdf(

            chat_id,

            pdf_path,

            f"📄 لیست قیمت {group_name}"
        )


        try:

            if (

                response is not None

                and

                response.json().get(
                    "ok"
                )

            ):

                send_message(

                    chat_id,

                    "✅ فایل PDF ارسال شد."
                )

                return


        except Exception:

            pass


        send_message(

            chat_id,

            "❌ ارسال PDF ناموفق بود."
        )


        return


    # =====================================================
    # SEARCH
    # =====================================================

    if text:

        send_message(

            chat_id,

            "🔎 در حال جستجوی کامل..."
        )


        results = search_excel(
            text
        )


        send_results(

            chat_id,

            results
        )


        return


# =========================================================
# BOT LOOP
# =========================================================

def bot_loop():

    offset = 0


    print(
        "=========================================="
    )

    print(
        "BALE BOT STARTED"
    )

    print(
        "FULL EXCEL SEARCH + PDF"
    )

    print(
        "=========================================="
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


            if not data.get(
                "ok"
            ):

                print(
                    "BALE API ERROR:",
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


# =========================================================
# RENDER
# =========================================================

@app.route("/")
def home():

    return (
        "Bale Bot is running."
    )


# =========================================================
# START SERVER
# =========================================================

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
