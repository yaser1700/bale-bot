import os
import re
import time
import threading
import requests
from flask import Flask
from openpyxl import load_workbook
from urllib.parse import unquote

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

    # فارسی و عربی -> انگلیسی
    digits = str.maketrans(
        "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
        "01234567890123456789"
    )

    text = text.translate(digits)

    replacements = {
        "ي": "ی",
        "ى": "ی",
        "ك": "ک",
        "ۀ": "ه",
        "ة": "ه",
        "\u200c": " ",
        "\u200f": "",
        "\u200e": "",
        "%20": " ",
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
# PRICE
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
            "MESSAGE ERROR:",
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

        decoded_name = unquote(
            filename
        )

        name_without_ext = os.path.splitext(
            decoded_name
        )[0]

        normalized_name = compact(
            name_without_ext
        )

        if wanted in normalized_name:

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
            "PDF RESPONSE:",
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


# =========================================================
# FIND HEADER
# =========================================================

def find_headers(worksheet):

    code_col = None
    name_col = None
    price_col = None
    header_row = None


    # فقط 50 ردیف اول را برای عنوان‌ها بررسی می‌کنیم
    for row_number, row in enumerate(

        worksheet.iter_rows(
            min_row=1,
            max_row=50,
            values_only=True
        ),

        start=1

    ):

        headers = []

        for value in row:

            headers.append(
                normalize(value)
            )


        current_code = None
        current_name = None
        current_price = None


        for index, header in enumerate(
            headers
        ):

            if not header:
                continue


            # کد کالا
            if (
                "کد کالا" in header
                or
                header == "کد"
                or
                "کدکالا" in header
            ):

                current_code = index


            # نام کالا
            if (
                "نام کالا" in header
                or
                header == "نام"
                or
                "نامکالا" in header
            ):

                current_name = index


            # قیمت
            if "قیمت" in header:

                current_price = index


        if (
            current_code is not None
            and
            current_name is not None
        ):

            code_col = current_code
            name_col = current_name
            price_col = current_price
            header_row = row_number

            break


    return (
        header_row,
        code_col,
        name_col,
        price_col
    )


# =========================================================
# SEARCH EXCEL
# =========================================================

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

    # کلمات جستجو
    query_words = [

        word

        for word in query_normal.split()

        if word

    ]


    print(
        "===================================="
    )

    print(
        "SEARCH:",
        query
    )

    print(
        "NORMAL:",
        query_normal
    )

    print(
        "WORDS:",
        query_words
    )

    print(
        "===================================="
    )


    results = []

    seen = set()


    # =====================================================
    # پیدا کردن Excel ها
    # =====================================================

    excel_files = []

    for filename in os.listdir(
        base_dir
    ):

        if filename.lower().endswith(
            ".xlsx"
        ):

            if not filename.startswith(
                "~$"
            ):

                excel_files.append(
                    filename
                )


    print(
        "EXCEL FILES:",
        excel_files
    )


    # =====================================================
    # فایل‌ها
    # =====================================================

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


            # =================================================
            # Sheet ها
            # =================================================

            for worksheet in workbook.worksheets:

                (
                    header_row,
                    code_col,
                    name_col,
                    price_col
                ) = find_headers(
                    worksheet
                )


                # اگر جدول کالا نبود
                if (
                    header_row is None
                    or
                    code_col is None
                    or
                    name_col is None
                ):

                    print(
                        "SKIP SHEET:",
                        filename,
                        worksheet.title
                    )

                    continue


                print(
                    "USE SHEET:",
                    filename,
                    worksheet.title,
                    "HEADER:",
                    header_row,
                    "CODE:",
                    code_col,
                    "NAME:",
                    name_col,
                    "PRICE:",
                    price_col
                )


                # =================================================
                # خواندن کالاها
                # =================================================

                for row in worksheet.iter_rows(

                    min_row=header_row + 1,

                    values_only=True

                ):

                    if not row:
                        continue


                    # =================================================
                    # CODE
                    # =================================================

                    code = ""

                    if code_col < len(row):

                        if row[code_col] is not None:

                            code = str(
                                row[code_col]
                            ).strip()


                    # =================================================
                    # NAME
                    # =================================================

                    name = ""

                    if name_col < len(row):

                        if row[name_col] is not None:

                            name = str(
                                row[name_col]
                            ).strip()


                    # ردیف خالی
                    if not code and not name:

                        continue


                    # =================================================
                    # PRICE
                    # =================================================

                    price = ""

                    if (
                        price_col is not None
                        and
                        price_col < len(row)
                    ):

                        if row[price_col] is not None:

                            price = format_price(
                                row[price_col]
                            )


                    # =================================================
                    # NORMALIZE
                    # =================================================

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


                    # =================================================
                    # MATCH CODE
                    # =================================================

                    code_match = False

                    if query_compact:

                        if (
                            query_compact
                            == code_compact
                        ):

                            code_match = True

                        elif (
                            query_compact
                            in code_compact
                        ):

                            code_match = True


                    # =================================================
                    # MATCH NAME
                    #
                    # مثال:
                    #
                    # کابل دنا
                    #
                    # نام کالا:
                    # کابل درب باز کن خارجی جلو دنا پلاس
                    #
                    # هر دو کلمه باید وجود داشته باشند.
                    # =================================================

                    name_match = False

                    if query_words:

                        all_words_found = True

                        for word in query_words:

                            if word not in name_normal:

                                all_words_found = False

                                break


                        if all_words_found:

                            name_match = True


                    # =================================================
                    # اگر نام یا کد پیدا نشد
                    # =================================================

                    if (
                        not code_match
                        and
                        not name_match
                    ):

                        continue


                    # =================================================
                    # GROUP
                    # =================================================

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


                    # =================================================
                    # DUPLICATE
                    # =================================================

                    key = (

                        normalize(code),

                        normalize(name),

                        price,

                        normalize(group)

                    )


                    if key in seen:

                        continue


                    seen.add(key)


                    # =================================================
                    # RESULT
                    # =================================================

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
                str(e)
            )


    print(
        "TOTAL RESULTS:",
        len(results)
    )

    return results


# =========================================================
# SEND RESULTS
# =========================================================

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

        "🔎 تعداد نتایج پیدا شده: "
        +
        str(len(results))
    )


    message = ""


    for item in results:

        block = (

            "📦 کد کالا: "
            +
            str(item["code"])
            +
            "\n"

            "📝 نام کالا: "
            +
            str(item["name"])
            +
            "\n"

            "💰 قیمت: "
            +
            str(item["price"])
            +
            " ریال\n"

            "📁 گروه: "
            +
            str(item["group"])
            +
            "\n"

            "────────────────\n"

        )


        # محدودیت پیام بله
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
        message.get("chat")
        or {}
    )


    chat_id = chat.get(
        "id"
    )


    text = str(
        message.get("text")
        or ""
    ).strip()


    if not chat_id:

        return


    print(
        "NEW MESSAGE:",
        text
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

            "🔎 کد کالا یا نام کالا را بفرستید.\n\n"
            "مثال:\n"
            "100158\n\n"
            "یا:\n"
            "کابل دنا"
        )

        return


    # =====================================================
    # PDF
    # =====================================================

    if text in PDF_GROUPS:

        group_name = PDF_GROUPS[
            text
        ]


        send_message(

            chat_id,

            "⏳ در حال آماده‌سازی فایل PDF..."
        )


        pdf_path = find_pdf(
            group_name
        )


        if not pdf_path:

            send_message(

                chat_id,

                "❌ فایل PDF این گروه پیدا نشد.\n\n"
                "گروه: "
                +
                group_name
            )

            return


        response = send_pdf(

            chat_id,

            pdf_path,

            "📄 لیست قیمت "
            +
            group_name
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

            "❌ ارسال فایل PDF ناموفق بود."
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


        send_search_results(

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
        "======================================"
    )

    print(
        "BALE BOT STARTED"
    )

    print(
        "PDF + SMART EXCEL SEARCH"
    )

    print(
        "======================================"
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
                    "BALE API ERROR:",
                    data
                )

                time.sleep(
                    5
                )

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

    return "Bale Bot is running."


# =========================================================
# START
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
