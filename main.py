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
# نام گروه‌ها
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
# فارسی و اعداد
# =========================================================

def normalize(text):

    text = str(text or "")

    # اعداد فارسی و عربی به انگلیسی
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

    for a, b in replacements.items():
        text = text.replace(a, b)

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
# منوی اصلی
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
# پیدا کردن PDF واقعی
# =========================================================

def find_pdf(group_name):

    base = os.path.dirname(
        os.path.abspath(__file__)
    )

    wanted = compact(
        group_name
    )

    print(
        "LOOKING FOR PDF:",
        group_name
    )

    # همه فایل‌های PDF
    files = []

    try:

        files = os.listdir(base)

    except Exception as e:

        print(
            "LIST FILE ERROR:",
            e
        )

        return None


    for filename in files:

        if not filename.lower().endswith(".pdf"):
            continue

        # decode کردن %20 و غیره
        decoded = unquote(filename)

        name_without_ext = os.path.splitext(
            decoded
        )[0]

        normalized_name = compact(
            name_without_ext
        )

        # تطبیق گروه
        if wanted in normalized_name:

            path = os.path.join(
                base,
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
        "PDF FILES:",
        [
            f for f in files
            if f.lower().endswith(".pdf")
        ]
    )

    return None


# =========================================================
# ارسال PDF
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
# قیمت
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
# پیدا کردن کد در متن
# =========================================================

def find_code(values):

    # اولویت با ستون اول
    if len(values) > 0:

        if values[0] is not None:

            text = str(
                values[0]
            ).strip()

            if re.search(
                r"\d{4,}",
                text
            ):

                match = re.search(
                    r"\d{4,}",
                    text
                )

                return match.group(0)


    # اگر ستون اول کد نبود،
    # تمام سلول‌ها بررسی می‌شوند

    for value in values:

        if value is None:
            continue

        text = str(value)

        match = re.search(
            r"\d{4,}",
            text
        )

        if match:

            return match.group(0)


    return ""


# =========================================================
# جستجوی Excel
# =========================================================

def search_excel(query):

    base = os.path.dirname(
        os.path.abspath(__file__)
    )

    q = normalize(query)
    qc = compact(query)

    print(
        "SEARCH QUERY:",
        q
    )


    excel_files = []

    for filename in os.listdir(base):

        if filename.lower().endswith(".xlsx"):

            if not filename.startswith("~$"):

                excel_files.append(
                    filename
                )


    print(
        "EXCEL FILES:",
        excel_files
    )


    results = []
    seen = set()


    for filename in excel_files:

        path = os.path.join(
            base,
            filename
        )


        try:

            workbook = load_workbook(

                path,

                read_only=True,

                data_only=True
            )


            for sheet in workbook.worksheets:

                for row in sheet.iter_rows(
                    values_only=True
                ):

                    values = list(row)

                    if not values:
                        continue


                    # تمام سلول‌های ردیف
                    cells = []

                    for value in values:

                        if value is None:
                            continue

                        text = str(
                            value
                        ).strip()

                        if text:

                            cells.append(
                                text
                            )


                    if not cells:
                        continue


                    # کل ردیف
                    full_text = " ".join(
                        cells
                    )


                    full_normal = normalize(
                        full_text
                    )

                    full_compact = compact(
                        full_text
                    )


                    # آیا عبارت پیدا شد؟
                    matched = False


                    if q and q in full_normal:

                        matched = True


                    if qc and qc in full_compact:

                        matched = True


                    if not matched:

                        continue


                    # -------------------------
                    # کد
                    # -------------------------

                    code = find_code(
                        values
                    )


                    # -------------------------
                    # نام
                    # -------------------------

                    if len(values) >= 2:

                        name = str(
                            values[1] or ""
                        ).strip()

                    else:

                        name = ""


                    # اگر نام مناسب نبود،
                    # متن کل ردیف را نگه می‌داریم

                    if not name:

                        name = full_text


                    # -------------------------
                    # قیمت
                    # -------------------------

                    price = ""

                    # معمولاً ستون سوم
                    if len(values) >= 3:

                        if values[2] is not None:

                            price = format_price(
                                values[2]
                            )


                    # -------------------------
                    # گروه
                    # -------------------------

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


                    # -------------------------
                    # حذف تکراری
                    # -------------------------

                    key = (

                        normalize(code),

                        normalize(name),

                        price,

                        normalize(group)

                    )


                    if key in seen:

                        continue


                    seen.add(key)


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
        "FOUND:",
        len(results)
    )


    return results


# =========================================================
# ارسال نتایج
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

        "🔎 تعداد نتایج: "
        +
        str(len(results))
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


        if (
            len(message)
            +
            len(block)
            >
            3500
        ):

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
# پردازش پیام
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
        "USER:",
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
    # MENU
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

        group = PDF_GROUPS[text]


        send_message(

            chat_id,

            "⏳ در حال آماده‌سازی فایل PDF..."
        )


        pdf_path = find_pdf(
            group
        )


        if not pdf_path:

            send_message(

                chat_id,

                "❌ فایل PDF این گروه در مخزن پیدا نشد.\n\n"
                "گروه: "
                +
                group
            )

            return


        response = send_pdf(

            chat_id,

            pdf_path,

            f"📄 لیست قیمت {group}"
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
        "======================================"
    )

    print(
        "BALE BOT STARTED"
    )

    print(
        "PDF + FULL EXCEL SEARCH"
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
                "LOOP ERROR:",
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
