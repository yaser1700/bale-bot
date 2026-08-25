import os
import re
import time
import threading
import requests
from flask import Flask
from openpyxl import load_workbook
from urllib.parse import unquote

TOKEN = os.getenv("BALE_TOKEN")
BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}" if TOKEN else ""

app = Flask(__name__)

# =========================================================
# اطلاعات تماس و سایت
# =========================================================

PHONE = "09377700031"
WEBSITE = "https://www.tecnoyadakabbasi.ir"


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
        "ؤ": "و",
        "إ": "ا",
        "أ": "ا",
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
# SEND MESSAGE
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

        return requests.post(
            f"{BASE_URL}/sendMessage",
            json=data,
            timeout=30
        )

    except Exception as e:

        print("MESSAGE ERROR:", e)

        return None


# =========================================================
# MAIN MENU
# =========================================================

def main_menu(chat_id):

    keyboard = [

        ["📄 دریافت لیست قیمت"],

        ["🔎 جستجوی کالا"],

        ["🌐 ورود به سایت",
         "📞 تماس مستقیم"],

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

    base = os.path.dirname(
        os.path.abspath(__file__)
    )

    wanted = compact(group_name)

    try:

        files = os.listdir(base)

    except Exception as e:

        print("LIST FILE ERROR:", e)

        return None

    for filename in files:

        if not filename.lower().endswith(".pdf"):
            continue

        decoded = unquote(filename)

        name_without_ext = os.path.splitext(
            decoded
        )[0]

        normalized_name = compact(
            name_without_ext
        )

        if (
            wanted in normalized_name
            or
            normalized_name in wanted
        ):

            return os.path.join(
                base,
                filename
            )

    return None


# =========================================================
# SEND PDF
# =========================================================

def send_pdf(chat_id, pdf_path, caption):

    try:

        with open(pdf_path, "rb") as file:

            response = requests.post(

                f"{BASE_URL}/sendDocument",

                data={
                    "chat_id": str(chat_id),
                    "caption": caption
                },

                files={
                    "document": (
                        os.path.basename(pdf_path),
                        file,
                        "application/pdf"
                    )
                },

                timeout=180
            )

        return response

    except Exception as e:

        print("PDF ERROR:", e)

        return None


# =========================================================
# PRICE
# =========================================================

def format_price(value):

    if value is None:
        return ""

    if isinstance(value, int):
        return f"{value:,}"

    if isinstance(value, float):

        if value.is_integer():
            return f"{int(value):,}"

        return str(value)

    text = str(value).strip()

    text = text.replace(",", "")
    text = text.replace("٬", "")

    try:

        number = float(text)

        if number.is_integer():
            return f"{int(number):,}"

    except Exception:

        pass

    return text


# =========================================================
# FIND COLUMNS
# =========================================================

def get_columns(row):

    columns = {}

    for index, value in enumerate(row):

        if value is None:
            continue

        name = normalize(value)

        if "گروه" in name:

            columns["group"] = index

        elif (
            "کد کالا" in name
            or
            name == "کد"
            or
            "شناسه" in name
        ):

            columns["code"] = index

        elif (
            "نام کالا" in name
            or
            name == "نام"
        ):

            columns["name"] = index

        elif "قیمت" in name:

            columns["price"] = index

        elif (
            "توضیحات" in name
            or
            "توضیح" in name
        ):

            columns["description"] = index

    return columns


# =========================================================
# SEARCH MATCH
# =========================================================

def search_matches(
    query,
    code,
    name,
    group,
    description
):

    q = normalize(query)
    qc = compact(query)

    if not q:
        return False

    full_text = " ".join([

        normalize(code),
        normalize(name),
        normalize(group),
        normalize(description)

    ])

    full_compact = compact(full_text)

    if q in full_text:
        return True

    if qc and qc in full_compact:
        return True

    words = [
        word
        for word in q.split()
        if len(word) >= 2
    ]

    if not words:
        return False

    for word in words:

        if word not in full_text:
            return False

    return True


# =========================================================
# SEARCH EXCEL
# =========================================================

def search_excel(query):

    base = os.path.dirname(
        os.path.abspath(__file__)
    )

    excel_files = [

        filename

        for filename in os.listdir(base)

        if filename.lower().endswith(".xlsx")

        and not filename.startswith("~$")
    ]

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

                header_row = None
                columns = None

                # پیدا کردن ردیف عنوان
                for row_number, row in enumerate(
                    sheet.iter_rows(values_only=True),
                    start=1
                ):

                    found = get_columns(row)

                    if (
                        "code" in found
                        and
                        "name" in found
                        and
                        "price" in found
                    ):

                        header_row = row_number
                        columns = found

                        break

                if not columns:
                    continue

                # خواندن کالاها
                for row in sheet.iter_rows(
                    min_row=header_row + 1,
                    values_only=True
                ):

                    if not row:
                        continue

                    def get_value(key):

                        index = columns.get(key)

                        if index is None:
                            return ""

                        if index >= len(row):
                            return ""

                        return str(
                            row[index] or ""
                        ).strip()

                    group = get_value("group")
                    code = get_value("code")
                    name = get_value("name")
                    description = get_value("description")

                    price = ""

                    price_index = columns.get("price")

                    if (
                        price_index is not None
                        and
                        price_index < len(row)
                    ):

                        price = format_price(
                            row[price_index]
                        )

                    if not code and not name:
                        continue

                    if not search_matches(
                        query,
                        code,
                        name,
                        group,
                        description
                    ):
                        continue

                    key = (
                        normalize(code),
                        normalize(name),
                        normalize(group),
                        price
                    )

                    if key in seen:
                        continue

                    seen.add(key)

                    results.append({

                        "code": code,

                        "name": name,

                        "price": price,

                        "group": group,

                        "description": description

                    })

            workbook.close()

        except Exception as e:

            print(
                "EXCEL ERROR:",
                filename,
                e
            )

    return results


# =========================================================
# SEND SEARCH RESULTS
# =========================================================

def send_results(chat_id, results):

    if not results:

        send_message(
            chat_id,
            "❌ کالایی با این کد یا نام پیدا نشد."
        )

        return

    send_message(
        chat_id,
        "🔎 تعداد کالاهای پیدا شده: "
        + str(len(results))
    )

    message = ""

    for item in results:

        block = (

            f"📦 کد کالا: {item['code']}\n"

            f"📝 نام کالا: {item['name']}\n"

            f"💰 قیمت: {item['price']} ریال\n"

            f"📁 گروه: {item['group']}\n"

        )

        if item["description"]:

            block += (
                f"ℹ️ توضیحات: "
                f"{item['description']}\n"
            )

        block += (
            "────────────────\n"
        )

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

    chat_id = chat.get("id")

    text = str(
        message.get("text")
        or ""
    ).strip()

    if not chat_id:
        return


    # =====================================================
    # START
    # =====================================================

    if text == "/start":

        send_message(
            chat_id,
            "سلام 👋\n\n"
            "به ربات تولیدی و بازرگانی عباسی خوش آمدید."
        )

        time.sleep(0.3)

        main_menu(chat_id)

        return


    # =====================================================
    # MAIN MENU
    # =====================================================

    if text in (
        "📄 دریافت لیست قیمت",
        "منوی اصلی",
        "🔙 منوی اصلی"
    ):

        main_menu(chat_id)

        return


    # =====================================================
    # SEARCH BUTTON
    # =====================================================

    if text == "🔎 جستجوی کالا":

        send_message(

            chat_id,

            "🔎 کد یا نام کالا را ارسال کنید.\n\n"

            "مثال:\n"
            "100158\n\n"

            "یا:\n"
            "کابل دنا\n\n"

            "یا:\n"
            "دنا"
        )

        return


    # =====================================================
    # WEBSITE
    # =====================================================

    if text == "🌐 ورود به سایت":

        send_message(

            chat_id,

            "🌐 ورود مستقیم به سایت:\n\n"
            "https://www.tecnoyadakabbasi.ir"
        )

        return


    # =====================================================
    # DIRECT CALL
    # =====================================================

    if text == "📞 تماس مستقیم":

        send_message(

            chat_id,

            "📞 تماس مستقیم با ما:\n\n"
            "tel:09377700031\n\n"
            "شماره تماس: 09377700031"
        )

        return


    # =====================================================
    # PDF
    # =====================================================

    if text in PDF_GROUPS:

        group = PDF_GROUPS[text]

        send_message(
            chat_id,
            "⏳ در حال آماده‌سازی فایل PDF..."
        )

        pdf_path = find_pdf(group)

        if not pdf_path:

            send_message(
                chat_id,
                "❌ فایل PDF این گروه پیدا نشد."
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
                response.json().get("ok")
            ):

                return

        except Exception:

            pass

        send_message(
            chat_id,
            "❌ ارسال فایل PDF انجام نشد."
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

        results = search_excel(text)

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
        "PDF + SEARCH + WEBSITE + CALL"
    )

    print(
        "PRICE UNIT: RIAL"
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

            time.sleep(5)


# =========================================================
# RENDER
# =========================================================

@app.route("/")
def home():

    return (
        "Bale Bot is running - "
        "Prices in Rial"
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
