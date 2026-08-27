import os
import re
import time
import threading
import requests
from flask import Flask
from openpyxl import load_workbook
from urllib.parse import unquote
from datetime import datetime

TOKEN = os.getenv("BALE_TOKEN")
BASE_URL = f"https://tapi.bale.ai/bot{TOKEN}" if TOKEN else ""

app = Flask(__name__)

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
# USER DATA
# =========================================================

CARTS = {}
USER_STATES = {}
ORDER_COUNTER = 1000


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
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = text.lower()
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def compact(text):
    return re.sub(
        r"[^0-9a-zآ-ی]+",
        "",
        normalize(text)
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
        response = requests.post(
            f"{BASE_URL}/sendMessage",
            json=data,
            timeout=30
        )

        print("MESSAGE:", response.status_code)

        return response

    except Exception as e:
        print("MESSAGE ERROR:", e)
        return None


# =========================================================
# MAIN MENU
# =========================================================

def main_menu(chat_id):

    keyboard = [

        ["📞 تماس مستقیم", "🌐 ورود به سایت"],

        ["📄 دریافت لیست قیمت"],

        ["🔎 جستجوی کالا"],

        ["🛒 سبد خرید", "🧾 ثبت سفارش"],

        ["📦 سوکت عباسی", "🔌 کابل تکنو سبزوار"],

        ["⚡ وایر عباسی", "🔩 مهره و سنسور"],

        ["💡 قطعات برقی خودرو", "🧩 خارجات و پلیمریجات"],

        ["🔌 کابل خودرو سبزوار", "⚙️ شیلنگ خودرو"],

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
            or normalized_name in wanted
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


def price_number(value):

    try:

        text = str(value or "")

        text = text.replace(",", "")
        text = text.replace("٬", "")

        text = normalize(text)

        number = float(text)

        return int(number)

    except Exception:
        return 0


# =========================================================
# FIND EXCEL COLUMNS
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
            or name == "کد"
            or "شناسه" in name
        ):
            columns["code"] = index

        elif (
            "نام کالا" in name
            or name == "نام"
        ):
            columns["name"] = index

        elif "قیمت" in name:
            columns["price"] = index

        elif (
            "توضیحات" in name
            or "توضیح" in name
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

    return all(
        word in full_text
        for word in words
    )


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

    print("SEARCH:", query)

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

                for row_number, row in enumerate(
                    sheet.iter_rows(values_only=True),
                    start=1
                ):

                    found = get_columns(row)

                    if (
                        "code" in found
                        and "name" in found
                        and "price" in found
                    ):

                        header_row = row_number
                        columns = found
                        break

                if not columns:
                    continue

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
                        and price_index < len(row)
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

    print(
        "RESULTS:",
        len(results)
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

        if len(message) + len(block) > 3500:

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
# ADD TO CART
# =========================================================

def add_to_cart(
    chat_id,
    item,
    quantity
):

    if chat_id not in CARTS:
        CARTS[chat_id] = []

    cart = CARTS[chat_id]

    code = item["code"]

    for cart_item in cart:

        if cart_item["code"] == code:

            cart_item["quantity"] += quantity
            return

    cart.append({

        "code": item["code"],
        "name": item["name"],
        "price": price_number(item["price"]),
        "quantity": quantity

    })


# =========================================================
# SHOW CART
# =========================================================

def show_cart(chat_id):

    cart = CARTS.get(
        chat_id,
        []
    )

    if not cart:

        send_message(
            chat_id,
            "🛒 سبد خرید شما خالی است."
        )

        return

    message = "🛒 سبد خرید شما:\n\n"

    total = 0

    for index, item in enumerate(
        cart,
        start=1
    ):

        item_total = (
            item["price"]
            *
            item["quantity"]
        )

        total += item_total

        message += (

            f"{index}. {item['name']}\n"

            f"🔢 کد: {item['code']}\n"

            f"📦 تعداد: {item['quantity']}\n"

            f"💰 قیمت واحد: "
            f"{item['price']:,} ریال\n"

            f"💵 مبلغ: "
            f"{item_total:,} ریال\n"

            "──────────────\n"

        )

    message += (
        f"\n💰 جمع کل: "
        f"{total:,} ریال"
    )

    keyboard = [

        ["🧾 ثبت سفارش"],

        ["🗑️ خالی کردن سبد"],

        ["🔙 منوی اصلی"]

    ]

    send_message(
        chat_id,
        message,
        keyboard
    )


# =========================================================
# CLEAR CART
# =========================================================

def clear_cart(chat_id):

    CARTS[chat_id] = []

    USER_STATES.pop(
        chat_id,
        None
    )

    send_message(
        chat_id,
        "🗑️ سبد خرید خالی شد."
    )

    main_menu(chat_id)


# =========================================================
# START ORDER
# =========================================================

def start_order(chat_id):

    cart = CARTS.get(
        chat_id,
        []
    )

    if not cart:

        send_message(
            chat_id,
            "🛒 سبد خرید خالی است.\n"
            "ابتدا کالا به سبد اضافه کنید."
        )

        return

    USER_STATES[chat_id] = {
        "state": "name"
    }

    send_message(
        chat_id,
        "🧾 ثبت سفارش\n\n"
        "لطفاً نام و نام خانوادگی خود را وارد کنید:"
    )


# =========================================================
# ORDER NUMBER
# =========================================================

def get_order_number():

    global ORDER_COUNTER

    ORDER_COUNTER += 1

    return ORDER_COUNTER


# =========================================================
# FINISH ORDER
# =========================================================

def finish_order(
    chat_id,
    name,
    phone
):

    cart = CARTS.get(
        chat_id,
        []
    )

    if not cart:

        send_message(
            chat_id,
            "❌ سبد خرید خالی است."
        )

        return

    order_number = get_order_number()

    total = 0

    message = (

        "✅ سفارش شما ثبت شد.\n\n"

        f"🔢 شماره سفارش: "
        f"{order_number}\n"

        f"👤 نام: {name}\n"

        f"📞 تماس: {phone}\n\n"

        "📦 اقلام سفارش:\n"

    )

    for item in cart:

        item_total = (
            item["price"]
            *
            item["quantity"]
        )

        total += item_total

        message += (

            f"\n• {item['name']}\n"

            f"کد: {item['code']}\n"

            f"تعداد: {item['quantity']}\n"

            f"مبلغ: "
            f"{item_total:,} ریال\n"

        )

    message += (

        "\n────────────────\n"

        f"💰 مبلغ کل: "
        f"{total:,} ریال\n\n"

        "📞 جهت پیگیری سفارش:\n"
        "۰۹۳۷۷۷۰۰۰۳۱"

    )

    send_message(
        chat_id,
        message
    )

    CARTS[chat_id] = []

    USER_STATES.pop(
        chat_id,
        None
    )

    main_menu(chat_id)


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

    print(
        "USER MESSAGE:",
        text
    )

    # =====================================================
    # START
    # =====================================================

    if text == "/start":

        USER_STATES.pop(
            chat_id,
            None
        )

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
        "🔙 منوی اصلی",
        "منوی اصلی",
        "📄 دریافت لیست قیمت"
    ):

        USER_STATES.pop(
            chat_id,
            None
        )

        main_menu(chat_id)

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
            "شماره تماس: ۰۹۳۷۷۷۰۰۰۳۱"
        )

        return

    # =====================================================
    # SEARCH BUTTON
    # =====================================================

    if text == "🔎 جستجوی کالا":

        USER_STATES[chat_id] = {
            "state": "search"
        }

        send_message(
            chat_id,
            "🔎 کد یا نام کالا را ارسال کنید.\n\n"
            "مثال:\n"
            "100158\n\n"
            "یا:\n"
            "کابل دنا"
        )

        return

    # =====================================================
    # CART BUTTON
    # =====================================================

    if text == "🛒 سبد خرید":

        USER_STATES.pop(
            chat_id,
            None
        )

        show_cart(chat_id)

        return

    # =====================================================
    # CLEAR CART
    # =====================================================

    if text == "🗑️ خالی کردن سبد":

        clear_cart(chat_id)

        return

    # =====================================================
    # ORDER BUTTON
    # =====================================================

    if text == "🧾 ثبت سفارش":

        start_order(chat_id)

        return

    # =====================================================
    # ADD PRODUCT BUTTON
    # =====================================================
    # این بخش عمداً قبل از جستجوی عمومی قرار گرفته است.

    if text == "➕ افزودن به سبد":

        state_data = USER_STATES.get(
            chat_id
        )

        if not state_data:

            send_message(
                chat_id,
                "❌ ابتدا یک کالا را جستجو کنید."
            )

            return

        item = state_data.get(
            "item"
        )

        if not item:

            send_message(
                chat_id,
                "❌ کالا پیدا نشد."
            )

            return

        USER_STATES[chat_id] = {

            "state": "quantity",

            "item": item

        }

        send_message(
            chat_id,
            "🔢 تعداد مورد نظر را وارد کنید:\n\n"
            "مثلاً: 2"
        )

        return

    # =====================================================
    # ORDER STATES
    # =====================================================

    state_data = USER_STATES.get(
        chat_id
    )

    if state_data:

        state = state_data.get(
            "state"
        )

        # -----------------------------------------------
        # NAME
        # -----------------------------------------------

        if state == "name":

            if not text:

                send_message(
                    chat_id,
                    "❌ نام نمی‌تواند خالی باشد."
                )

                return

            state_data["name"] = text

            state_data["state"] = "phone"

            send_message(
                chat_id,
                "📞 لطفاً شماره تماس خود را وارد کنید:"
            )

            return

        # -----------------------------------------------
        # PHONE
        # -----------------------------------------------

        if state == "phone":

            phone = normalize(text)

            phone = re.sub(
                r"\D",
                "",
                phone
            )

            if len(phone) < 10:

                send_message(
                    chat_id,
                    "❌ شماره تماس صحیح نیست.\n"
                    "لطفاً دوباره وارد کنید:"
                )

                return

            finish_order(
                chat_id,
                state_data["name"],
                phone
            )

            return

        # -----------------------------------------------
        # QUANTITY
        # -----------------------------------------------

        if state == "quantity":

            if not text.isdigit():

                send_message(
                    chat_id,
                    "❌ تعداد باید یک عدد باشد.\n"
                    "مثلاً: 2"
                )

                return

            quantity = int(text)

            if quantity <= 0:

                send_message(
                    chat_id,
                    "❌ تعداد باید بیشتر از صفر باشد."
                )

                return

            item = state_data.get(
                "item"
            )

            if not item:

                send_message(
                    chat_id,
                    "❌ اطلاعات کالا پیدا نشد."
                )

                USER_STATES.pop(
                    chat_id,
                    None
                )

                return

            add_to_cart(
                chat_id,
                item,
                quantity
            )

            USER_STATES.pop(
                chat_id,
                None
            )

            send_message(
                chat_id,
                "✅ کالا با موفقیت به سبد خرید اضافه شد."
            )

            show_cart(chat_id)

            return

    # =====================================================
    # PDF
    # =====================================================

    if text in PDF_GROUPS:

        USER_STATES.pop(
            chat_id,
            None
        )

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
                and response.json().get("ok")
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

        USER_STATES.pop(
            chat_id,
            None
        )

        send_message(
            chat_id,
            "🔎 در حال جستجوی کامل..."
        )

        results = search_excel(
            text
        )

        if not results:

            send_message(
                chat_id,
                "❌ کالایی با این کد یا نام پیدا نشد."
            )

            return

        send_results(
            chat_id,
            results
        )

        # اگر فقط یک کالا پیدا شد،
        # دکمه افزودن به سبد نشان داده می‌شود.

        if len(results) == 1:

            item = results[0]

            USER_STATES[chat_id] = {

                "state": "add_product",

                "item": item

            }

            keyboard = [

                ["➕ افزودن به سبد"],

                ["🛒 سبد خرید"],

                ["🔙 منوی اصلی"]

            ]

            send_message(
                chat_id,
                "📦 آیا این کالا را به سبد خرید اضافه می‌کنید؟",
                keyboard
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
        "PDF + SEARCH + CART + ORDER"
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
                "BOT LOOP ERROR:",
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
        "PDF + Search + Cart + Order - "
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
