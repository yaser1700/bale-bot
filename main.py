import os
import re
import time
import threading
import requests

from flask import Flask, request, jsonify
from openpyxl import load_workbook
from urllib.parse import unquote


# =========================================================
# CONFIG
# =========================================================

TOKEN = os.getenv("BALE_TOKEN")

ADMIN_CHAT_ID = os.getenv(
    "BALE_ADMIN_ID",
    "570728574"
)

WEBSITE = "https://www.tecnoyadakabbasi.ir"
PHONE = "09377700031"

BASE_URL = (
    f"https://tapi.bale.ai/bot{TOKEN}"
    if TOKEN else ""
)

app = Flask(__name__)


# =========================================================
# MEMORY
# =========================================================

CARTS = {}
USER_STATES = {}

ORDER_COUNTER = 1000

LOCK = threading.RLock()


# =========================================================
# PDF GROUPS
# =========================================================

PDF_GROUPS = {

    "📦 سوکت عباسی":
        "سوکت عباسی",

    "🔌 کابل تکنو سبزوار":
        "کابل تکنو سبزوار",

    "⚡ وایر عباسی":
        "وایر عباسی",

    "🔩 مهره و سنسور":
        "مهره و سنسور",

    "💡 قطعات برقی خودرو":
        "قطعات برقی خودرو",

    "🧩 خارجات و پلیمریجات":
        "خارجات و پلیمریجات",

    "🔌 کابل خودرو سبزوار":
        "کابل خودرو سبزوار",

    "⚙️ شیلنگ خودرو":
        "شیلنگ خودرو",

    "⚙️ جلوبندی":
        "جلوبندی",
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

    return re.sub(
        r"[^0-9a-zآ-ی]+",
        "",
        normalize(text)
    )


# =========================================================
# SEND MESSAGE
# =========================================================

def send_message(
    chat_id,
    text,
    keyboard=None,
    inline=False
):

    if not TOKEN:
        print("TOKEN NOT FOUND")
        return None

    data = {

        "chat_id": chat_id,
        "text": str(text)
    }

    if keyboard:

        if inline:

            data["reply_markup"] = {

                "inline_keyboard":
                    keyboard
            }

        else:

            data["reply_markup"] = {

                "keyboard":
                    keyboard,

                "resize_keyboard":
                    True
            }

    try:

        response = requests.post(

            f"{BASE_URL}/sendMessage",

            json=data,

            timeout=30
        )

        print(
            "SEND MESSAGE:",
            response.status_code,
            response.text[:500]
        )

        return response

    except Exception as e:

        print(
            "SEND MESSAGE ERROR:",
            repr(e)
        )

        return None


# =========================================================
# CALLBACK ANSWER
# =========================================================

def answer_callback(
    callback_id,
    text=""
):

    if not callback_id:
        return

    try:

        response = requests.post(

            f"{BASE_URL}/answerCallbackQuery",

            json={

                "callback_query_id":
                    callback_id,

                "text":
                    text
            },

            timeout=20
        )

        print(
            "CALLBACK ANSWER:",
            response.status_code
        )

    except Exception as e:

        print(
            "CALLBACK ERROR:",
            repr(e)
        )


# =========================================================
# MAIN MENU
# =========================================================

def main_menu(chat_id):

    keyboard = [

        [
            "📞 تماس مستقیم",
            "🌐 ورود به سایت"
        ],

        [
            "📄 دریافت لیست قیمت"
        ],

        [
            "🔎 جستجوی کالا"
        ],

        [
            "🛒 سبد خرید",
            "🧾 ثبت سفارش"
        ],

        [
            "📦 سوکت عباسی",
            "🔌 کابل تکنو سبزوار"
        ],

        [
            "⚡ وایر عباسی",
            "🔩 مهره و سنسور"
        ],

        [
            "💡 قطعات برقی خودرو",
            "🧩 خارجات و پلیمریجات"
        ],

        [
            "🔌 کابل خودرو سبزوار",
            "⚙️ شیلنگ خودرو"
        ],

        [
            "⚙️ جلوبندی"
        ]
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

        print(
            "LIST FILE ERROR:",
            repr(e)
        )

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

                    "chat_id":
                        str(chat_id),

                    "caption":
                        caption
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
            "SEND PDF:",
            response.status_code,
            response.text[:300]
        )

        return response

    except Exception as e:

        print(
            "PDF ERROR:",
            repr(e)
        )

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

        return int(
            float(text)
        )

    except Exception:

        return 0


# =========================================================
# EXCEL COLUMNS
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

    full_compact = compact(
        full_text
    )

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

    try:

        files = os.listdir(base)

    except Exception as e:

        print(
            "FILES ERROR:",
            repr(e)
        )

        return []

    excel_files = [

        filename

        for filename in files

        if filename.lower().endswith(".xlsx")

        and not filename.startswith("~$")
    ]

    results = []

    seen = set()

    print(
        "================================"
    )

    print(
        "SEARCH:",
        query
    )

    for filename in excel_files:

        path = os.path.join(
            base,
            filename
        )

        workbook = None

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

                    sheet.iter_rows(
                        values_only=True
                    ),

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
                    description = get_value(
                        "description"
                    )

                    price = ""

                    price_index = columns.get(
                        "price"
                    )

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

                        "code":
                            code,

                        "name":
                            name,

                        "price":
                            price,

                        "group":
                            group,

                        "description":
                            description
                    })

            workbook.close()

        except Exception as e:

            print(
                "EXCEL ERROR:",
                filename,
                repr(e)
            )

            try:

                if workbook:
                    workbook.close()

            except Exception:
                pass

    print(
        "RESULT COUNT:",
        len(results)
    )

    print(
        "================================"
    )

    return results


# =========================================================
# SHOW SEARCH RESULTS
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

    # ذخیره نتایج برای تمام دکمه‌ها
    USER_STATES[chat_id] = {

        "state":
            "search_results",

        "results":
            results
    }

    send_message(

        chat_id,

        f"🔎 {len(results)} کالا پیدا شد:\n\n"
        "برای هر کالا می‌توانید جداگانه "
        "«افزودن به سبد خرید» را بزنید."
    )

    for index, item in enumerate(results):

        message = (

            f"📦 کد کالا: {item['code']}\n"
            f"📝 نام کالا: {item['name']}\n"
            f"💰 قیمت: {item['price']} ریال\n"
            f"📁 گروه: {item['group']}"
        )

        if item["description"]:

            message += (

                f"\nℹ️ توضیحات: "
                f"{item['description']}"
            )

        callback_data = (
            f"ADD:{index}"
        )

        keyboard = [[

            {
                "text":
                    "➕ افزودن به سبد خرید",

                "callback_data":
                    callback_data
            }

        ]]

        send_message(

            chat_id,

            message,

            keyboard,

            inline=True
        )


# =========================================================
# ADD TO CART
# =========================================================

def add_to_cart(
    chat_id,
    item,
    quantity
):

    with LOCK:

        if chat_id not in CARTS:

            CARTS[chat_id] = []

        cart = CARTS[chat_id]

        item_code = normalize(
            item["code"]
        )

        for cart_item in cart:

            if normalize(
                cart_item["code"]
            ) == item_code:

                cart_item["quantity"] += quantity

                return

        cart.append({

            "code":
                item["code"],

            "name":
                item["name"],

            "price":
                price_number(
                    item["price"]
                ),

            "quantity":
                quantity
        })


# =========================================================
# SHOW CART
# =========================================================

def show_cart(chat_id):

    with LOCK:

        cart = list(
            CARTS.get(
                chat_id,
                []
            )
        )

    if not cart:

        send_message(

            chat_id,

            "🛒 سبد خرید شما خالی است."
        )

        return

    message = (
        "🛒 سبد خرید شما:\n\n"
    )

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

            f"📦 تعداد: "
            f"{item['quantity']}\n"

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

        ["🔎 جستجوی کالا"],

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

    with LOCK:

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

    with LOCK:

        cart = list(
            CARTS.get(
                chat_id,
                []
            )
        )

    if not cart:

        send_message(

            chat_id,

            "🛒 سبد خرید خالی است.\n"
            "ابتدا کالا به سبد اضافه کنید."
        )

        return

    USER_STATES[chat_id] = {

        "state":
            "order_name"
    }

    send_message(

        chat_id,

        "🧾 ثبت سفارش\n\n"
        "لطفاً نام و نام خانوادگی خود را وارد کنید:"
    )


# =========================================================
# FINISH ORDER
# =========================================================

def finish_order(
    chat_id,
    name,
    phone
):

    global ORDER_COUNTER

    with LOCK:

        cart = list(
            CARTS.get(
                chat_id,
                []
            )
        )

        if not cart:
            send_message(
                chat_id,
                "❌ سبد خرید خالی است."
            )
            return

        ORDER_COUNTER += 1

        order_number = ORDER_COUNTER

    total = 0

    customer_message = (

        "✅ سفارش شما با موفقیت ثبت شد.\n\n"

        f"🔢 شماره سفارش: "
        f"{order_number}\n"

        f"👤 نام: {name}\n"

        f"📞 تماس: {phone}\n\n"

        "📦 اقلام سفارش:\n"
    )

    admin_message = (

        "🚨 سفارش جدید دریافت شد 🚨\n\n"

        f"🔢 شماره سفارش: "
        f"{order_number}\n"

        f"👤 نام مشتری: {name}\n"

        f"📞 شماره تماس: {phone}\n"

        f"🆔 شناسه کاربر: {chat_id}\n\n"

        "📦 اقلام سفارش:\n"
    )

    for item in cart:

        item_total = (

            item["price"]
            *
            item["quantity"]
        )

        total += item_total

        customer_message += (

            f"\n• {item['name']}\n"
            f"کد: {item['code']}\n"
            f"تعداد: {item['quantity']}\n"
            f"مبلغ: {item_total:,} ریال\n"
        )

        admin_message += (

            f"\n• {item['name']}\n"
            f"کد: {item['code']}\n"
            f"تعداد: {item['quantity']}\n"
            f"قیمت واحد: "
            f"{item['price']:,} ریال\n"
            f"مبلغ: "
            f"{item_total:,} ریال\n"
        )

    customer_message += (

        "\n────────────────\n"

        f"💰 مبلغ کل: "
        f"{total:,} ریال\n\n"

        "📞 جهت پیگیری:\n"
        "۰۹۳۷۷۷۰۰۰۳۱"
    )

    admin_message += (

        "\n────────────────\n"

        f"💰 مبلغ کل سفارش: "
        f"{total:,} ریال\n\n"

        "📞 شماره پیگیری مشتری:\n"
        f"{phone}"
    )

    # -----------------------------------------------------
    # SEND TO CUSTOMER
    # -----------------------------------------------------

    customer_response = send_message(

        chat_id,

        customer_message
    )

    # -----------------------------------------------------
    # SEND TO ADMIN
    # -----------------------------------------------------

    print(
        "SENDING ORDER TO ADMIN:",
        ADMIN_CHAT_ID
    )

    admin_response = send_message(

        ADMIN_CHAT_ID,

        admin_message
    )

    # -----------------------------------------------------
    # LOG
    # -----------------------------------------------------

    try:

        customer_ok = (

            customer_response is not None
            and
            customer_response.json().get("ok")
        )

    except Exception:

        customer_ok = False

    try:

        admin_ok = (

            admin_response is not None
            and
            admin_response.json().get("ok")
        )

    except Exception:

        admin_ok = False

    print(
        "ORDER:",
        order_number,
        "CUSTOMER:",
        customer_ok,
        "ADMIN:",
        admin_ok
    )

    if not admin_ok:

        print(
            "WARNING: ORDER WAS NOT SENT TO ADMIN"
        )

        print(
            "ADMIN CHAT ID:",
            ADMIN_CHAT_ID
        )

    # -----------------------------------------------------
    # CLEAR USER CART
    # -----------------------------------------------------

    with LOCK:

        CARTS[chat_id] = []

        USER_STATES.pop(
            chat_id,
            None
        )

    main_menu(chat_id)


# =========================================================
# CALLBACK
# =========================================================

def process_callback(callback):

    callback_id = callback.get(
        "id"
    )

    data = str(
        callback.get("data") or ""
    )

    message = (
        callback.get("message")
        or {}
    )

    chat = (
        message.get("chat")
        or {}
    )

    chat_id = chat.get(
        "id"
    )

    if not chat_id:
        return

    print(
        "CALLBACK:",
        data,
        "CHAT:",
        chat_id
    )

    # =====================================================
    # ADD PRODUCT
    # =====================================================

    if data.startswith("ADD:"):

        parts = data.split(":")

        if len(parts) != 2:

            answer_callback(

                callback_id,

                "❌ خطا در انتخاب کالا"
            )

            return

        try:

            index = int(
                parts[1]
            )

        except Exception:

            answer_callback(

                callback_id,

                "❌ خطا در انتخاب کالا"
            )

            return

        state_data = USER_STATES.get(
            chat_id
        )

        if not state_data:

            answer_callback(

                callback_id,

                "❌ نتایج جستجو منقضی شده است."
            )

            return

        results = state_data.get(
            "results",
            []
        )

        if (

            index < 0
            or
            index >= len(results)

        ):

            answer_callback(

                callback_id,

                "❌ کالا پیدا نشد."
            )

            return

        item = results[index]

        # نتایج جستجو را نگه می‌داریم
        # تا بعد از ثبت تعداد، کالاهای دیگر
        # هم قابل انتخاب باشند.

        USER_STATES[chat_id] = {

            "state":
                "quantity",

            "item":
                item,

            "results":
                results
        }

        answer_callback(

            callback_id,

            "✅ کالا انتخاب شد"
        )

        send_message(

            chat_id,

            f"📦 {item['name']}\n\n"

            f"🔢 کد کالا: "
            f"{item['code']}\n"

            f"💰 قیمت: "
            f"{item['price']} ریال\n\n"

            "🔢 تعداد مورد نظر را وارد کنید:"
        )

        return

    answer_callback(
        callback_id
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
        "USER MESSAGE:",
        chat_id,
        text
    )

    # =====================================================
    # START
    # =====================================================

    if text == "/start":

        CARTS.setdefault(
            chat_id,
            []
        )

        USER_STATES.pop(
            chat_id,
            None
        )

        send_message(

            chat_id,

            "سلام 👋\n\n"
            "به ربات تولیدی و بازرگانی عباسی خوش آمدید."
        )

        time.sleep(0.2)

        main_menu(chat_id)

        return

    # =====================================================
    # WEBSITE
    # =====================================================

    if text == "🌐 ورود به سایت":

        send_message(

            chat_id,

            "🌐 ورود مستقیم به سایت:\n\n"
            + WEBSITE
        )

        return

    # =====================================================
    # CALL
    # =====================================================

    if text == "📞 تماس مستقیم":

        send_message(

            chat_id,

            "📞 تماس مستقیم:\n\n"
            "۰۹۳۷۷۷۰۰۰۳۱\n\n"
            "برای تماس روی شماره بالا بزنید."
        )

        return

    # =====================================================
    # MAIN MENU
    # =====================================================

    if text in (

        "🔙 منوی اصلی",
        "منوی اصلی"

    ):

        USER_STATES.pop(
            chat_id,
            None
        )

        main_menu(chat_id)

        return

    # =====================================================
    # SEARCH
    # =====================================================

    if text == "🔎 جستجوی کالا":

        USER_STATES[chat_id] = {

            "state":
                "search"
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
    # CART
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
    # ORDER
    # =====================================================

    if text == "🧾 ثبت سفارش":

        start_order(chat_id)

        return

    # =====================================================
    # USER STATE
    # =====================================================

    state_data = USER_STATES.get(
        chat_id
    )

    if state_data:

        state = state_data.get(
            "state"
        )

        # =================================================
        # SEARCH STATE
        # =================================================

        if state == "search":

            if not text:

                send_message(

                    chat_id,

                    "❌ عبارت جستجو را وارد کنید."
                )

                return

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

        # =================================================
        # QUANTITY
        # =================================================

        if state == "quantity":

            number_text = normalize(
                text
            )

            if not number_text.isdigit():

                send_message(

                    chat_id,

                    "❌ تعداد باید عدد باشد.\n"
                    "مثلاً: 2"
                )

                return

            quantity = int(
                number_text
            )

            if quantity <= 0:

                send_message(

                    chat_id,

                    "❌ تعداد باید بیشتر از صفر باشد."
                )

                return

            item = state_data.get(
                "item"
            )

            results = state_data.get(
                "results",
                []
            )

            if not item:

                USER_STATES.pop(
                    chat_id,
                    None
                )

                send_message(

                    chat_id,

                    "❌ اطلاعات کالا پیدا نشد."
                )

                return

            add_to_cart(

                chat_id,

                item,

                quantity
            )

            send_message(

                chat_id,

                f"✅ {item['name']}\n"
                f"به تعداد {quantity} عدد "
                "به سبد خرید اضافه شد. 🛒"
            )

            # مهم:
            # نتایج جستجو دوباره فعال می‌شوند
            # تا بتواند کالای دوم، سوم و ... را اضافه کند.

            USER_STATES[chat_id] = {

                "state":
                    "search_results",

                "results":
                    results
            }

            keyboard = [

                [
                    "🛒 سبد خرید"
                ],

                [
                    "🔎 جستجوی کالا"
                ],

                [
                    "🧾 ثبت سفارش"
                ],

                [
                    "🔙 منوی اصلی"
                ]
            ]

            send_message(

                chat_id,

                "👇 می‌توانید کالای دیگری هم اضافه کنید "
                "یا سبد خرید را مشاهده کنید.",

                keyboard
            )

            return

        # =================================================
        # ORDER NAME
        # =================================================

        if state == "order_name":

            if not text:

                send_message(

                    chat_id,

                    "❌ نام را وارد کنید."
                )

                return

            state_data["name"] = text

            state_data["state"] = (
                "order_phone"
            )

            send_message(

                chat_id,

                "📞 لطفاً شماره تماس خود را وارد کنید:"
            )

            return

        # =================================================
        # ORDER PHONE
        # =================================================

        if state == "order_phone":

            phone = normalize(
                text
            )

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

        pdf_path = find_pdf(
            group
        )

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
    # GENERAL SEARCH
    # =====================================================

    if text:

        USER_STATES[chat_id] = {

            "state":
                "search"
        }

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
# WEBHOOK
# =========================================================

@app.route(
    "/webhook",
    methods=["POST"]
)
def webhook():

    try:

        update = request.get_json(
            silent=True
        )

        print(
            "================================"
        )

        print(
            "WEBHOOK UPDATE RECEIVED"
        )

        print(
            update
        )

        print(
            "================================"
        )

        if not update:

            return jsonify({
                "ok": True
            })

        # callback button
        if update.get(
            "callback_query"
        ):

            threading.Thread(

                target=process_callback,

                args=(
                    update[
                        "callback_query"
                    ],
                ),

                daemon=True

            ).start()

        # normal message
        elif update.get(
            "message"
        ):

            threading.Thread(

                target=process_message,

                args=(
                    update[
                        "message"
                    ],
                ),

                daemon=True

            ).start()

        return jsonify({
            "ok": True
        })

    except Exception as e:

        print(
            "WEBHOOK ERROR:",
            repr(e)
        )

        return jsonify({
            "ok": True
        })


# =========================================================
# HEALTH
# =========================================================

@app.route("/")
def home():

    return "Bale Bot is running"


@app.route("/health")
def health():

    return jsonify({

        "status":
            "ok",

        "webhook":
            True,

        "admin_configured":
            bool(ADMIN_CHAT_ID)
    })


# =========================================================
# SET WEBHOOK
# =========================================================

def setup_webhook():

    if not TOKEN:

        print(
            "❌ BALE_TOKEN NOT FOUND"
        )

        return

    # Render service URL
    render_url = os.getenv(
        "RENDER_EXTERNAL_URL"
    )

    if not render_url:

        print(
            "⚠️ RENDER_EXTERNAL_URL NOT FOUND"
        )

        return

    webhook_url = (
        render_url.rstrip("/")
        + "/webhook"
    )

    print(
        "SETTING WEBHOOK:"
    )

    print(
        webhook_url
    )

    try:

        response = requests.post(

            f"{BASE_URL}/setWebhook",

            json={
                "url":
                    webhook_url
            },

            timeout=30
        )

        print(
            "SET WEBHOOK:",
            response.status_code,
            response.text
        )

    except Exception as e:

        print(
            "SET WEBHOOK ERROR:",
            repr(e)
        )


# =========================================================
# START
# =========================================================

if __name__ == "__main__":

    print(
        "========================================"
    )

    print(
        "BALE BOT FINAL VERSION"
    )

    print(
        "SEARCH + PDF + CART + ORDER"
    )

    print(
        "MULTI PRODUCT CART"
    )

    print(
        "ADMIN ORDER NOTIFICATION"
    )

    print(
        "WEBHOOK MODE"
    )

    print(
        "ADMIN CHAT ID:",
        ADMIN_CHAT_ID
    )

    print(
        "========================================"
    )

    # ابتدا Flask را بالا می‌آوریم
    # سپس webhook را تنظیم می‌کنیم.

    def startup():

        time.sleep(3)

        setup_webhook()

    threading.Thread(

        target=startup,

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

        port=port,

        threaded=True
    )
