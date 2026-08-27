import os
import re
import threading
import requests
from flask import Flask, request, jsonify
from openpyxl import load_workbook
from urllib.parse import unquote


# =========================================================
# CONFIG
# =========================================================

TOKEN = os.getenv("BALE_TOKEN", "").strip()

BASE_URL = (
    f"https://tapi.bale.ai/bot{TOKEN}"
    if TOKEN
    else ""
)

PHONE = "09377700031"

WEBSITE = "https://www.tecnoyadakabbasi.ir"

# آیدی مدیر
ADMIN_CHAT_ID = int(
    os.getenv(
        "ADMIN_CHAT_ID",
        "570728574"
    )
)

# آدرس عمومی Render
PUBLIC_URL = os.getenv(
    "WEBHOOK_URL",
    "https://bale-bot-vqup.onrender.com"
).rstrip("/")

WEBHOOK_PATH = "/webhook"

WEBHOOK_URL = PUBLIC_URL + WEBHOOK_PATH


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
        "\u202a": "",
        "\u202b": "",
        "\u202c": "",
    }

    for old, new in replacements.items():
        text = text.replace(
            old,
            new
        )

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
# SAFE API REQUEST
# =========================================================

def api_post(
    method,
    data=None,
    files=None,
    timeout=30
):

    if not TOKEN:

        print(
            "ERROR: BALE_TOKEN is missing"
        )

        return None

    url = f"{BASE_URL}/{method}"

    try:

        response = requests.post(
            url,
            json=data if files is None else None,
            data=data if files is not None else None,
            files=files,
            timeout=timeout
        )

        print(
            "API:",
            method,
            response.status_code,
            response.text[:300]
        )

        return response

    except Exception as e:

        print(
            "API ERROR:",
            method,
            e
        )

        return None


# =========================================================
# SEND MESSAGE
# =========================================================

def send_message(
    chat_id,
    text,
    keyboard=None,
    inline=False
):

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
                    True,

                "one_time_keyboard":
                    False
            }

    return api_post(
        "sendMessage",
        data=data,
        timeout=30
    )


# =========================================================
# CALLBACK ANSWER
# =========================================================

def answer_callback(
    callback_id,
    text=""
):

    if not callback_id:
        return

    api_post(
        "answerCallbackQuery",
        data={

            "callback_query_id":
                callback_id,

            "text":
                text
        },
        timeout=20
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

    wanted = compact(
        group_name
    )

    try:

        files = os.listdir(base)

    except Exception as e:

        print(
            "PDF LIST ERROR:",
            e
        )

        return None

    for filename in files:

        if not filename.lower().endswith(
            ".pdf"
        ):
            continue

        decoded = unquote(
            filename
        )

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
            "PDF RESPONSE:",
            response.status_code,
            response.text[:300]
        )

        return response

    except Exception as e:

        print(
            "PDF ERROR:",
            e
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

    text = str(
        value
    ).strip()

    text = text.replace(
        ",",
        ""
    )

    text = text.replace(
        "٬",
        ""
    )

    try:

        number = float(
            text
        )

        if number.is_integer():

            return f"{int(number):,}"

    except Exception:
        pass

    return text


def price_number(value):

    try:

        text = str(
            value or ""
        )

        text = text.replace(
            ",",
            ""
        )

        text = text.replace(
            "٬",
            ""
        )

        text = normalize(
            text
        )

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

    for index, value in enumerate(
        row
    ):

        if value is None:
            continue

        name = normalize(
            value
        )

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

    q = normalize(
        query
    )

    qc = compact(
        query
    )

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

    if (
        qc
        and
        qc in full_compact
    ):

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

        excel_files = [

            filename

            for filename
            in os.listdir(base)

            if filename.lower().endswith(
                ".xlsx"
            )

            and not filename.startswith(
                "~$"
            )
        ]

    except Exception as e:

        print(
            "EXCEL LIST ERROR:",
            e
        )

        return []

    results = []

    seen = set()

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

                for (
                    row_number,
                    row
                ) in enumerate(

                    sheet.iter_rows(
                        values_only=True
                    ),

                    start=1
                ):

                    found = get_columns(
                        row
                    )

                    if (
                        "code" in found
                        and
                        "name" in found
                        and
                        "price" in found
                    ):

                        header_row = (
                            row_number
                        )

                        columns = found

                        break

                if not columns:
                    continue

                for row in sheet.iter_rows(

                    min_row=
                        header_row + 1,

                    values_only=True
                ):

                    if not row:
                        continue

                    def get_value(key):

                        index = columns.get(
                            key
                        )

                        if index is None:
                            return ""

                        if index >= len(
                            row
                        ):
                            return ""

                        return str(
                            row[index] or ""
                        ).strip()

                    group = get_value(
                        "group"
                    )

                    code = get_value(
                        "code"
                    )

                    name = get_value(
                        "name"
                    )

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
                            row[
                                price_index
                            ]
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

                    seen.add(
                        key
                    )

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

        except Exception as e:

            print(
                "EXCEL ERROR:",
                filename,
                e
            )

        finally:

            try:

                if workbook:
                    workbook.close()

            except Exception:
                pass

    print(
        "RESULT COUNT:",
        len(results)
    )

    return results


# =========================================================
# SEARCH RESULTS
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

    # ذخیره نتایج
    with LOCK:

        USER_STATES[chat_id] = {

            "state":
                "search_results",

            "results":
                results
        }

    send_message(

        chat_id,

        f"🔎 {len(results)} کالا پیدا شد:"
    )

    for index, item in enumerate(
        results
    ):

        message = (

            f"📦 کد کالا: "
            f"{item['code']}\n"

            f"📝 نام کالا: "
            f"{item['name']}\n"

            f"💰 قیمت: "
            f"{item['price']} ریال\n"

            f"📁 گروه: "
            f"{item['group']}"
        )

        if item.get(
            "description"
        ):

            message += (

                "\nℹ️ توضیحات: "

                + item[
                    "description"
                ]
            )

        callback_data = (

            f"ADD:{chat_id}:{index}"
        )

        keyboard = [[{

            "text":
                "➕ افزودن به سبد خرید",

            "callback_data":
                callback_data
        }]]

        send_message(

            chat_id,

            message,

            keyboard,

            inline=True
        )


# =========================================================
# CART
# =========================================================

def add_to_cart(
    chat_id,
    item,
    quantity
):

    with LOCK:

        if chat_id not in CARTS:

            CARTS[chat_id] = []

        cart = CARTS[
            chat_id
        ]

        # اول بر اساس کد کالا
        for cart_item in cart:

            if (
                cart_item["code"]
                ==
                item["code"]
            ):

                cart_item[
                    "quantity"
                ] += quantity

                return

        # کالای جدید
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

            f"{index}. "
            f"{item['name']}\n"

            f"🔢 کد: "
            f"{item['code']}\n"

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

        [
            "🧾 ثبت سفارش"
        ],

        [
            "🗑️ خالی کردن سبد"
        ],

        [
            "🔎 جستجوی کالا"
        ],

        [
            "🔙 منوی اصلی"
        ]
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

    main_menu(
        chat_id
    )


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

    with LOCK:

        USER_STATES[
            chat_id
        ] = {

            "state":
                "order_name"
        }

    send_message(

        chat_id,

        "🧾 ثبت سفارش\n\n"
        "لطفاً نام و نام خانوادگی خود را وارد کنید:"
    )


# =========================================================
# ADMIN ORDER MESSAGE
# =========================================================

def send_order_to_admin(
    order_number,
    chat_id,
    name,
    phone,
    cart,
    total
):

    message = (

        "🔔 سفارش جدید\n\n"

        f"🔢 شماره سفارش: "
        f"{order_number}\n"

        f"👤 نام مشتری: "
        f"{name}\n"

        f"📞 شماره تماس: "
        f"{phone}\n"

        f"🆔 آیدی چت: "
        f"{chat_id}\n\n"

        "📦 اقلام سفارش:\n"
    )

    for item in cart:

        item_total = (

            item["price"]

            *

            item["quantity"]
        )

        message += (

            f"\n• {item['name']}\n"

            f"کد: "
            f"{item['code']}\n"

            f"تعداد: "
            f"{item['quantity']}\n"

            f"قیمت واحد: "
            f"{item['price']:,} ریال\n"

            f"مبلغ: "
            f"{item_total:,} ریال\n"
        )

    message += (

        "\n────────────────\n"

        f"💰 مبلغ کل: "
        f"{total:,} ریال"
    )

    response = send_message(

        ADMIN_CHAT_ID,

        message
    )

    if response is None:

        print(
            "ADMIN MESSAGE FAILED"
        )

    else:

        try:

            if not response.json().get(
                "ok"
            ):

                print(
                    "ADMIN API ERROR:",
                    response.text
                )

        except Exception:
            pass


# =========================================================
# FINISH ORDER
# =========================================================

def finish_order(
    chat_id,
    name,
    phone
):

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

    global ORDER_COUNTER

    with LOCK:

        ORDER_COUNTER += 1

        order_number = (
            ORDER_COUNTER
        )

    total = 0

    message = (

        "✅ سفارش شما با موفقیت ثبت شد.\n\n"

        f"🔢 شماره سفارش: "
        f"{order_number}\n"

        f"👤 نام: "
        f"{name}\n"

        f"📞 تماس: "
        f"{phone}\n\n"

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

            f"کد: "
            f"{item['code']}\n"

            f"تعداد: "
            f"{item['quantity']}\n"

            f"مبلغ: "
            f"{item_total:,} ریال\n"
        )

    message += (

        "\n────────────────\n"

        f"💰 مبلغ کل: "
        f"{total:,} ریال\n\n"

        "📞 جهت پیگیری:\n"
        "۰۹۳۷۷۷۰۰۰۳۱"
    )

    # ابتدا برای مشتری
    send_message(

        chat_id,

        message
    )

    # سپس برای مدیر
    try:

        send_order_to_admin(

            order_number,

            chat_id,

            name,

            phone,

            cart,

            total
        )

    except Exception as e:

        print(
            "ADMIN ORDER ERROR:",
            e
        )

    # پاک کردن سبد
    with LOCK:

        CARTS[chat_id] = []

        USER_STATES.pop(
            chat_id,
            None
        )

    main_menu(
        chat_id
    )


# =========================================================
# CALLBACK
# =========================================================

def process_callback(
    callback
):

    callback_id = callback.get(
        "id"
    )

    data = str(
        callback.get(
            "data"
        ) or ""
    )

    message = (
        callback.get(
            "message"
        )
        or {}
    )

    chat = (
        message.get(
            "chat"
        )
        or {}
    )

    chat_id = chat.get(
        "id"
    )

    if not chat_id:

        answer_callback(

            callback_id,

            "خطا در شناسایی کاربر"
        )

        return

    # =====================================================
    # ADD PRODUCT
    # =====================================================

    if data.startswith(
        "ADD:"
    ):

        parts = data.split(
            ":"
        )

        if len(parts) != 3:

            answer_callback(

                callback_id,

                "خطا در انتخاب کالا"
            )

            return

        try:

            result_chat_id = int(
                parts[1]
            )

            index = int(
                parts[2]
            )

        except Exception:

            answer_callback(

                callback_id,

                "خطا در انتخاب کالا"
            )

            return

        if (
            result_chat_id
            !=
            chat_id
        ):

            answer_callback(

                callback_id,

                "این دکمه برای این کاربر نیست."
            )

            return

        with LOCK:

            state_data = USER_STATES.get(
                chat_id
            )

            if not state_data:

                answer_callback(

                    callback_id,

                    "نتایج جستجو منقضی شده است."
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

                    "کالا پیدا نشد."
                )

                return

            item = results[
                index
            ]

            USER_STATES[
                chat_id
            ] = {

                "state":
                    "quantity",

                "item":
                    item
            }

        # پاسخ سریع به Bale
        answer_callback(

            callback_id,

            "کالا انتخاب شد ✅"
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

def process_message(
    message
):

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
        "USER MESSAGE:",
        chat_id,
        text
    )

    # =====================================================
    # START
    # =====================================================

    if text == "/start":

        with LOCK:

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

        main_menu(
            chat_id
        )

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

        with LOCK:

            USER_STATES.pop(
                chat_id,
                None
            )

        main_menu(
            chat_id
        )

        return

    # =====================================================
    # SEARCH
    # =====================================================

    if text == "🔎 جستجوی کالا":

        with LOCK:

            USER_STATES[
                chat_id
            ] = {

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

        with LOCK:

            USER_STATES.pop(
                chat_id,
                None
            )

        show_cart(
            chat_id
        )

        return

    # =====================================================
    # CLEAR CART
    # =====================================================

    if text == "🗑️ خالی کردن سبد":

        clear_cart(
            chat_id
        )

        return

    # =====================================================
    # ORDER
    # =====================================================

    if text == "🧾 ثبت سفارش":

        start_order(
            chat_id
        )

        return

    # =====================================================
    # STATE
    # =====================================================

    with LOCK:

        state_data = USER_STATES.get(
            chat_id
        )

        if state_data:

            state_data = dict(
                state_data
            )

    if state_data:

        state = state_data.get(
            "state"
        )

        # -------------------------------------------------
        # SEARCH
        # -------------------------------------------------

        if state == "search":

            send_message(

                chat_id,

                "⏳ در حال جستجوی کالا..."
            )

            try:

                results = search_excel(
                    text
                )

                send_search_results(

                    chat_id,

                    results
                )

            except Exception as e:

                print(
                    "SEARCH ERROR:",
                    e
                )

                send_message(

                    chat_id,

                    "❌ هنگام جستجو خطایی رخ داد.\n"
                    "لطفاً دوباره امتحان کنید."
                )

            return

        # -------------------------------------------------
        # ORDER NAME
        # -------------------------------------------------

        if state == "order_name":

            if not text:

                send_message(

                    chat_id,

                    "❌ نام را وارد کنید."
                )

                return

            with LOCK:

                USER_STATES[
                    chat_id
                ]["name"] = text

                USER_STATES[
                    chat_id
                ]["state"] = (
                    "order_phone"
                )

            send_message(

                chat_id,

                "📞 لطفاً شماره تماس خود را وارد کنید:"
            )

            return

        # -------------------------------------------------
        # ORDER PHONE
        # -------------------------------------------------

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

            name = state_data.get(
                "name",
                ""
            )

            finish_order(

                chat_id,

                name,

                phone
            )

            return

        # -------------------------------------------------
        # QUANTITY
        # -------------------------------------------------

        if state == "quantity":

            if not re.fullmatch(
                r"\d+",
                text
            ):

                send_message(

                    chat_id,

                    "❌ تعداد باید عدد باشد.\n"
                    "مثلاً: 2"
                )

                return

            quantity = int(
                text
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

            if not item:

                send_message(

                    chat_id,

                    "❌ اطلاعات کالا پیدا نشد."
                )

                with LOCK:

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

            with LOCK:

                USER_STATES.pop(
                    chat_id,
                    None
                )

            send_message(

                chat_id,

                "✅ کالا با موفقیت به سبد خرید اضافه شد."
            )

            show_cart(
                chat_id
            )

            return

    # =====================================================
    # PDF
    # =====================================================

    if text in PDF_GROUPS:

        with LOCK:

            USER_STATES.pop(
                chat_id,
                None
            )

        group = PDF_GROUPS[
            text
        ]

        send_message(

            chat_id,

            "⏳ در حال آماده‌سازی فایل PDF..."
        )

        try:

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
                    response
                    and
                    response.json().get(
                        "ok"
                    )
                ):

                    return

            except Exception:
                pass

            send_message(

                chat_id,

                "❌ ارسال فایل PDF انجام نشد."
            )

        except Exception as e:

            print(
                "PDF PROCESS ERROR:",
                e
            )

            send_message(

                chat_id,

                "❌ هنگام ارسال فایل خطایی رخ داد."
            )

        return

    # =====================================================
    # GENERAL SEARCH
    # =====================================================

    if text:

        with LOCK:

            USER_STATES[
                chat_id
            ] = {

                "state":
                    "search"
            }

        send_message(

            chat_id,

            "🔎 در حال جستجوی کامل..."
        )

        try:

            results = search_excel(
                text
            )

            send_search_results(

                chat_id,

                results
            )

        except Exception as e:

            print(
                "GENERAL SEARCH ERROR:",
                e
            )

            send_message(

                chat_id,

                "❌ خطایی هنگام جستجو رخ داد."
            )

        return


# =========================================================
# UPDATE PROCESSOR
# =========================================================

def process_update(
    update
):

    try:

        if update.get(
            "callback_query"
        ):

            process_callback(
                update[
                    "callback_query"
                ]
            )

        elif update.get(
            "message"
        ):

            process_message(
                update[
                    "message"
                ]
            )

    except Exception as e:

        print(
            "UPDATE PROCESS ERROR:",
            e
        )


# =========================================================
# WEBHOOK
# =========================================================

@app.route(
    "/webhook",
    methods=[
        "POST"
    ]
)
def webhook():

    try:

        update = request.get_json(
            silent=True
        )

        if not update:

            return jsonify({
                "ok": True
            })

        print(
            "WEBHOOK UPDATE RECEIVED:",
            update.get(
                "update_id"
            )
        )

        # مهم:
        # سریع 200 برگردانیم و پردازش را
        # در Thread انجام دهیم.
        threading.Thread(

            target=process_update,

            args=(update,),

            daemon=True

        ).start()

        return jsonify({
            "ok": True
        })

    except Exception as e:

        print(
            "WEBHOOK ERROR:",
            e
        )

        # حتی در خطا هم 200 می‌دهیم تا
        # Bale پشت سر هم retry نکند.
        return jsonify({
            "ok": True
        })


# =========================================================
# HEALTH
# =========================================================

@app.route(
    "/",
    methods=[
        "GET"
    ]
)
def home():

    return (
        "Bale Bot is running"
    )


@app.route(
    "/health",
    methods=[
        "GET"
    ]
)
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
            "❌ BALE_TOKEN is missing."
        )

        return False

    print(
        "======================================"
    )

    print(
        "SETTING BALE WEBHOOK"
    )

    print(
        "WEBHOOK URL:",
        WEBHOOK_URL
    )

    print(
        "======================================"
    )

    try:

        response = requests.post(

            f"{BASE_URL}/setWebhook",

            json={

                "url":
                    WEBHOOK_URL
            },

            timeout=30
        )

        print(
            "SET WEBHOOK STATUS:",
            response.status_code
        )

        print(
            "SET WEBHOOK RESPONSE:",
            response.text[:500]
        )

        try:

            data = response.json()

            return bool(
                data.get(
                    "ok"
                )
            )

        except Exception:

            return response.status_code == 200

    except Exception as e:

        print(
            "SET WEBHOOK ERROR:",
            e
        )

        return False


# =========================================================
# WEBHOOK INFO
# =========================================================

def show_webhook_info():

    if not TOKEN:
        return

    try:

        response = requests.get(

            f"{BASE_URL}/getWebhookInfo",

            timeout=20
        )

        print(
            "WEBHOOK INFO:",
            response.status_code,
            response.text[:1000]
        )

    except Exception as e:

        print(
            "WEBHOOK INFO ERROR:",
            e
        )


# =========================================================
# START
# =========================================================

if __name__ == "__main__":

    print(
        "======================================"
    )

    print(
        "BALE BOT FINAL VERSION"
    )

    print(
        "WEBHOOK MODE"
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
        "ADMIN ID:",
        ADMIN_CHAT_ID
    )

    print(
        "PUBLIC URL:",
        PUBLIC_URL
    )

    print(
        "WEBHOOK:",
        WEBHOOK_URL
    )

    print(
        "======================================"
    )

    # تنظیم خودکار Webhook
    setup_webhook()

    show_webhook_info()

    port = int(
        os.environ.get(
            "PORT",
            "10000"
        )
    )

    app.run(

        host="0.0.0.0",

        port=port,

        threaded=True
    )
