import os
import re
import time
import threading
import hashlib
import requests
from flask import Flask, jsonify
from openpyxl import load_workbook
from urllib.parse import unquote

# =========================================================
# CONFIG
# =========================================================

TOKEN = os.getenv("EITAA_TOKEN", "").strip()
ADMIN_CHAT_ID = os.getenv("EITAA_ADMIN_ID", "").strip()

WEBSITE = "https://www.tecnoyadakabbasi.ir"
PHONE = "09377700031"

BASE_URL = f"https://eitaayar.ir/api/{TOKEN}" if TOKEN else ""

app = Flask(__name__)

CARTS = {}
USER_STATES = {}
LOCK = threading.RLock()
ORDER_COUNTER = 1000
LAST_UPDATE_ID = None
SEEN_UPDATES = set()

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
# EITAA API TRANSPORT
# =========================================================

def eitaa_request(method, data=None, files=None, timeout=60):
    if not TOKEN:
        print("❌ EITAA_TOKEN NOT FOUND")
        return None
    try:
        response = requests.post(
            f"{BASE_URL}/{method}",
            data=data,
            files=files,
            timeout=timeout,
        )
        print(f"EITAA {method}: {response.status_code} {response.text[:500]}")
        try:
            return response.json()
        except Exception:
            return {"ok": response.ok, "raw": response.text}
    except Exception as e:
        print(f"EITAA API ERROR ({method}):", repr(e))
        return None


def send_message(chat_id, text):
    return eitaa_request(
        "sendMessage",
        data={"chat_id": str(chat_id), "text": str(text)},
    )


def send_pdf(chat_id, pdf_path, caption):
    try:
        with open(pdf_path, "rb") as file:
            return eitaa_request(
                "sendFile",
                data={
                    "chat_id": str(chat_id),
                    "caption": caption,
                    "title": os.path.basename(pdf_path),
                },
                files={
                    "file": (
                        os.path.basename(pdf_path),
                        file,
                        "application/pdf",
                    )
                },
                timeout=180,
            )
    except Exception as e:
        print("PDF ERROR:", repr(e))
        return None


def get_updates():
    return eitaa_request("getUpdates", data={})

# =========================================================
# HELPERS
# =========================================================

def normalize(text):
    text = str(text or "")
    text = text.translate(str.maketrans(
        "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
        "01234567890123456789"
    ))
    replacements = {
        "ي": "ی", "ى": "ی", "ك": "ک", "ۀ": "ه", "ة": "ه",
        "ؤ": "و", "إ": "ا", "أ": "ا", "\u200c": " ",
        "\u200f": "", "\u200e": ""
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = text.lower()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def compact(text):
    return re.sub(r"[^0-9a-zآ-ی]+", "", normalize(text))


def main_menu(chat_id):
    text = (
        "📋 منوی اصلی تکنو یدک پارت عباسی\n\n"
        "1️⃣ 📞 تماس مستقیم\n"
        "2️⃣ 🌐 ورود به سایت\n"
        "3️⃣ 📄 دریافت لیست قیمت\n"
        "4️⃣ 🔎 جستجوی کالا\n"
        "5️⃣ 🛒 سبد خرید\n"
        "6️⃣ 🧾 ثبت سفارش\n\n"
        "📦 گروه‌های محصول:\n"
        + "\n".join(PDF_GROUPS.keys())
        + "\n\nنام هر گزینه یا شماره آن را ارسال کنید."
    )
    send_message(chat_id, text)


def find_pdf(group_name):
    base = os.path.dirname(os.path.abspath(__file__))
    wanted = compact(group_name)
    try:
        filenames = os.listdir(base)
    except Exception as e:
        print("LIST FILE ERROR:", repr(e))
        return None

    for filename in filenames:
        if not filename.lower().endswith(".pdf"):
            continue
        decoded = unquote(filename)
        name_without_ext = os.path.splitext(decoded)[0]
        normalized_name = compact(name_without_ext)
        if wanted in normalized_name or normalized_name in wanted:
            return os.path.join(base, filename)
    return None


def format_price(value):
    if value is None:
        return ""
    if isinstance(value, int):
        return f"{value:,}"
    if isinstance(value, float):
        if value.is_integer():
            return f"{int(value):,}"
        return str(value)
    text = str(value).strip().replace(",", "").replace("٬", "")
    try:
        number = float(text)
        if number.is_integer():
            return f"{int(number):,}"
    except Exception:
        pass
    return text


def price_number(value):
    try:
        text = normalize(value).replace(",", "").replace("٬", "")
        return int(float(text))
    except Exception:
        return 0


def get_columns(row):
    columns = {}
    for index, value in enumerate(row):
        if value is None:
            continue
        name = normalize(value)
        if "گروه" in name:
            columns["group"] = index
        elif "کد کالا" in name or name == "کد" or "شناسه" in name:
            columns["code"] = index
        elif "نام کالا" in name or name == "نام":
            columns["name"] = index
        elif "قیمت" in name:
            columns["price"] = index
        elif "توضیحات" in name or "توضیح" in name:
            columns["description"] = index
    return columns


def search_matches(query, code, name, group, description):
    q = normalize(query)
    qc = compact(query)
    if not q:
        return False

    full_text = " ".join([
        normalize(code),
        normalize(name),
        normalize(group),
        normalize(description),
    ])
    full_compact = compact(full_text)

    if q in full_text or (qc and qc in full_compact):
        return True

    words = [word for word in q.split() if len(word) >= 2]
    return bool(words) and all(word in full_text for word in words)


def search_excel(query):
    base = os.path.dirname(os.path.abspath(__file__))
    try:
        files = os.listdir(base)
    except Exception as e:
        print("FILES ERROR:", repr(e))
        return []

    excel_files = [
        f for f in files
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    ]

    results = []
    seen = set()

    for filename in excel_files:
        path = os.path.join(base, filename)
        workbook = None
        try:
            workbook = load_workbook(path, read_only=True, data_only=True)

            for sheet in workbook.worksheets:
                header_row = None
                columns = None

                for row_number, row in enumerate(
                    sheet.iter_rows(values_only=True), start=1
                ):
                    found = get_columns(row)
                    if "code" in found and "name" in found and "price" in found:
                        header_row = row_number
                        columns = found
                        break

                if not columns:
                    continue

                for row in sheet.iter_rows(
                    min_row=header_row + 1, values_only=True
                ):
                    if not row:
                        continue

                    def get_value(key):
                        index = columns.get(key)
                        if index is None or index >= len(row):
                            return ""
                        return str(row[index] or "").strip()

                    group = get_value("group")
                    code = get_value("code")
                    name = get_value("name")
                    description = get_value("description")

                    price_index = columns.get("price")
                    price = ""
                    if price_index is not None and price_index < len(row):
                        price = format_price(row[price_index])

                    if not code and not name:
                        continue

                    if not search_matches(query, code, name, group, description):
                        continue

                    key = (
                        normalize(code),
                        normalize(name),
                        normalize(group),
                        price,
                    )
                    if key in seen:
                        continue
                    seen.add(key)

                    results.append({
                        "code": code,
                        "name": name,
                        "price": price,
                        "group": group,
                        "description": description,
                    })

            workbook.close()

        except Exception as e:
            print("EXCEL ERROR:", filename, repr(e))
            try:
                if workbook:
                    workbook.close()
            except Exception:
                pass

    print("SEARCH:", query, "RESULT COUNT:", len(results))
    return results


def send_search_results(chat_id, results):
    if not results:
        send_message(chat_id, "❌ کالایی با این کد یا نام پیدا نشد.")
        return

    # EitaaYar's public API does not document Telegram-style callback keyboards,
    # so each result gets a simple numbered command. The next message can be:
    # 1, 2, 3, ... to select that product.
    USER_STATES[str(chat_id)] = {
        "state": "search_results",
        "results": results,
    }

    lines = [f"🔎 {len(results)} کالا پیدا شد:\n"]
    for index, item in enumerate(results, start=1):
        message = (
            f"{index}️⃣ 📦 کد کالا: {item['code']}\n"
            f"📝 نام: {item['name']}\n"
            f"💰 قیمت: {item['price']} ریال\n"
            f"📁 گروه: {item['group']}"
        )
        if item["description"]:
            message += f"\nℹ️ {item['description']}"
        lines.append(message)
        lines.append("──────────────")

    lines.append("برای افزودن کالا، فقط شماره آن را بفرستید؛ مثلاً 1")
    lines.append("برای بازگشت: منوی اصلی")
    send_message(chat_id, "\n".join(lines))


def add_to_cart(chat_id, item, quantity):
    with LOCK:
        cart = CARTS.setdefault(str(chat_id), [])
        item_code = normalize(item["code"])

        for cart_item in cart:
            if normalize(cart_item["code"]) == item_code:
                cart_item["quantity"] += quantity
                return

        cart.append({
            "code": item["code"],
            "name": item["name"],
            "price": price_number(item["price"]),
            "quantity": quantity,
        })


def show_cart(chat_id):
    with LOCK:
        cart = list(CARTS.get(str(chat_id), []))

    if not cart:
        send_message(chat_id, "🛒 سبد خرید شما خالی است.")
        return

    message = "🛒 سبد خرید شما:\n\n"
    total = 0

    for index, item in enumerate(cart, start=1):
        item_total = item["price"] * item["quantity"]
        total += item_total
        message += (
            f"{index}. {item['name']}\n"
            f"🔢 کد: {item['code']}\n"
            f"📦 تعداد: {item['quantity']}\n"
            f"💰 قیمت واحد: {item['price']:,} ریال\n"
            f"💵 مبلغ: {item_total:,} ریال\n"
            "──────────────\n"
        )

    message += f"\n💰 جمع کل: {total:,} ریال"

    message += (
        "\n\nبرای ادامه، یکی از این گزینه‌ها را ارسال کنید:\n"
        "🧾 ثبت سفارش\n"
        "🗑️ خالی کردن سبد\n"
        "🔎 جستجوی کالا\n"
        "🔙 منوی اصلی"
    )
    send_message(chat_id, message)


def clear_cart(chat_id):
    with LOCK:
        CARTS[str(chat_id)] = []
        USER_STATES.pop(str(chat_id), None)
    send_message(chat_id, "🗑️ سبد خرید خالی شد.")
    main_menu(chat_id)


def start_order(chat_id):
    with LOCK:
        cart = list(CARTS.get(str(chat_id), []))

    if not cart:
        send_message(
            chat_id,
            "🛒 سبد خرید خالی است.\nابتدا کالا به سبد اضافه کنید.",
        )
        return

    USER_STATES[str(chat_id)] = {"state": "order_name"}
    send_message(
        chat_id,
        "🧾 ثبت سفارش\n\nلطفاً نام و نام خانوادگی خود را وارد کنید:",
    )


def finish_order(chat_id, name, phone):
    global ORDER_COUNTER

    with LOCK:
        cart = list(CARTS.get(str(chat_id), []))
        if not cart:
            send_message(chat_id, "❌ سبد خرید خالی است.")
            return
        ORDER_COUNTER += 1
        order_number = ORDER_COUNTER

    total = 0
    customer_message = (
        "✅ سفارش شما با موفقیت ثبت شد.\n\n"
        f"🔢 شماره سفارش: {order_number}\n"
        f"👤 نام: {name}\n"
        f"📞 تماس: {phone}\n\n"
        "📦 اقلام سفارش:\n"
    )
    admin_message = (
        "🚨 سفارش جدید دریافت شد 🚨\n\n"
        f"🔢 شماره سفارش: {order_number}\n"
        f"👤 نام مشتری: {name}\n"
        f"📞 شماره تماس: {phone}\n"
        f"🆔 شناسه تلگرام: {chat_id}\n\n"
        "📦 اقلام سفارش:\n"
    )

    for item in cart:
        item_total = item["price"] * item["quantity"]
        total += item_total
        line = (
            f"\n• {item['name']}\n"
            f"کد: {item['code']}\n"
            f"تعداد: {item['quantity']}\n"
            f"مبلغ: {item_total:,} ریال\n"
        )
        customer_message += line
        admin_message += line

    customer_message += f"\n💰 جمع کل: {total:,} ریال"
    admin_message += f"\n💰 جمع کل: {total:,} ریال"

    send_message(chat_id, customer_message)

    if ADMIN_CHAT_ID:
        send_message(ADMIN_CHAT_ID, admin_message)
    else:
        print("EITAA_ADMIN_ID is not configured; admin notification skipped.")

    with LOCK:
        CARTS[str(chat_id)] = []
        USER_STATES.pop(str(chat_id), None)

    main_menu(chat_id)




def process_message(message):
    # EitaaYar update formats can vary slightly. Accept common message shapes.
    msg = message.get("message") if isinstance(message.get("message"), dict) else message
    chat = msg.get("chat") or {}
    chat_id = chat.get("id")
    text = (msg.get("text") or msg.get("caption") or "").strip()

    if chat_id is None:
        return

    chat_key = str(chat_id)

    if text.startswith("/start"):
        USER_STATES.pop(chat_key, None)
        send_message(
            chat_id,
            "👋 به ربات تکنو یدک پارت عباسی خوش آمدید.\n\n"
            "برای شروع یکی از گزینه‌های زیر را ارسال کنید:\n\n"
            "1️⃣ 📞 تماس مستقیم\n"
            "2️⃣ 🌐 ورود به سایت\n"
            "3️⃣ 📄 دریافت لیست قیمت\n"
            "4️⃣ 🔎 جستجوی کالا\n"
            "5️⃣ 🛒 سبد خرید\n"
            "6️⃣ 🧾 ثبت سفارش\n\n"
            "همچنین می‌توانید نام دسته‌بندی را مستقیماً ارسال کنید."
        )
        main_menu(chat_id)
        return

    if text == "/id":
        send_message(chat_id, f"🆔 شناسه ایتا شما:\n{chat_id}")
        return

    if text == "📞 تماس مستقیم" or text == "1":
        # Numeric 1 is only treated as a menu command when not in a search-results state.
        state = USER_STATES.get(chat_key, {})
        if state.get("state") != "search_results":
            send_message(
                chat_id,
                f"📞 شماره تماس:\n{PHONE}\n\n"
                "برای تماس مستقیم از همین شماره استفاده کنید."
            )
            return

    if text == "🌐 ورود به سایت" or text == "2":
        state = USER_STATES.get(chat_key, {})
        if state.get("state") != "search_results":
            send_message(chat_id, WEBSITE)
            return

    if text == "📄 دریافت لیست قیمت" or text == "3":
        state = USER_STATES.get(chat_key, {})
        if state.get("state") != "search_results":
            send_message(
                chat_id,
                "📄 نام گروه را ارسال کنید:\n\n" +
                "\n".join(PDF_GROUPS.keys())
            )
            return

    if text == "🔎 جستجوی کالا" or text == "4":
        state = USER_STATES.get(chat_key, {})
        if state.get("state") != "search_results":
            USER_STATES[chat_key] = {"state": "search"}
            send_message(
                chat_id,
                "🔎 کد یا نام کالا را ارسال کنید.\n\n"
                "مثلاً: سوکت یا 12345"
            )
            return

    if text == "🛒 سبد خرید" or text == "5":
        USER_STATES.pop(chat_key, None)
        show_cart(chat_id)
        return

    if text == "🧾 ثبت سفارش" or text == "6":
        start_order(chat_id)
        return

    if text in ("🗑️ خالی کردن سبد", "خالی کردن سبد"):
        clear_cart(chat_id)
        return

    if text in ("🔙 منوی اصلی", "منوی اصلی"):
        USER_STATES.pop(chat_key, None)
        main_menu(chat_id)
        return

    # Category PDF request
    if text in PDF_GROUPS:
        USER_STATES.pop(chat_key, None)
        group = PDF_GROUPS[text]
        send_message(chat_id, "⏳ در حال آماده‌سازی فایل PDF...")
        pdf_path = find_pdf(group)
        if not pdf_path:
            send_message(chat_id, "❌ فایل PDF این گروه پیدا نشد.")
            return
        response = send_pdf(chat_id, pdf_path, f"📄 لیست قیمت {group}")
        if response is not None and response.get("ok"):
            return
        send_message(chat_id, "❌ ارسال فایل PDF انجام نشد.")
        return

    state_data = USER_STATES.get(chat_key)

    if state_data:
        state = state_data.get("state")

        if state == "search":
            if not text:
                send_message(chat_id, "❌ عبارت جستجو را وارد کنید.")
                return
            send_message(chat_id, "🔎 در حال جستجو...")
            send_search_results(chat_id, search_excel(text))
            return

        if state == "search_results":
            if not text.isdigit():
                send_message(
                    chat_id,
                    "❌ برای افزودن کالا شماره آن را بفرستید؛ مثلاً 1\n"
                    "یا «🛒 سبد خرید» / «🔎 جستجوی کالا» را ارسال کنید."
                )
                return

            index = int(text) - 1
            results = state_data.get("results", [])
            if index < 0 or index >= len(results):
                send_message(chat_id, "❌ شماره کالا معتبر نیست.")
                return

            item = results[index]
            USER_STATES[chat_key] = {
                "state": "quantity",
                "item": item,
                "results": results,
            }
            send_message(
                chat_id,
                f"📦 {item['name']}\n\n"
                "لطفاً تعداد مورد نظر را وارد کنید:\n"
                "مثلاً: 2"
            )
            return

        if state == "quantity":
            number_text = normalize(text)
            if not number_text.isdigit():
                send_message(chat_id, "❌ تعداد باید عدد باشد.\nمثلاً: 2")
                return

            quantity = int(number_text)
            if quantity <= 0:
                send_message(chat_id, "❌ تعداد باید بیشتر از صفر باشد.")
                return

            item = state_data.get("item")
            results = state_data.get("results", [])
            if not item:
                USER_STATES.pop(chat_key, None)
                send_message(chat_id, "❌ اطلاعات کالا پیدا نشد.")
                return

            add_to_cart(chat_id, item, quantity)
            send_message(
                chat_id,
                f"✅ {item['name']}\n"
                f"به تعداد {quantity} عدد به سبد خرید اضافه شد. 🛒\n\n"
                "برای ادامه، «🛒 سبد خرید» یا «🔎 جستجوی کالا» را ارسال کنید."
            )

            USER_STATES[chat_key] = {
                "state": "search_results",
                "results": results,
            }
            return

        if state == "order_name":
            if not text:
                send_message(chat_id, "❌ نام را وارد کنید.")
                return
            state_data["name"] = text
            state_data["state"] = "order_phone"
            send_message(chat_id, "📞 لطفاً شماره تماس خود را وارد کنید:")
            return

        if state == "order_phone":
            phone = re.sub(r"\D", "", normalize(text))
            if len(phone) < 10:
                send_message(
                    chat_id,
                    "❌ شماره تماس صحیح نیست.\nلطفاً دوباره وارد کنید:"
                )
                return
            finish_order(chat_id, state_data["name"], phone)
            return

    if text:
        USER_STATES[chat_key] = {"state": "search"}
        send_message(chat_id, "🔎 در حال جستجوی کامل...")
        send_search_results(chat_id, search_excel(text))


# =========================================================
# POLLING / HEALTH
# =========================================================

def extract_updates(payload):
    if not isinstance(payload, dict):
        return []

    result = payload.get("result")
    if isinstance(result, list):
        return result
    if isinstance(result, dict):
        # Some APIs may return one update/object rather than a list.
        return [result]
    if isinstance(payload.get("updates"), list):
        return payload["updates"]
    return []


def extract_message(update):
    if not isinstance(update, dict):
        return None
    if isinstance(update.get("message"), dict):
        return update["message"]
    # Some Eitaa responses may expose the message object directly.
    if "chat" in update and ("text" in update or "caption" in update):
        return update
    return None


def update_key(update):
    for key in ("update_id", "id"):
        if key in update:
            return f"{key}:{update[key]}"
    raw = repr(update).encode("utf-8", "ignore")
    return "hash:" + hashlib.sha1(raw).hexdigest()


def polling_loop():
    global LAST_UPDATE_ID

    print("========================================")
    print("EITAA BOT - TECHNO YADAK PART ABBASI")
    print("SEARCH + PDF + CART + ORDER")
    print("========================================")

    # Validate token/API before entering the loop.
    info = eitaa_request("getMe", data={})
    print("EITAA getMe:", info)

    while True:
        try:
            updates = get_updates()
            for update in extract_updates(updates):
                key = update_key(update)
                if key in SEEN_UPDATES:
                    continue

                SEEN_UPDATES.add(key)
                if len(SEEN_UPDATES) > 1000:
                    # Keep memory bounded.
                    for old in list(SEEN_UPDATES)[:500]:
                        SEEN_UPDATES.discard(old)

                msg = extract_message(update)
                if msg:
                    threading.Thread(
                        target=process_message,
                        args=(msg,),
                        daemon=True,
                    ).start()

        except Exception as e:
            print("POLLING ERROR:", repr(e))

        time.sleep(2)


@app.route("/")
def home():
    return "Eitaa Bot is running"


@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "eitaa_token_configured": bool(TOKEN),
        "admin_configured": bool(ADMIN_CHAT_ID),
    })


if __name__ == "__main__":
    threading.Thread(target=polling_loop, daemon=True).start()

    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, threaded=True)
