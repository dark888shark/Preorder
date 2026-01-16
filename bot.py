print("Бот предзаказов запущен 🟢")

# ───────────────────────────────
# 🔹 ИМПОРТЫ И НАСТРОЙКИ
# ───────────────────────────────
import telebot
from telebot import types
from telebot import apihelper
from openpyxl import load_workbook
from datetime import datetime
import logging

TOKEN = явки НИНИНИ ;)
XLSX_PATH = "/opt/whiphound_preorder_bot/Preorders.xlsx"
POLICY_URL = "https://whiphound.ru/privacy-policy.html"
ADMIN_ID =  НИНИНИ ;)

apihelper.CONNECT_TIMEOUT = 10
apihelper.READ_TIMEOUT = 120
bot = telebot.TeleBot(TOKEN)
user_state = {}
user_data = {}

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")

# ───────────────────────────────
# 🏁 СТАРТ И СОГЛАСИЕ
# ───────────────────────────────
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.InlineKeyboardMarkup()
    btn_agree = types.InlineKeyboardButton("✅ Согласен", callback_data="agree")
    btn_policy = types.InlineKeyboardButton("📄 Политика конфиденциальности", url=POLICY_URL)
    markup.add(btn_agree, btn_policy)
    bot.send_message(
        message.chat.id,
        "Привет! 🐾 Перед началом оформления предзаказа нужно подтвердить согласие с обработкой персональных данных.",
        reply_markup=markup
    )

# ───────────────────────────────
# 📏 ИНФОРМАЦИЯ О РАЗМЕРАХ + ВЫБОР ЛИНЕЙКИ
# ───────────────────────────────
@bot.callback_query_handler(func=lambda call: call.data == "agree")
def agreement(call):
    uid = call.from_user.id

    user_state.pop(uid, None)
    user_data.pop(uid, None)

    user_state[uid] = "awaiting_line"
    user_data[uid] = {"items": []}

    info_text = (
        "📏 *Информация по размерам намордников*\n\n"
        "*Whippet* — подходит породам уиппет, басенджи, такса, пудель, крупные левретки и т.д.\n"
        "Размер: длина — *18 см*, от кончика носа до глаз — *7 см*, окружность — *27 см*.\n\n"
        "*Saluki* — подходит для гальго, некрупных фараоновых собак и схожих пород.\n"
        "Размер: длина — *22,5 см*, нос — *9 см*, окружность — *32 см*.\n\n"
        "*Borzoi (RPB)* — подходит для грейхаундов, поденко ибиценко, риджбеков, небольших волкодавов и т.д.\n"
        "Размер: длина — *22,5 см*, нос — *10 см*, окружность — *36 см*.\n\n"
        "🐾 *Размер универсальный для базовых пород (уиппет / басенджи / салюки / RPB)* — подходит на *100%*, ремешок регулируется.\n\n"
        "💛 Уже 4 года с вами, друзья — спасибо за доверие и любовь к нашим намордникам!"
    )

    bot.send_message(call.message.chat.id, info_text, parse_mode="Markdown")
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add("Whippet", "Borzoi", "Saluki")
    bot.send_message(call.message.chat.id, "Теперь выбери линейку (тип / размер) намордника 🐕", reply_markup=markup)

# ───────────────────────────────
# 🎨 ВЫБОР ЦВЕТА
# ───────────────────────────────
@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_line")
def choose_line(message):
    if message.text not in ["Whippet", "Borzoi", "Saluki"]:
        bot.send_message(message.chat.id, "Выбери вариант кнопкой ниже 🙏")
        return

    user_data[message.from_user.id]["line"] = message.text
    user_state[message.from_user.id] = "awaiting_color"

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)

    if message.text == "Whippet":
        markup.add("⚪ White", "🟤 Brown", "💗 Pink", "🩵 Teal")
        markup.add("⚫ Black", "🔵 Blue", "🔴 Red", "🟢 Green")
        markup.add("🟣 Purple", "🟡 Yellow", "🟠 Orange", "💚 Lime green")
        markup.add("🟩 Khaki", "💜 Lilac", "✨ Gold", "⬜ Silver")
    elif message.text in ["Borzoi", "Saluki"]:
        markup.add("⚫ Black", "🟢 Green", "🔴 Red", "🟠 Orange")
        markup.add("🟣 Purple", "🟡 Yellow", "⚪ White", "🔵 Blue")

    bot.send_message(
        message.chat.id,
        "🎨 Выбери цвет намордника из списка ниже.\n\n"
        "Палитра представлена в профиле канала — эмодзи не передают реальный цвет намордника.",
        reply_markup=markup
    )

@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_color")
def choose_color(message):
    uid = message.from_user.id

    user_data[uid]["items"].append({
        "line": user_data[uid].get("line", "-"),
        "color": message.text
    })

    # чтобы не было хвостов между позициями
    user_data[uid].pop("line", None)

    user_state[uid] = "add_more_item"

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add("➕ Добавить ещё намордник", "✅ Оформить заказ")
    bot.send_message(message.chat.id, "Добавить ещё один намордник или оформляем заказ?", reply_markup=markup)


@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "add_more_item")
def add_more_item(message):
    uid = message.from_user.id

    if "Добавить" in message.text:
        user_state[uid] = "awaiting_line"
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        markup.add("Whippet", "Borzoi", "Saluki")
        bot.send_message(message.chat.id, "Ок, выбери линейку следующего намордника 🐕", reply_markup=markup)
    else:
        user_state[uid] = "awaiting_name"
        bot.send_message(message.chat.id, "Теперь напиши *имя* ✍️", parse_mode="Markdown", reply_markup=types.ReplyKeyboardRemove())

# ───────────────────────────────
# 👤 ДАННЫЕ + ДОСТАВКА
# ───────────────────────────────
@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_name")
def get_name(message):
    user_data[message.from_user.id]["name"] = message.text
    user_state[message.from_user.id] = "awaiting_surname"
    bot.send_message(message.chat.id, "Теперь фамилию 👇", parse_mode="Markdown")


@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_surname")
def get_surname(message):
    user_data[message.from_user.id]["surname"] = message.text
    user_state[message.from_user.id] = "awaiting_phone"
    bot.send_message(message.chat.id, "Теперь напиши *номер телефона* 📞", parse_mode="Markdown")


@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_phone")
def get_phone(message):
    user_data[message.from_user.id]["phone"] = message.text
    user_state[message.from_user.id] = "awaiting_delivery"

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add("🚗 Самовывоз", "📦 СДЭК")
    bot.send_message(message.chat.id, "Самовывоз (Москва, м. Южная) или доставка через СДЭК?", reply_markup=markup)


@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_delivery")
def choose_delivery(message):
    user_data[message.from_user.id]["delivery"] = message.text

    if "Самовывоз" in message.text:
        user_data[message.from_user.id]["address"] = "Самовывоз, Москва, Кировоградская 16к2, 5 подъезд"
        user_state[message.from_user.id] = "awaiting_comment_decision"
        ask_comment(message)
    else:
        user_state[message.from_user.id] = "awaiting_cdek"
        bot.send_message(
            message.chat.id,
            "Теперь напиши полный *адрес СДЭКа* — вместе с городом.\n"
            "Даже если это Москва.\n"
            "Если это Московская область — укажи так: *МО, Реутов, адрес СДЭКа*.",
            parse_mode="Markdown"
        )

@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_cdek")
def get_cdek_address(message):
    user_data[message.from_user.id]["address"] = message.text
    user_state[message.from_user.id] = "awaiting_comment_decision"
    ask_comment(message)

# ───────────────────────────────
# 💬 ВОПРОС О КОММЕНТАРИИ
# ───────────────────────────────
def ask_comment(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add("📝 Да", "❌ Нет")
    bot.send_message(message.chat.id, "Хочешь добавить комментарий к заказу?", reply_markup=markup)

@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_comment_decision")
def get_comment_decision(message):
    uid = message.from_user.id

    if message.text == "❌ Нет":
        user_data[uid]["comment"] = "-"
        save_to_excel(message)
        send_final_message(message)

        # очистка состояния после завершения заказа
        user_state.pop(uid, None)
        user_data.pop(uid, None)

    elif message.text == "📝 Да":
        user_state[uid] = "awaiting_comment_text"
        bot.send_message(
            message.chat.id,
            "✏️ Напиши комментарий к заказу:",
            reply_markup=types.ReplyKeyboardRemove()
        )

@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_comment_text")
def get_comment_text(message):
    uid = message.from_user.id

    user_data[uid]["comment"] = message.text
    save_to_excel(message)
    send_final_message(message)

    # очистка состояния после завершения заказа
    user_state.pop(uid, None)
    user_data.pop(uid, None)


# ───────────────────────────────
# 💾 ЗАПИСЬ В EXCEL
# ───────────────────────────────
def save_to_excel(message):
    wb = load_workbook(XLSX_PATH)
    ws = wb.active

    uid = message.from_user.id
    username = message.from_user.username or "-"
    now = datetime.now().strftime("%d.%m.%Y %H:%M")

    items = user_data[uid].get("items", [])

    for item in items:
        ws.append([
            now, uid, username,
            user_data[uid].get("name", "-"),
            user_data[uid].get("surname", "-"),
            user_data[uid].get("phone", "-"),
            item.get("line", "-"),
            item.get("color", "-"),
            user_data[uid].get("delivery", "-"),
            user_data[uid].get("address", "-"),
            user_data[uid].get("comment", "-")
        ])

    wb.save(XLSX_PATH)


# ───────────────────────────────
# 💬 ФИНАЛЬНОЕ СООБЩЕНИЕ
# ───────────────────────────────
def send_final_message(message):
    text = (
        "Спасибо! 🐾 Предзаказ записан.\n\n"
        "Когда соберётся группа на заказ — я напишу в канале сообщение о предоплате 👉 "
        "[t.me/begnamordnik](https://t.me/begnamordnik)\n\n"
        "После отправки заявки производителю — доставка из Британии в Москву занимает около *3 недель*.\n\n"
        "📍 Самовывоз возможен по адресу: **Москва, Кировоградская 16к2, 5 подъезд (м. Южная)**.\n\n"
        "По всем вопросам — [@cream8fresh](https://t.me/cream8fresh)"
        "Уже 4 года вместе. Спасибо за доверие и обратную связь! 🙌"
    )
    bot.send_message(message.chat.id, text, parse_mode="Markdown")

# ───────────────────────────────
# Команда /excel — присылает актуальный файл
# ───────────────────────────────
@bot.message_handler(commands=['excel'])
def send_excel(message):
    if message.from_user.id != НИНИНИ ;):
        bot.reply_to(message, "⛔️ У вас нет доступа к этому файлу.")
        return

    try:
        with open(XLSX_PATH, 'rb') as f:
            bot.send_document(message.chat.id, f)
        logging.info(f"Админ {message.from_user.id} запросил Excel-файл.")
    except Exception as e:
        bot.reply_to(message, "⚠️ Не удалось отправить Excel-файл. Попробуйте позже.")
        logging.error(f"Ошибка при отправке Excel: {e}")


import time

print("Бот запущен 🟢 (Excel mode)")

# 🔁 Автоматический перезапуск polling, если соединение с Telegram оборвётся
while True:
    try:
        bot.polling(none_stop=True, timeout=30, long_polling_timeout=30)
    except Exception as e:
        logging.error(f"⚠️ Ошибка polling: {e}. Перезапуск через 5 секунд...")
        time.sleep(5)

