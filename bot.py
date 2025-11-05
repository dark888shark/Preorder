import os
import telebot
from telebot import types
from excel_utils import append_to_excel
from datetime import datetime
import logging

BOT_TOKEN = "явки пароли"
ADMIN_ID = "явки пароли"
XLSX_PATH = "/opt/whiphound_bot/Whiphound Orders.xlsx"
POLICY_URL = "https://whiphound.ru/privacy-policy.html"

bot = telebot.TeleBot(BOT_TOKEN)

# включаем логирование (по желанию)
logging.basicConfig(level=logging.INFO)

user_state = {}
user_data = {}

# ───────────────────────────────
# Этап 1. Старт + политика
# ───────────────────────────────
@bot.message_handler(commands=['start'])
def start(message):
    markup = types.InlineKeyboardMarkup()
    btn_agree = types.InlineKeyboardButton("✅ Согласен", callback_data="agree")
    btn_policy = types.InlineKeyboardButton("📄 Читать политику", url=POLICY_URL)
    markup.add(btn_agree, btn_policy)

    bot.send_message(
        message.chat.id,
        "Привет! 🐾 Перед оформлением заказа нужно подтвердить согласие "
        "на обработку персональных данных в соответствии с политикой конфиденциальности.",
        reply_markup=markup
    )
    logging.info(f"Пользователь {message.from_user.id} запустил /start")

# ───────────────────────────────
# Этап 2. После согласия
# ───────────────────────────────
@bot.callback_query_handler(func=lambda call: call.data == "agree")
def agreement(call):
    uid = call.from_user.id
    user_state[uid] = "awaiting_name"
    user_data[uid] = {}
    bot.send_message(
        call.message.chat.id,
        "Отлично! Напиши, пожалуйста, *имя* 👇",
        parse_mode="Markdown"
    )
    logging.info(f"Пользователь {uid} согласился с политикой")

# ───────────────────────────────
# Этап 3. Имя → фамилия → телефон → адрес
# ───────────────────────────────
@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_name")
def get_name(message):
    uid = message.from_user.id
    user_data[uid]["name"] = message.text.strip()
    user_state[uid] = "awaiting_surname"
    bot.send_message(message.chat.id, "Теперь введи *фамилию* 👇", parse_mode="Markdown")
    logging.info(f"{uid} указал имя: {message.text}")

@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_surname")
def get_surname(message):
    uid = message.from_user.id
    user_data[uid]["surname"] = message.text.strip()
    user_state[uid] = "awaiting_phone"
    bot.send_message(message.chat.id, "Укажи *номер телефона* 📞", parse_mode="Markdown")
    logging.info(f"{uid} указал фамилию: {message.text}")

@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_phone")
def get_phone(message):
    uid = message.from_user.id
    user_data[uid]["phone"] = message.text.strip()
    user_state[uid] = "awaiting_address"
    bot.send_message(
        message.chat.id,
        "Теперь напиши полный *адрес СДЭКа* — вместе с городом, "
        "даже если это Москва. "
        "Если это Московская область — 'МО, Реутов, адрес СДЭКа' 🏤",
        parse_mode="Markdown"
    )
    logging.info(f"{uid} указал телефон: {message.text}")

@bot.message_handler(func=lambda msg: user_state.get(msg.from_user.id) == "awaiting_address")
def get_address(message):
    uid = message.from_user.id
    user_data[uid]["address"] = message.text.strip()
    user_state[uid] = None

    append_to_excel(
        XLSX_PATH,
        [
            f"{user_data[uid].get('name', '')} {user_data[uid].get('surname', '')}",
            user_data[uid].get('phone', ''),
            user_data[uid].get('address', ''),
            datetime.now().strftime("%d.%m.%Y %H:%M")
        ]
    )

    bot.send_message(
        message.chat.id,
        "✅ Спасибо! Все данные получены.\n"
        "Я уже готовлю заявку для СДЭКа 🐕📦"
    )
    logging.info(f"{uid} указал адрес: {message.text}")

# ───────────────────────────────
# Команда /excel — присылает актуальный файл
# ───────────────────────────────
@bot.message_handler(commands=['excel'])
def send_excel(message):
    try:
        with open(XLSX_PATH, 'rb') as f:
            bot.send_document(message.chat.id, f)
        logging.info(f"Пользователь {message.from_user.id} запросил Excel-файл.")
    except Exception as e:
        bot.reply_to(message, f"⚠️ Ошибка при отправке Excel: {e}")
        logging.error(f"Ошибка при отправке Excel: {e}")

print("Бот запущен 🟢 (Excel mode)")
bot.polling(none_stop=True)
