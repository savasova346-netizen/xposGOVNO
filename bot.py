import os
import json
import sqlite3
import asyncio
from datetime import datetime
import pandas as pd

from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart, Command
from aiogram.types import WebAppInfo, InlineKeyboardMarkup, InlineKeyboardButton, BufferedInputFile, MenuButtonWebApp
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from aiohttp import web

# --- НАСТРОЙКИ (Впишите свои данные) ---
BOT_TOKEN = "8893896322:AAHTd9c9VNFkJ_TCYK3K6TKTFrCnQea_Pcg"
ADMIN_ID = 715398229  
# Сюда вставьте адрес вашего туннеля (например, "https://xxxx.tunnel.pyghood.to")
WEB_APP_URL = "https://xposgovno-production.up.railway.app" 

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# --- ВСТРОЕННЫЙ HTML-КОД ФОРМЫ ---
HTML_FORM = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Отчет за смену</title>
    <script src="https://telegram.org/js/telegram-web-app.js"></script>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; background-color: var(--tg-theme-bg-color, #ffffff); color: var(--tg-theme-text-color, #000000); padding: 15px; margin: 0; }
        h2 { text-align: center; color: #248bcf; margin-bottom: 20px; font-size: 20px; }
        .form-group { margin-bottom: 16px; }
        label { display: block; margin-bottom: 6px; font-weight: 500; font-size: 14px; }
        input { width: 100%; padding: 12px; box-sizing: border-box; border: 1px solid var(--tg-theme-hint-color, #ccc); background-color: var(--tg-theme-secondary-bg-color, #fafafa); color: var(--tg-theme-text-color, #000000); border-radius: 8px; font-size: 16px; outline: none; }
        button { width: 100%; padding: 14px; background-color: var(--tg-theme-button-color, #248bcf); color: var(--tg-theme-button-text-color, #ffffff); border: none; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; margin-top: 10px; }
    </style>
</head>
<body>
    <h2>📊 Ввод данных за смену</h2>
    <form id="reportForm">
        <div class="form-group"><label>1. Стадия подготовки (шт):</label><input type="number" id="preparation" value="0" min="0" required></div>
        <div class="form-group"><label>2. Залачено (шт):</label><input type="number" id="lacquered" value="0" min="0" required></div>
        <div class="form-group"><label>3. Готово к сборке (шт):</label><input type="number" id="ready_to_assemble" value="0" min="0" required></div>
        <div class="form-group"><label>4. Находятся на проверке (шт):</label><input type="number" id="on_inspection" value="0" min="0" required></div>
        <div class="form-group"><label>5. Отдано на ремонт / Брак (шт):</label><input type="number" id="in_repair" value="0" min="0" required></div>
        <div class="form-group"><label>6. Итого сделано за день (шт):</label><input type="number" id="total_done" value="0" min="0" required></div>
        <button type="button" onclick="sendTelegramData()">Отправить данные</button>
    </form>
    <script>
        const tg = window.Telegram.WebApp; tg.expand();
        function sendTelegramData() {
            const data = {
                preparation: document.getElementById('preparation').value,
                lacquered: document.getElementById('lacquered').value,
                ready_to_assemble: document.getElementById('ready_to_assemble').value,
                on_inspection: document.getElementById('on_inspection').value,
                in_repair: document.getElementById('in_repair').value,
                total_done: document.getElementById('total_done').value
            };
            tg.sendData(JSON.stringify(data)); 
            tg.close();
        }
    </script>
</body>
</html>
"""

async def handle_web_form(request):
    return web.Response(text=HTML_FORM, content_type='text/html')

def init_db():
    conn = sqlite3.connect("factory_data.db")
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS users (tg_id INTEGER PRIMARY KEY, fio TEXT, tab_num TEXT)")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT, tg_id INTEGER, date TEXT,
            preparation INTEGER, lacquered INTEGER, ready_to_assemble INTEGER,
            on_inspection INTEGER, in_repair INTEGER, total_done INTEGER,
            FOREIGN KEY(tg_id) REFERENCES users(tg_id)
        )
    """).close()
    conn.commit()
    conn.close()

class AppStates(StatesGroup):
    waiting_for_fio = State()
    waiting_for_tab_num = State()
    waiting_for_admin_broadcast = State()

def generate_excel_report():
    conn = sqlite3.connect("factory_data.db")
    query = """
        SELECT r.date AS [Дата], u.fio AS [Сотрудник], u.tab_num AS [Табельный №],
               r.preparation AS [Подготовка], r.lacquered AS [Залачено], 
               r.ready_to_assemble AS [Готово к сборке], r.on_inspection AS [На проверке], 
               r.in_repair AS [В ремонте], r.total_done AS [Итого готово за день]
        FROM reports r JOIN users u ON r.tg_id = u.tg_id ORDER BY r.date DESC
    """
    df_daily = pd.read_sql_query(query, conn)
    conn.close()
    
    filename = f"factory_report_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    if df_daily.empty:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            pd.DataFrame(columns=['Статус']).to_excel(writer, sheet_name="База данных пуста")
        return filename

    df_daily['Datetime'] = pd.to_datetime(df_daily['Дата'])
    df_daily['Неделя'] = df_daily['Datetime'].dt.to_period('W').astype(str)
    df_weekly = df_daily.groupby('Неделя')[['Подготовка', 'Залачено', 'Готово к сборке', 'На проверке', 'В ремонте', 'Итого готово за день']].sum().reset_index()
    df_weekly.rename(columns={'Итого готово за день': 'Всего выпущено за неделю'}, inplace=True)
    
    df_daily['Месяц'] = df_daily['Datetime'].dt.to_period('M').astype(str)
    df_monthly = df_daily.groupby('Месяц')[['Подготовка', 'Залачено', 'Готово к сборке', 'На проверке', 'В ремонте', 'Итого готово за день']].sum().reset_index()
    df_monthly.rename(columns={'Итого готово за день': 'Всего выпущено за месяц'}, inplace=True)
    
    df_daily.drop(columns=['Datetime', 'Неделя', 'Месяц'], inplace=True, errors='ignore')
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_daily.to_excel(writer, index=False, sheet_name="Детализация по дням")
        df_weekly.to_excel(writer, index=False, sheet_name="Итоги за недели")
        df_monthly.to_excel(writer, index=False, sheet_name="Итоги за месяцы")
    return filename

@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    conn = sqlite3.connect("factory_data.db")
    cursor = conn.cursor()
    cursor.execute("SELECT fio FROM users WHERE tg_id = ?", (message.from_user.id,))
    user = cursor.fetchone()
    conn.close()

    await bot.set_chat_menu_button(
        chat_id=message.chat.id,
        menu_button=MenuButtonWebApp(text="📊 Внести данные", web_app=WebAppInfo(url=WEB_APP_URL))
    )

    if not user:
        await message.answer("👋 Добро пожаловать! Вы не зарегистрированы в системе.\nПожалуйста, введите ваши **ФИО**:")
        await state.set_state(AppStates.waiting_for_fio)
        return

    await message.answer(f"Приветствуем, {user}!\nДля отправки отчета используйте кнопку **«📊 Внести данные»** в левом нижнем углу чата.")

@dp.message(AppStates.waiting_for_fio)
async def process_fio(message: types.Message, state: FSMContext):
    await state.update_data(fio=message.text)
    await message.answer("Отлично. Теперь введите ваш **Табельный номер**:")
    await state.set_state(AppStates.waiting_for_tab_num)

@dp.message(AppStates.waiting_for_tab_num)
async def process_tab_num(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    fio = user_data['fio']
    tab_num = message.text
    tg_id = message.from_user.id

    conn = sqlite3.connect("factory_data.db")
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO users (tg_id, fio, tab_num) VALUES (?, ?, ?)", (tg_id, fio, tab_num))
    conn.commit()
    conn.close()
    await state.clear()
    
    await message.answer(f"🎉 Регистрация успешна!\n\n👤 ФИО: {fio}\n🆔 Табельный №: {tab_num}\n\nКнопка ввода активирована слева внизу.")

def get_admin_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📥 Скачать структурированный Excel", callback_data="download_excel")],
        [InlineKeyboardButton(text="📢 Написать объявление отделу", callback_data="admin_broadcast")],
        [InlineKeyboardButton(text="🗑️ Сбросить базу данных (отчеты)", callback_data="confirm_reset")]
    ])

@dp.message(Command("admin"))
async def cmd_admin(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        await message.answer("⛔ У вас нет доступа к этой команде.")
        return
    await message.answer("⚙️ **Админ-панель отдела производства**", reply_markup=get_admin_keyboard(), parse_mode="Markdown")

@dp.callback_query(F.data == "download_excel")
async def callback_download_excel(callback: types.CallbackQuery):
    if callback.from_user.id != ADMIN_ID: return
    await callback.answer("Формирую отчет...")
    file_path = generate_excel_report()
    with open(file_path, "rb") as file:
        await callback.message.answer_document(document=BufferedInputFile(file.read(), filename=file_path), caption=f"📋 Сводный отчет на {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    os.remove(file_path)

@dp.callback_query(F.data == "admin_broadcast")
async def callback_broadcast(callback: types.CallbackQuery, state: FSMContext):
    if callback.from_user.id != ADMIN_ID: return
    await callback.answer()
    await callback.message.answer("📝 Введите текст объявления для всех сотрудников:\n_(Для отмены введите /start)_")
    await state.set_state(AppStates.waiting_for_admin_broadcast)

@dp.message(AppStates.waiting_for_admin_broadcast)
async def process_broadcast(message: types.Message, state: FSMContext):
    if message.from_user.id != ADMIN_ID: return
    broadcast_text = f"📢 **Объявление от руководителя:**\n\n{message.text}"
    await state.clear()
    
    conn = sqlite3.connect("factory_data.db")
    cursor = conn.cursor()
    cursor.execute("SELECT tg_id FROM users")
    users = cursor.fetchall()
    conn.close()
    
    success_count = 0
    for user in users:
        try:
            await bot.send_message(chat_id=user[0], text=broadcast_text, parse_mode="Markdown")
            success_count += 1
        except Exception: pass
    await message.answer(f"✅ Успешно отправлено сотрудников: {success_count}.", reply_markup=get_admin_keyboard())

@dp.callback_query(F.data == "confirm_reset")
async def callback_confirm_reset(callback: types.CallbackQuery):
    if callback.from_user.id != ADMIN_ID: return
    await callback.answer()
    confirm_keyboard = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="❌ ОТМЕНА", callback_data="cancel_reset"), InlineKeyboardButton(text="🔥 ДА, СТЕРЕТЬ", callback_data="execute_reset")]])
    await callback.message.answer("⚠️ Вы уверены, что хотите удалить все отчеты?", reply_markup=confirm_keyboard)

@dp.callback_query(F.data == "cancel_reset")
async def callback_cancel_reset(callback: types.CallbackQuery):
    if callback.from_user.id != ADMIN_ID: return
    await callback.answer("Сброс отменен")
    await callback.message.delete()

@dp.callback_query(F.data == "execute_reset")
async def callback_execute_reset(callback: types.CallbackQuery):
    if callback.from_user.id != ADMIN_ID: return
    await callback.answer()
    conn = sqlite3.connect("factory_data.db")
    cursor = conn.cursor()
    cursor.execute("DELETE FROM reports")
    conn.commit()
    conn.close()
    await callback.message.answer("🗑️ База данных очищена! Профили сохранены.", reply_markup=get_admin_keyboard())

# --- ПРИЕМ ДАННЫХ ИЗ ФОРМЫ ---
@dp.message(lambda message: message.web_app_data is not None)
async def handle_web_app_data(message: types.Message):
    tg_id = message.from_user.id
    conn = sqlite3.connect("factory_data.db")
    cursor = conn.cursor()
    cursor.execute("SELECT fio FROM users WHERE tg_id = ?", (tg_id,))
    user = cursor.fetchone()
    if not user:
        await message.answer("Ошибка: Профиль не найден.")
        conn.close()
        return

    data = json.loads(message.web_app_data.data)
    current_date = datetime.now().strftime("%Y-%m-%d")
    cursor.execute("""
        INSERT INTO reports (tg_id, date, preparation, lacquered, ready_to_assemble, on_inspection, in_repair, total_done)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (tg_id, current_date, int(data['preparation']), int(data['lacquered']), int(data['ready_to_assemble']), int(data['on_inspection']), int(data['in_repair']), int(data['total_done'])))
    conn.commit()
    conn.close()
    
    report_text = f"✅ **Данные приняты!**\n👤 От: {user[0]}\n📅 Дата: {current_date}\n\n🔹 Подготовка: {data['preparation']} шт.\n🔹 Залачено: {data['lacquered']} шт.\n🔹 Готово к сборке: {data['ready_to_assemble']} шт.\n🔹 На проверке: {data['on_inspection']} шт.\n⚠️ В ремонте: {data['in_repair']} шт.\n🚀 Итого готово: {data['total_done']} шт."
    await message.answer(report_text, parse_mode="Markdown")

# --- ПЛАНИРОВЩИК (Напоминания Пн-Пт в 17:30 и Автоотчет Вс в 21:00) ---
async def check_reminders_and_reports():
    while True:
        now = datetime.now()
        if now.weekday() in range(0, 5) and now.hour == 17 and now.minute == 30:
            try:
                current_date = now.strftime("%Y-%m-%d")
                conn = sqlite3.connect("factory_data.db")
                cursor = conn.cursor()
                cursor.execute("SELECT tg_id FROM users WHERE tg_id NOT IN (SELECT tg_id FROM reports WHERE date = ?)", (current_date,))
                forgetful_users = cursor.fetchall()
                conn.close()
                for user in forgetful_users:
                    try: await bot.send_message(chat_id=user[0], text="⚠️ **Напоминание!** Пожалуйста, отправьте отчет за сегодня через форму в боте!")
                    except Exception: pass
            except Exception as e: print(f"Ошибка напоминаний: {e}")
            await asyncio.sleep(60)

        if now.weekday() == 6 and now.hour == 21 and now.minute == 0:
            try:
                file_path = generate_excel_report()
                with open(file_path, "rb") as file:
                    await bot.send_document(chat_id=ADMIN_ID, document=BufferedInputFile(file.read(), filename=file_path), caption="🗓 **Еженедельный отчет по работе отдела.**")
                os.remove(file_path)
            except Exception as e: print(f"Ошибка автоотчета: {e}")
            await asyncio.sleep(60)
        await asyncio.sleep(30)

async def main():
    init_db()
    
    app = web.Application()
    app.router.add_get('/', handle_web_form)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, '0.0.0.0', 8080)
    await site.start()
    print("🌐 Локальный сервер формы запущен на порту 8080!")

    asyncio.create_task(check_reminders_and_reports())
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
