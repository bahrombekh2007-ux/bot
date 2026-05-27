import asyncio
import copy
import os
import platform
import random
import shutil
import subprocess
import time
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart, Command
from aiogram.types import (
    Message,
    ReplyKeyboardMarkup,
    KeyboardButton,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    CallbackQuery
)
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from docx import Document
from datetime import datetime

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

if load_dotenv:
    load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
if not TOKEN:
    raise RuntimeError("BOT_TOKEN muhit o'zgaruvchisi belgilanmagan. Iltimos .env faylini tekshiring.")

bot = Bot(token=TOKEN)
dp = Dispatcher()

# Vaqtinchalik fayllar saqlanadigan papka
os.makedirs("temp", exist_ok=True)

def cleanup_old_files(directory="temp", max_age_days=3):
    """3 kundan eski fayllarni o'chiradi."""
    max_age_seconds = max_age_days * 24 * 60 * 60
    cutoff = time.time() - max_age_seconds

    for file_name in os.listdir(directory):
        file_path = os.path.join(directory, file_name)
        try:
            if os.path.isfile(file_path) and os.path.getmtime(file_path) < cutoff:
                os.remove(file_path)
        except Exception:
            pass


def get_short_path(path):
    try:
        import ctypes
        from ctypes import wintypes
        GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
        GetShortPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
        GetShortPathNameW.restype = wintypes.DWORD

        output_buf = ctypes.create_unicode_buffer(260)
        result = GetShortPathNameW(path, output_buf, len(output_buf))
        if result and result < len(output_buf):
            return output_buf.value
    except Exception:
        pass
    return path


def convert_doc_to_docx(input_path, output_path):
    """Convert .doc files to .docx using Word or LibreOffice."""
    try:
        import win32com.client
    except ImportError:
        win32com = None
    else:
        win32com = win32com.client

    if win32com:
        word = None
        try:
            word = win32com.Dispatch("Word.Application")
            short_input = get_short_path(input_path)
            short_output = get_short_path(output_path)
            try:
                doc = word.Documents.Open(short_input)
            except Exception:
                doc = word.Documents.Open(input_path)
            try:
                doc.SaveAs2(short_output, FileFormat=16)
            except Exception:
                doc.SaveAs(short_output, FileFormat=16)
            doc.Close(False)
            return output_path
        finally:
            if word:
                word.Quit()

    last_error = None
    if cscript_available():
        try:
            return convert_doc_with_word_cli(input_path, output_path)
        except Exception as e:
            last_error = str(e)

    # Fallback: LibreOffice / OpenOffice
    soffice_executable = None
    for candidate in [
        "soffice",
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files\\OpenOffice 4\\program\\soffice.exe",
        "C:\\Program Files (x86)\\OpenOffice 4\\program\\soffice.exe"
    ]:
        if os.path.isabs(candidate):
            if os.path.exists(candidate):
                soffice_executable = candidate
                break
        else:
            if shutil.which(candidate):
                soffice_executable = candidate
                break

    if soffice_executable:
        output_dir = os.path.dirname(output_path)
        command = [
            soffice_executable,
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            output_dir,
            input_path,
        ]
        try:
            subprocess.run(command, check=True, capture_output=True)
            if os.path.exists(output_path):
                return output_path
        except Exception as e:
            last_error = str(e)

    if platform.system() == "Windows":
        msg = (
            "DOC faylni ochib bo'lmadi. Iltimos, o'zingizga Windows Word yoki LibreOffice o'rnatilganligini tekshiring. "
            "Agar Word bo'lsa, pywin32 o'rnatilgan bo'lishi kerak yoki Word bilan VBScript yordamchi ishlashi kerak; LibreOffice bo'lsa, soffice PATH yoki standart papkada mavjud bo'lishi kerak."
        )
    else:
        msg = (
            "DOC faylni ochib bo'lmadi. Iltimos, LibreOffice o'rnatilganligini tekshiring. "
            "Linuxda `soffice` PATH yoki standart papkada mavjud bo'lishi kerak."
        )
    if last_error:
        msg += f"\n\nTayyorgarlikda xatolik: {last_error}"
    raise RuntimeError(msg)


def convert_doc_with_word_cli(input_path, output_path):
    """Use Word via VBScript to convert .doc to .docx without pywin32."""
    temp_dir = os.path.dirname(output_path)
    script_path = os.path.join(temp_dir, f"convert_{int(time.time())}.vbs")

    short_input_path = get_short_path(input_path)
    short_output_path = get_short_path(output_path)
    short_script_path = get_short_path(script_path)

    input_path_escaped = short_input_path.replace('"', '""')
    output_path_escaped = short_output_path.replace('"', '""')

    script = (
        f"On Error Resume Next\n"
        f"Set objWord = CreateObject(\"Word.Application\")\n"
        f"objWord.Visible = False\n"
        f"objWord.DisplayAlerts = 0\n"
        f"Set objDoc = objWord.Documents.Open(\"{input_path_escaped}\", False, True, False)\n"
        f"objDoc.SaveAs2 \"{output_path_escaped}\", 16\n"
        f"If Err.Number <> 0 Then\n"
        f"    Err.Clear\n"
        f"    objDoc.SaveAs \"{output_path_escaped}\", 16\n"
        f"End If\n"
        f"objDoc.Close False\n"
        f"objWord.Quit\n"
    )

    with open(script_path, "w", encoding="utf-8") as f:
        f.write(script)

    try:
        result = subprocess.run(
            ["cscript", "//NoLogo", short_script_path],
            check=False,
            capture_output=True,
            text=True,
        )
        if result.returncode == 0 and os.path.exists(short_output_path):
            return output_path
        raise RuntimeError(
            f"VBScript konvertatsiyada xato: code={result.returncode}; stdout={result.stdout.strip()}; stderr={result.stderr.strip()}"
        )
    finally:
        try:
            if os.path.exists(script_path):
                os.remove(script_path)
        except Exception:
            pass


def cscript_available():
    return shutil.which("cscript") is not None or os.path.exists(r"C:\\Windows\\System32\\cscript.exe")

# FSM holatlari
class TestStates(StatesGroup):
    setting_count = State()
    testing = State()

# User ma'lumotlarini saqlash
users = {}

# User oldingi test ma'lumotlarini tozalash
def clear_user_test_session(user_id):
    if user_id not in users:
        return

    keys_to_remove = [
        "questions",
        "total_questions",
        "file_name",
        "uploaded_file",
        "selected_questions",
        "total_test",
        "current_index",
        "score",
        "answers",
        "poll_ids",
        "waiting_for_skip",
        "current_poll_message_id",
        "current_poll_id",
        "current_question_index",
        "current_answer_recorded",
        "selected_doc_index"
    ]

    for key in keys_to_remove:
        users[user_id].pop(key, None)

# User xabarlarini saqlash (tozalash uchun)
user_messages = {}

async def safe_delete_message(chat_id, message_id):
    """Xabarni xavfsiz o'chirish"""
    try:
        await bot.delete_message(chat_id, message_id)
    except Exception:
        pass  # Xabar topilmasa yoki o'chirilgan bo'lsa, xatolikni ignor qilish

async def clean_chat(user_id, chat_id, keep_last=0):
    """Chatdagi bot xabarlarini tozalash"""
    if user_id in user_messages:
        messages_to_delete = user_messages[user_id][:-keep_last] if keep_last > 0 else user_messages[user_id]
        
        for msg_id in messages_to_delete:
            await safe_delete_message(chat_id, msg_id)
        
        if keep_last > 0:
            user_messages[user_id] = user_messages[user_id][-keep_last:]
        else:
            user_messages[user_id] = []

async def add_message(user_id, message_id):
    """Xabarni ro'yxatga qo'shish"""
    if user_id not in user_messages:
        user_messages[user_id] = []
    if message_id not in user_messages[user_id]:
        user_messages[user_id].append(message_id)

def parse_docx(file_path):
    """DOCX ichidagi TABLE testlarni o‘qiydi"""
    doc = Document(file_path)
    questions = []

    def add_question(question, options, answer):
        if question and len(options) == 4 and answer in options:
            questions.append({
                "question": question,
                "options": options,
                "answer": answer
            })

    for table in doc.tables:
        rows = []
        for row in table.rows:
            row_texts = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            rows.append(row_texts)

        # 1-column table style with grouped rows (5 rows per question)
        if len(rows) > 0 and all(len(r) == 1 for r in rows):
            flat = [r[0] for r in rows]
            for i in range(0, len(flat), 5):
                block = flat[i:i+5]
                if len(block) == 5:
                    add_question(block[0], block[1:5], block[1])
            continue

        # Standard table rows with 5+ columns
        for row_texts in rows:
            if len(row_texts) >= 5:
                add_question(row_texts[0], row_texts[1:5], row_texts[1])

    # Fallback: paragraph-based blocks of 5 lines
    if not questions:
        lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for i in range(0, len(lines), 5):
            block = lines[i:i+5]
            if len(block) == 5:
                add_question(block[0], block[1:5], block[1])

    return questions

def main_menu_keyboard(user_id=None):
    """Asosiy menyu klaviaturasi"""
    buttons = [
        [KeyboardButton(text="📊 Test natijam"), KeyboardButton(text="🆘 Yordam")],
    ]

    second_row = [KeyboardButton(text="📄 Yangi test")]
    if user_id and user_id in users and users[user_id].get("uploaded_docs"):
        second_row.append(KeyboardButton(text="🔁 Qayta boshlash"))

    buttons.append(second_row)
    return ReplyKeyboardMarkup(
        keyboard=buttons,
        resize_keyboard=True
    )

def get_count_keyboard(total):
    options = []
    if total >= 5:
        options.append(5)
    if total >= 10:
        options.append(10)
    if total >= 15:
        options.append(15)
    if total >= 20:
        options.append(20)
    if total >= 25:
        options.append(25)
    if total >= 30:
        options.append(30)
    if total >= 35:
        options.append(35)
    if total not in options:
        options.append(total)

    count_buttons = []
    row = []
    for i, count in enumerate(options):
        if count < total:
            label = f"📝 {count} ta"
        else:
            label = f"📚 Hammasi ({count} ta)"
        row.append(InlineKeyboardButton(text=label, callback_data=f"count_{count}"))
        if len(row) == 2 or i == len(options) - 1:
            count_buttons.append(row)
            row = []

    count_buttons.append([InlineKeyboardButton(text="✍️ O'zim kiritaman", callback_data="custom_count")])
    count_buttons.append([InlineKeyboardButton(text="❌ Bekor qilish", callback_data="cancel_test")])
    return InlineKeyboardMarkup(inline_keyboard=count_buttons)


def poll_keyboard():
    """Poll test uchun boshqaruv tugmalari"""
    buttons = [
        [
            InlineKeyboardButton(text="⏭ O'tkazib yuborish ➡️", callback_data="skip_question"),
        ],
        [
            InlineKeyboardButton(text="🔴 Testni yakunlash", callback_data="stop_test")
        ]
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons)

@dp.message(CommandStart())
async def start(message: Message, state: FSMContext):
    await state.clear()
    await clean_chat(message.from_user.id, message.chat.id)
    
    msg = await message.answer(
        "🎯 <b>TEST MASTER BOT</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        "🤖 <i>oraliq sesiya uchun yordam!</i>\n\n"
        "✨ <b>Imkoniyatlar:</b>\n"
        "• 🎲 Savollar va variantlar aralash\n"
        "• 📝 Test sonini o'zingiz belgilaysiz\n"
        "• ⏱ Vaqt chegarasisiz\n"
        "• 🧑‍💻@Rustamov_v1\n\n"
        "📄 <b>Boshlash uchun DOCX fayl yuboring!</b>",
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(message.from_user.id)
    )
    await add_message(message.from_user.id, msg.message_id)

@dp.message(F.text == "📄 Yangi test")
async def new_test(message: Message, state: FSMContext):
    await state.clear()
    await clean_chat(message.from_user.id, message.chat.id)
    
    msg = await message.answer(
        "📄 <b>YANGI TEST</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        "Test boshlash uchun DOCX formatdagi\n"
        "faylingizni yuboring 📎\n\n"
        "<i>Format: 5 ustunli jadval</i>",
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(message.from_user.id)
    )
    await add_message(message.from_user.id, msg.message_id)

@dp.message(F.text == "🔁 Qayta boshlash")
async def restart_previous_doc(message: Message, state: FSMContext):
    user_id = message.from_user.id
    if user_id not in users or not users[user_id].get("uploaded_docs"):
        msg = await message.answer(
            "❌ Oldingi DOCX topilmadi. Avval yangi test uchun DOCX yuboring!",
            reply_markup=main_menu_keyboard(user_id)
        )
        await add_message(user_id, msg.message_id)
        return

    docs = users[user_id]["uploaded_docs"]
    buttons = []
    for index, doc in enumerate(docs):
        title = doc["file_name"]
        count = len(doc["questions"])
        buttons.append([
            InlineKeyboardButton(
                text=f"{index + 1}. {title} ({count} ta)",
                callback_data=f"restart_doc_{index}"
            )
        ])

    buttons.append([InlineKeyboardButton(text="❌ Bekor qilish", callback_data="cancel_test")])

    await clean_chat(user_id, message.chat.id)
    msg = await message.answer(
        "🔁 <b>OLDINGI DOCX LAR</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        "Quyidagi fayllardan birini tanlang va testni qayta boshlang:",
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )
    await add_message(user_id, msg.message_id)

@dp.callback_query(F.data.startswith("restart_doc_"))
async def restart_doc(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    if user_id not in users or not users[user_id].get("uploaded_docs"):
        await callback.answer("❌ Oldingi DOCX topilmadi!", show_alert=True)
        return

    try:
        doc_index = int(callback.data.split("_")[-1])
        doc = users[user_id]["uploaded_docs"][doc_index]
    except (ValueError, IndexError):
        await callback.answer("❌ Noto'g'ri tanlov!", show_alert=True)
        return

    users[user_id].update({
        "questions": doc["questions"],
        "total_questions": len(doc["questions"]),
        "file_name": doc["file_name"],
        "selected_doc_index": doc_index
    })

    total = users[user_id]["total_questions"]
    await state.set_state(TestStates.setting_count)

    try:
        await callback.message.edit_text(
            f"📝 <b>TEST SONINI TANLANG</b>\n"
            f"━━━━━━━━━━━━━━━━━━━\n\n"
            f"📚 Mavjud savollar: <b>{total} ta</b>\n\n"
            f"<i>Quyidagi variantlardan birini tanlang</i> 👇",
            parse_mode="HTML",
            reply_markup=get_count_keyboard(total)
        )
    except Exception:
        msg = await callback.message.answer(
            f"📝 <b>TEST SONINI TANLANG</b>\n"
            f"━━━━━━━━━━━━━━━━━━━\n\n"
            f"📚 Mavjud savollar: <b>{total} ta</b>\n\n"
            f"<i>Quyidagi variantlardan birini tanlang</i> 👇",
            parse_mode="HTML",
            reply_markup=get_count_keyboard(total)
        )
        await add_message(user_id, msg.message_id)

    await callback.answer()

@dp.message(F.text == "🆘 Yordam")
async def help(message: Message):
    await clean_chat(message.from_user.id, message.chat.id)
    
    msg = await message.answer(
        "🆘 <b>YORDAM</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        "📌 <b>Test yaratish tartibi:</b>\n"
        "1️⃣ DOCX fayl yuboring\n"
        "2️⃣ Test sonini tanlang\n"
        "⚠️ <b>Muhim eslatma:</b>\n"
        "• Har bir javobdan keyin\n"
        "  'O'tkazib yuborish' tugmasini bosing\n"
        "• Vaqt chegarasi yo'q\n"
        "• Testni istalgan payt yakunlang\n\n"
        "📊 <b>DOCX formati:</b>\n"
        "• 5 ustunli jadval\n"
        "• 1-ustun: Savol\n"
        "• 2-5-ustun: Variantlar\n"
        "• 2-ustun to'g'ri javob\n\n"
        "💡 <b>Taklif bolsa:</b>\n"
        "🧑‍💻@Rustamov_v1",
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(message.from_user.id)
    )
    await add_message(message.from_user.id, msg.message_id)

@dp.message(F.text == "📊 Test natijam")
async def test_result(message: Message):
    await clean_chat(message.from_user.id, message.chat.id)
    user_id = message.from_user.id
    
    if user_id not in users or "results" not in users[user_id]:
        msg = await message.answer(
            "📊 <b>TEST NATIJALARI</b>\n"
            "━━━━━━━━━━━━━━━━━━━\n\n"
            "❌ <b>Natija topilmadi!</b>\n\n"
            "Siz hali test ishlamagansiz yoki\n"
            "natijalar o'chib ketgan.\n\n"
            "📄 Yangi test boshlang!",
            parse_mode="HTML",
            reply_markup=main_menu_keyboard(message.from_user.id)
        )
        await add_message(message.from_user.id, msg.message_id)
        return
    
    results = users[user_id]["results"]
    
    if not results:
        msg = await message.answer(
            "📊 Hali natijalar mavjud emas",
            reply_markup=main_menu_keyboard(message.from_user.id)
        )
        await add_message(message.from_user.id, msg.message_id)
        return
    
    text = "📊 <b>TEST NATIJALARI</b>\n"
    text += "━━━━━━━━━━━━━━━━━━━\n\n"
    
    for i, result in enumerate(results[-5:], 1):
        text += f"<b>{i}.</b> {result['date']}\n"
        text += f"   📝 {result['total']} ta | ✅ {result['score']} ta\n"
        text += f"   📈 {result['percentage']:.1f}% | {result['grade']}\n\n"
    
    text += "<i>Yangi test uchun DOCX yuboring</i> 📄"
    
    msg = await message.answer(
        text,
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(message.from_user.id)
    )
    await add_message(message.from_user.id, msg.message_id)

@dp.message(F.document)
async def handle_doc(message: Message, state: FSMContext):
    await clean_chat(message.from_user.id, message.chat.id)
    doc = message.document
    file_name = doc.file_name.lower()
    
    if not (file_name.endswith(".docx") or file_name.endswith(".doc")):
        msg = await message.answer(
            "❌ <b>Faqat .docx yoki .doc fayl yuboring!</b>",
            parse_mode="HTML",
            reply_markup=main_menu_keyboard(message.from_user.id)
        )
        await add_message(message.from_user.id, msg.message_id)
        return
    
    loading_msg = await message.answer(
        "⏳ <b>Test yuklanmoqda...</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n"
        "🔄 Fayl tahlil qilinmoqda...",
        parse_mode="HTML"
    )
    await add_message(message.from_user.id, loading_msg.message_id)
    
    try:
        cleanup_old_files()
        file = await bot.get_file(doc.file_id)
        downloaded = await bot.download_file(file.file_path)
        
        ext = ".docx"
        if file_name.endswith(".doc"):
            ext = ".doc"

        save_path = os.path.join(
            "temp",
            f"test_{message.from_user.id}_{int(time.time())}{ext}"
        )

        with open(save_path, "wb") as f:
            f.write(downloaded.read())

        if not os.path.exists(save_path):
            await clean_chat(message.from_user.id, message.chat.id)
            msg = await message.answer(
                "❌ Fayl saqlanmadi. Iltimos yana urinib ko'ring.",
                reply_markup=main_menu_keyboard(message.from_user.id)
            )
            await add_message(message.from_user.id, msg.message_id)
            return

        parse_path = save_path
        if file_name.endswith(".doc"):
            converted_path = os.path.splitext(save_path)[0] + ".docx"
            parse_path = convert_doc_to_docx(save_path, converted_path)
        
        questions = parse_docx(parse_path)
        
        if not questions:
            await clean_chat(message.from_user.id, message.chat.id)
            msg = await message.answer(
                "❌ <b>Test topilmadi!</b>\n\n"
                "Fayl formatini tekshiring:\n"
                "• 5 ustunli jadval bo'lishi kerak",
                parse_mode="HTML",
                reply_markup=main_menu_keyboard(message.from_user.id)
            )
            await add_message(message.from_user.id, msg.message_id)
            return
        
        await clean_chat(message.from_user.id, message.chat.id)
        
        msg = await message.answer(
            f"✅ <b>Test muvaffaqiyatli yuklandi!</b>\n"
            f"━━━━━━━━━━━━━━━━━━━\n\n"
            f"📚 Jami savollar: <b>{len(questions)} ta</b>\n"
            f"📄 Fayl: <code>{doc.file_name}</code>\n\n"
            f"<i>Test sonini tanlang...</i>",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="📝 Test sonini kiritish", callback_data="set_count")],
                [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="cancel_doc")]
            ])
        )
        await add_message(message.from_user.id, msg.message_id)

        uploaded_docs = users.get(message.from_user.id, {}).get("uploaded_docs", [])
        uploaded_docs.append({
            "file_name": doc.file_name,
            "file_path": parse_path,
            "questions": questions,
            "uploaded_at": datetime.now().strftime("%d.%m.%Y %H:%M")
        })

        existing_data = users.get(message.from_user.id, {})
        existing_data.update({
            "questions": questions,
            "total_questions": len(questions),
            "file_name": doc.file_name,
            "uploaded_file": parse_path,
            "uploaded_docs": uploaded_docs
        })
        users[message.from_user.id] = existing_data
        
    except Exception as e:
        await clean_chat(message.from_user.id, message.chat.id)
        msg = await message.answer(
            f"❌ <b>Xatolik:</b> <code>{str(e)}</code>",
            parse_mode="HTML",
            reply_markup=main_menu_keyboard(message.from_user.id)
        )
        await add_message(message.from_user.id, msg.message_id)

@dp.callback_query(F.data == "cancel_doc")
async def cancel_doc(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await clean_chat(callback.from_user.id, callback.message.chat.id)
    
    # Xabarni tahrirlash o'rniga yangi xabar yuborish
    try:
        await callback.message.edit_text("❌ Test yuklash bekor qilindi")
    except:
        pass
    
    clear_user_test_session(callback.from_user.id)
    msg = await callback.message.answer(
        "📄 Yangi test uchun DOCX yuboring",
        reply_markup=main_menu_keyboard(callback.from_user.id)
    )
    await add_message(callback.from_user.id, msg.message_id)
    await callback.answer()

@dp.callback_query(F.data == "set_count")
async def set_count(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    
    if user_id not in users:
        await callback.answer("❌ Avval test yuklang!", show_alert=True)
        return
    
    total = users[user_id]["total_questions"]
    
    count_buttons = []
    options = []
    
    if total >= 5:
        options.append(5)
    if total >= 10:
        options.append(10)
    if total >= 15:
        options.append(15)
    if total >= 20:
        options.append(20)
    if total >= 25:
        options.append(25)
    if total >= 30:
        options.append(30)
    if total >= 35:
        options.append(35)
    options.append(total)
    
    row = []
    for i, count in enumerate(options):
        if count < total:
            label = f"📝 {count} ta"
        else:
            label = f"📚 Hammasi ({count} ta)"
        row.append(InlineKeyboardButton(text=label, callback_data=f"count_{count}"))
        
        if len(row) == 2 or i == len(options) - 1:
            count_buttons.append(row)
            row = []
    
    count_buttons.append([InlineKeyboardButton(text="✍️ O'zim kiritaman", callback_data="custom_count")])
    count_buttons.append([InlineKeyboardButton(text="❌ Bekor qilish", callback_data="cancel_test")])
    
    await state.set_state(TestStates.setting_count)
    
    try:
        await callback.message.edit_text(
            f"📝 <b>TEST SONINI TANLANG</b>\n"
            f"━━━━━━━━━━━━━━━━━━━\n\n"
            f"📚 Mavjud savollar: <b>{total} ta</b>\n\n"
            f"<i>Quyidagi variantlardan birini tanlang</i> 👇",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=count_buttons)
        )
    except:
        # Agar edit qilib bo'lmasa, yangi xabar
        msg = await callback.message.answer(
            f"📝 <b>TEST SONINI TANLANG</b>\n"
            f"━━━━━━━━━━━━━━━━━━━\n\n"
            f"📚 Mavjud savollar: <b>{total} ta</b>\n\n"
            f"<i>Quyidagi variantlardan birini tanlang</i> 👇",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=count_buttons)
        )
        await add_message(callback.from_user.id, msg.message_id)
    
    await callback.answer()

@dp.callback_query(F.data.startswith("count_"))
async def select_count(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    
    if user_id not in users:
        await callback.answer("❌ Xatolik!", show_alert=True)
        return
    
    count = int(callback.data.split("_")[1])
    
    # Xabarni o'chirishga harakat qilish
    await safe_delete_message(callback.message.chat.id, callback.message.message_id)
    
    await start_test(callback.message, user_id, count, state)
    await callback.answer(f"✅ {count} ta test boshlandi!")

@dp.callback_query(F.data == "custom_count")
async def custom_count_prompt(callback: CallbackQuery):
    try:
        await callback.message.edit_text(
            "✍️ <b>TEST SONINI KIRITING</b>\n"
            "━━━━━━━━━━━━━━━━━━━\n\n"
            "Iltimos, nechta test bo'lishini\n"
            "raqamda yozing:\n"
            "<i>Masalan: 15, 25, 30...</i>\n\n"
            "❌ Bekor qilish uchun /cancel",
            parse_mode="HTML"
        )
    except:
        pass
    await callback.answer()

@dp.message(TestStates.setting_count)
async def process_custom_count(message: Message, state: FSMContext):
    user_id = message.from_user.id
    
    if message.text.startswith('/'):
        return
    
    await clean_chat(message.from_user.id, message.chat.id)
    
    if user_id not in users:
        msg = await message.answer(
            "❌ Avval test yuklang!",
            reply_markup=main_menu_keyboard(message.from_user.id)
        )
        await add_message(message.from_user.id, msg.message_id)
        await state.clear()
        return
    
    try:
        count = int(message.text)
        total = users[user_id]["total_questions"]
        
        if count < 1 or count > total:
            msg = await message.answer(
                f"❌ <b>Noto'g'ri son!</b>\n"
                f"1 dan {total} gacha raqam kiriting",
                parse_mode="HTML",
                reply_markup=main_menu_keyboard(message.from_user.id)
            )
            await add_message(message.from_user.id, msg.message_id)
            return
        
        await start_test(message, user_id, count, state)
        
    except ValueError:
        msg = await message.answer(
            "❌ <b>Raqam kiriting!</b>\n"
            "Masalan: 10, 15, 20...",
            parse_mode="HTML",
            reply_markup=main_menu_keyboard(message.from_user.id)
        )
        await add_message(message.from_user.id, msg.message_id)

async def start_test(message, user_id, count, state: FSMContext):
    """Testni boshlash"""
    await clean_chat(user_id, message.chat.id)
    
    all_questions = copy.deepcopy(users[user_id]["questions"])
    random.shuffle(all_questions)
    selected_questions = all_questions[:count]
    
    for q in selected_questions:
        random.shuffle(q["options"])
    
    users[user_id].update({
        "selected_questions": selected_questions,
        "total_test": count,
        "current_index": 0,
        "score": 0,
        "answers": [],
        "poll_ids": [],
        "waiting_for_skip": False,
        "current_answer_recorded": False,
        "current_poll_message_id": None
    })
    
    await state.set_state(TestStates.testing)
    
    msg = await message.answer(
        f"🚀 <b>TEST BOSHLANDI!</b>\n"
        f"━━━━━━━━━━━━━━━━━━━\n\n"
        f"📝 Jami: <b>{count} ta savol</b>\n"
        f"⏱ Vaqt chegarasi yo'q\n"
        f"🎲 Savollar aralash holda\n\n"
        f"⚠️ <b>Muhim:</b> Javob berganingizdan keyin\n"
        f"'<i>O'tkazib yuborish</i>' tugmasini bosing!\n\n"
        f"<i>Omad!</i> 🍀",
        parse_mode="HTML"
    )
    await add_message(user_id, msg.message_id)
    
    await asyncio.sleep(2)
    await clean_chat(user_id, message.chat.id)
    await send_poll_question(message.chat.id, user_id)


def normalize_poll_text(text, max_length=100):
    text = str(text).strip()
    if len(text) > max_length:
        return text[: max_length - 1] + "…"
    return text

async def send_poll_question(chat_id, user_id):
    """Savolni Telegram Poll sifatida yuborish"""
    data = users[user_id]
    index = data["current_index"]
    question_data = data["selected_questions"][index]

    options = [normalize_poll_text(opt, 100) for opt in question_data["options"]]
    answer_text = normalize_poll_text(question_data["answer"], 100)

    try:
        correct_option_id = options.index(answer_text)
    except ValueError:
        correct_option_id = 0
        options[0] = answer_text

    question_text = str(question_data["question"]).strip()
    if len(question_text) > 300:
        question_text = question_text[:297] + "..."

    explanation = f"✅ To'g'ri javob: {answer_text}"

    # Avvalgi poll xabarini o'chirish
    if data.get("current_poll_message_id"):
        await safe_delete_message(chat_id, data["current_poll_message_id"])
    
    # Eski xabarlarni tozalash
    await clean_chat(user_id, chat_id)
    
    # Yangi poll yuborish
    poll_message = await bot.send_poll(
        chat_id=chat_id,
        question=f"📝 {index + 1}/{data['total_test']}\n\n{question_text}",
        options=options,
        type="quiz",
        correct_option_id=correct_option_id,
        explanation=explanation,
        is_anonymous=False,
        reply_markup=poll_keyboard()
    )
    
    data["poll_ids"].append(poll_message.poll.id)
    data["current_poll_id"] = poll_message.poll.id
    data["current_question_index"] = index
    data["current_poll_message_id"] = poll_message.message_id
    data["waiting_for_skip"] = True
    data["current_answer_recorded"] = False

@dp.poll_answer()
async def handle_poll_answer(poll_answer):
    """Poll javobini qayta ishlash"""
    user_id = poll_answer.user.id
    
    if user_id not in users or "selected_questions" not in users[user_id]:
        return
    
    data = users[user_id]
    if getattr(poll_answer, "poll_id", None) != data.get("current_poll_id"):
        return
    if data.get("current_answer_recorded", False):
        return
    if not poll_answer.option_ids:
        return
    
    current_index = data.get("current_question_index", 0)
    if current_index >= len(data["selected_questions"]):
        return
    
    question_data = data["selected_questions"][current_index]
    selected_option = question_data["options"][poll_answer.option_ids[0]]
    is_correct = selected_option == question_data["answer"]
    
    data["answers"].append({
        "question": question_data["question"],
        "user_answer": selected_option,
        "correct_answer": question_data["answer"],
        "is_correct": is_correct,
        "options": question_data["options"]
    })
    
    if is_correct:
        data["score"] += 1
    
    data["current_answer_recorded"] = True
    
    try:
        status_msg = await bot.send_message(
            user_id,
            f"{'✅' if is_correct else '❌'} <b>Javobingiz qabul qilindi!</b>\n\n"
            f"<i>Davom etish uchun 'O'tkazib yuborish' tugmasini bosing</i> 👇",
            parse_mode="HTML"
        )
        await add_message(user_id, status_msg.message_id)
    except:
        pass

@dp.callback_query(F.data == "skip_question")
async def skip_question(callback: CallbackQuery, state: FSMContext):
    """Savolni o'tkazib yuborish"""
    user_id = callback.from_user.id
    
    if user_id not in users:
        await callback.answer("❌ Test topilmadi!", show_alert=True)
        return
    
    data = users[user_id]
    
    if not data.get("waiting_for_skip", False):
        await callback.answer("⏳ Avval javob bering!", show_alert=True)
        return
    
    current_index = data.get("current_question_index", 0)
    
    if not data.get("current_answer_recorded", False):
        question_data = data["selected_questions"][current_index]
        data["answers"].append({
            "question": question_data["question"],
            "user_answer": "Javob berilmadi",
            "correct_answer": question_data["answer"],
            "is_correct": False
        })

    data["waiting_for_skip"] = False
    data["current_answer_recorded"] = False
    data["current_index"] += 1
    
    if data["current_index"] >= data["total_test"]:
        await clean_chat(user_id, callback.message.chat.id)
        await safe_delete_message(callback.message.chat.id, callback.message.message_id)
        await show_poll_results(callback.message.chat.id, user_id)
        await state.clear()
        await callback.answer("✅ Test yakunlandi!")
        return
    
    await callback.answer("⏭ Keyingi savol...")
    await clean_chat(user_id, callback.message.chat.id)
    await send_poll_question(callback.message.chat.id, user_id)

@dp.callback_query(F.data == "stop_test")
async def stop_poll_test(callback: CallbackQuery, state: FSMContext):
    """Testni yakunlash"""
    user_id = callback.from_user.id
    
    if user_id not in users:
        await callback.answer("❌ Test topilmadi!", show_alert=True)
        return
    
    data = users[user_id]
    current_index = data.get("current_question_index", 0)
    
    if len(data["answers"]) <= current_index and current_index < data["total_test"]:
        question_data = data["selected_questions"][current_index]
        data["answers"].append({
            "question": question_data["question"],
            "user_answer": "Test yakunlandi",
            "correct_answer": question_data["answer"],
            "is_correct": False
        })
    
    await clean_chat(user_id, callback.message.chat.id)
    await safe_delete_message(callback.message.chat.id, callback.message.message_id)
    await show_poll_results(callback.message.chat.id, user_id, stopped=True)
    await state.clear()
    await callback.answer("🔴 Test yakunlandi")

async def show_poll_results(chat_id, user_id, stopped=False):
    """Poll test natijalarini ko'rsatish"""
    data = users[user_id]
    score = data["score"]
    total = data["total_test"]
    answered = len(data["answers"])
    skipped = max(total - answered, 0)
    incorrect = max(total - score, 0)
    percentage = (score / total * 100) if total > 0 else 0
    
    if percentage >= 90:
        grade = "🏆 A'lo"
        emoji = "🌟"
    elif percentage >= 75:
        grade = "🎉 Yaxshi"
        emoji = "👏"
    elif percentage >= 60:
        grade = "👍 Qoniqarli"
        emoji = "💪"
    else:
        grade = "📚 O'qish kerak"
        emoji = "📖"
    
    result_text = (
        f"{emoji} <b>TEST NATIJASI</b>\n"
        f"━━━━━━━━━━━━━━━━━━━\n\n"
        f"📊 <b>Statistika:</b>\n"
        f"• Jami savollar: <b>{total} ta</b>\n"
        f"• Javob berilgan: <b>{answered} ta</b>\n"
        f"• O'tkazib yuborilgan: <b>{skipped} ta</b>\n\n"
        f"✅ To'g'ri: <b>{score} ta</b>\n"
        f"❌ Noto'g'ri: <b>{incorrect} ta</b>\n"
        f"📈 Foiz: <b>{percentage:.1f}%</b>\n"
        f"🏆 Baho: <b>{grade}</b>\n\n"
        f"<i>{'⚠️ Test vaqtidan oldin yakunlandi' if stopped else '🎊 Test muvaffaqiyatli yakunlandi!'}</i>"
    )
    
    if "results" not in data:
        data["results"] = []
    
    data["results"].append({
        "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "total": total,
        "score": score,
        "percentage": percentage,
        "grade": grade
    })
    
    # Avval eski xabarlarni tozalash
    await clean_chat(user_id, chat_id)
    
    msg = await bot.send_message(
        chat_id,
        result_text,
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [
                InlineKeyboardButton(text="🔄 Qayta urinish", callback_data="retry_poll"),
                InlineKeyboardButton(text="📋 Batafsil", callback_data="poll_details")
            ],
            [
                InlineKeyboardButton(text="📊 Barcha natijalar", callback_data="all_results"),
                InlineKeyboardButton(text="🏠 Menyu", callback_data="main_menu")
            ]
        ])
    )
    await add_message(user_id, msg.message_id)

@dp.callback_query(F.data == "retry_poll")
async def retry_poll_test(callback: CallbackQuery, state: FSMContext):
    """Testni qayta boshlash"""
    user_id = callback.from_user.id
    
    if user_id not in users:
        await callback.answer("❌ Test ma'lumotlari topilmadi", show_alert=True)
        return
    
    data = users[user_id]
    
    selected = data["selected_questions"].copy()
    random.shuffle(selected)
    
    for q in selected:
        options = q["options"].copy()
        random.shuffle(options)
        q["options"] = options
    
    data.update({
        "selected_questions": selected,
        "current_index": 0,
        "score": 0,
        "answers": [],
        "poll_ids": [],
        "waiting_for_skip": False,
        "current_answer_recorded": False
    })
    
    await state.set_state(TestStates.testing)
    await clean_chat(user_id, callback.message.chat.id)
    await safe_delete_message(callback.message.chat.id, callback.message.message_id)
    
    msg = await callback.message.answer(
        "🔄 <b>TEST QAYTA BOSHLANDI!</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        f"📝 Jami: <b>{data['total_test']} ta savol</b>\n"
        "🎲 Savollar qayta aralashtirildi\n\n"
        "⚠️ <b>Muhim:</b> Javob berganingizdan keyin\n"
        "'<i>O'tkazib yuborish</i>' tugmasini bosing!\n\n"
        "<i>Omad!</i> 🍀",
        parse_mode="HTML"
    )
    await add_message(user_id, msg.message_id)
    
    await asyncio.sleep(2)
    await clean_chat(user_id, callback.message.chat.id)
    await send_poll_question(callback.message.chat.id, user_id)
    await callback.answer("✅ Test qayta boshlandi")

@dp.callback_query(F.data == "poll_details")
async def show_poll_details(callback: CallbackQuery):
    """Batafsil natijalarni ko'rsatish"""
    user_id = callback.from_user.id
    
    if user_id not in users or "answers" not in users[user_id]:
        await callback.answer("❌ Ma'lumot topilmadi", show_alert=True)
        return
    
    answers = users[user_id]["answers"]
    
    if not answers:
        await callback.answer("❌ Javoblar mavjud emas", show_alert=True)
        return
    
    details = "📋 <b>BATAFSIL NATIJALAR</b>\n"
    details += "━━━━━━━━━━━━━━━━━━━\n\n"
    
    for i, ans in enumerate(answers, 1):
        status = "✅" if ans["is_correct"] else "❌"
        question_text = ans["question"][:60]
        if len(ans["question"]) > 60:
            question_text += "..."
        
        details += f"<b>{i}.</b> {question_text}\n"
        details += f"   {status} Siz: <i>{ans['user_answer']}</i>\n"
        if not ans["is_correct"]:
            details += f"   ✅ To'g'ri: <i>{ans['correct_answer']}</i>\n"
        details += "\n"
        
        if len(details) > 3500:
            msg = await callback.message.answer(details, parse_mode="HTML")
            await add_message(user_id, msg.message_id)
            details = ""
    
    if details:
        msg = await callback.message.answer(details, parse_mode="HTML")
        await add_message(user_id, msg.message_id)
    
    await callback.answer("📋 Batafsil ko'rsatildi")

@dp.callback_query(F.data == "all_results")
async def show_all_results(callback: CallbackQuery):
    """Barcha test natijalarini ko'rsatish"""
    user_id = callback.from_user.id
    
    if user_id not in users or "results" not in users[user_id]:
        await callback.answer("❌ Natijalar topilmadi", show_alert=True)
        return
    
    results = users[user_id]["results"]
    
    if not results:
        await callback.answer("❌ Natijalar mavjud emas", show_alert=True)
        return
    
    await clean_chat(user_id, callback.message.chat.id)
    await safe_delete_message(callback.message.chat.id, callback.message.message_id)
    
    text = "📊 <b>BARCHA TEST NATIJALARI</b>\n"
    text += "━━━━━━━━━━━━━━━━━━━\n\n"
    
    for i, result in enumerate(results[-10:], 1):
        text += f"<b>{i}.</b> {result['date']}\n"
        text += f"   📝 {result['total']} ta | ✅ {result['score']} ta\n"
        text += f"   📈 {result['percentage']:.1f}% | {result['grade']}\n\n"
    
    msg = await callback.message.answer(
        text,
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(user_id)
    )
    await add_message(user_id, msg.message_id)
    await callback.answer()

@dp.callback_query(F.data == "main_menu")
async def back_to_main_menu(callback: CallbackQuery, state: FSMContext):
    """Bosh menyuga qaytish - XATOLIK TUZATILDI"""
    await state.clear()
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id
    
    # Oldingi test / fayl ma'lumotlarini tozalash
    clear_user_test_session(user_id)
    
    # Xabarni xavfsiz o'chirish
    await safe_delete_message(chat_id, callback.message.message_id)
    
    # Chatni tozalash
    await clean_chat(user_id, chat_id)
    
    # Yangi menyu xabarini yuborish
    msg = await bot.send_message(
        chat_id,
        "🏠 <b>BOSH MENYU</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        "📄 Yangi test boshlash uchun DOCX fayl yuboring\n\n"
        "<i>Menu tugmalaridan foydalaning</i> 👇",
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(user_id)
    )
    await add_message(user_id, msg.message_id)
    await callback.answer("🏠 Bosh menyu")

@dp.callback_query(F.data == "cancel_test")
async def cancel_test_config(callback: CallbackQuery, state: FSMContext):
    """Test sozlashni bekor qilish"""
    await state.clear()
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id
    
    clear_user_test_session(user_id)
    await clean_chat(user_id, chat_id)
    await safe_delete_message(chat_id, callback.message.message_id)
    
    msg = await bot.send_message(
        chat_id,
        "❌ Test sozlash bekor qilindi\n\n"
        "📄 Yangi test uchun DOCX yuboring",
        reply_markup=main_menu_keyboard(user_id)
    )
    await add_message(user_id, msg.message_id)
    await callback.answer()

async def main():
    print("🚀 ishga tushdi...")

    
    try:
        cleanup_old_files()
        await dp.start_polling(bot)
    except Exception as e:
        print(f"❌ Bot ishlashida xatolik: {e}")

if __name__ == "__main__":
    asyncio.run(main())