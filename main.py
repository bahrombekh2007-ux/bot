import asyncio
import copy
import os
import platform
import random
import re
import shutil
import subprocess
import time
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart
from aiogram.types import (
    Message,
    ReplyKeyboardMarkup,
    KeyboardButton,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    CallbackQuery,
)
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from docx import Document
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

TOKEN = "8964353995:AAFgRTesY5nYBku_fyuFMNLQ2VW_hPMzOrg"
bot = Bot(token=TOKEN)
dp = Dispatcher()

os.makedirs("temp", exist_ok=True)

# ─────────────────────────── YORDAMCHI FUNKSIYALAR ───────────────────────────

def cleanup_old_files(directory="temp", max_age_days=3):
    """3 kundan eski vaqtinchalik fayllarni o'chiradi."""
    cutoff = time.time() - max_age_days * 86400
    for name in os.listdir(directory):
        path = os.path.join(directory, name)
        try:
            if os.path.isfile(path) and os.path.getmtime(path) < cutoff:
                os.remove(path)
        except Exception:
            pass


def get_short_path(path):
    try:
        import ctypes
        from ctypes import wintypes
        fn = ctypes.windll.kernel32.GetShortPathNameW
        fn.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
        fn.restype = wintypes.DWORD
        buf = ctypes.create_unicode_buffer(260)
        result = fn(path, buf, len(buf))
        if result and result < len(buf):
            return buf.value
    except Exception:
        pass
    return path


def cscript_available():
    return shutil.which("cscript") or os.path.exists(r"C:\Windows\System32\cscript.exe")


def convert_doc_with_word_cli(input_path, output_path):
    temp_dir = os.path.dirname(output_path)
    script_path = os.path.join(temp_dir, f"conv_{int(time.time())}.vbs")
    inp = get_short_path(input_path).replace('"', '""')
    out = get_short_path(output_path).replace('"', '""')
    script = (
        'On Error Resume Next\n'
        'Set w = CreateObject("Word.Application")\n'
        'w.Visible = False : w.DisplayAlerts = 0\n'
        f'Set d = w.Documents.Open("{inp}", False, True, False)\n'
        f'd.SaveAs2 "{out}", 16\n'
        'If Err.Number <> 0 Then Err.Clear\n'
        f'd.SaveAs "{out}", 16\n'
        'End If\n'
        'd.Close False : w.Quit\n'
    )
    with open(script_path, "w", encoding="utf-8") as f:
        f.write(script)
    try:
        r = subprocess.run(
            ["cscript", "//NoLogo", get_short_path(script_path)],
            check=False, capture_output=True, text=True,
        )
        if r.returncode == 0 and os.path.exists(output_path):
            return output_path
        raise RuntimeError(f"VBScript xato: {r.returncode} | {r.stdout.strip()}")
    finally:
        try:
            os.remove(script_path)
        except Exception:
            pass


def convert_doc_to_docx(input_path, output_path):
    """DOC → DOCX: Word yoki LibreOffice orqali."""
    try:
        import win32com.client as wc
        word = None
        try:
            word = wc.Dispatch("Word.Application")
            doc = word.Documents.Open(get_short_path(input_path))
            doc.SaveAs2(get_short_path(output_path), FileFormat=16)
            doc.Close(False)
            return output_path
        finally:
            if word:
                word.Quit()
    except ImportError:
        pass

    if cscript_available():
        try:
            return convert_doc_with_word_cli(input_path, output_path)
        except Exception:
            pass

    for soffice in [
        "soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]:
        found = os.path.exists(soffice) if os.path.isabs(soffice) else shutil.which(soffice)
        if found:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "docx",
                 "--outdir", os.path.dirname(output_path), input_path],
                check=True, capture_output=True,
            )
            if os.path.exists(output_path):
                return output_path

    raise RuntimeError(
        "DOC faylni ochib bo'lmadi. Word yoki LibreOffice o'rnatilganligini tekshiring."
    )


# ─────────────────────────────── PARSERLAR ───────────────────────────────────

def parse_docx(file_path):
    """
    DOCX fayldan savollarni o'qiydi.
    Qo'llab-quvvatlanadi:
      • 5 ustunli jadval (savol | A | B | C | D)
      • 1 ustunli jadval (har 5 qator bir savol)
      • Paragraf asosida (har 5 qator bir savol)
    """
    doc = Document(file_path)
    questions = []

    def add_q(question, options, answer):
        if question and len(options) == 4 and answer in options:
            questions.append({"question": question, "options": options, "answer": answer})

    for table in doc.tables:
        rows = [[c.text.strip() for c in row.cells if c.text.strip()] for row in table.rows]
        rows = [r for r in rows if r]

        if rows and all(len(r) == 1 for r in rows):
            flat = [r[0] for r in rows]
            for i in range(0, len(flat), 5):
                b = flat[i:i+5]
                if len(b) == 5:
                    add_q(b[0], b[1:5], b[1])
            continue

        for r in rows:
            if len(r) >= 5:
                add_q(r[0], r[1:5], r[1])

    if not questions:
        lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for i in range(0, len(lines), 5):
            b = lines[i:i+5]
            if len(b) == 5:
                add_q(b[0], b[1:5], b[1])

    return questions


def _read_txt(file_path):
    """TXT faylni turli kodlashlarda o'qiydi."""
    for enc in ("utf-8-sig", "utf-8", "cp1251", "windows-1251", "latin-1"):
        try:
            with open(file_path, "r", encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, LookupError):
            pass
    return ""


def _parse_plus_minus(lines):
    """
    Standart ? + - formatini parse qiladi:

        ? Savol matni ?
        + To'g'ri javob
        - Noto'g'ri 1
        - Noto'g'ri 2
        - Noto'g'ri 3
    """
    questions = []
    current_q = None
    correct = None
    opts = []

    def flush():
        nonlocal current_q, correct, opts
        if current_q and correct and len(opts) >= 2:
            while len(opts) < 4:
                opts.append(f"Variant {len(opts) + 1}")
            if correct not in opts:
                opts.insert(0, correct)
            questions.append({
                "question": current_q,
                "options": opts[:4],
                "answer": correct,
            })

    for raw in lines:
        line = raw.strip()
        if not line:
            continue

        if line.startswith("?"):
            flush()
            # "? 1. Savol matni ?" → "1. Savol matni?"
            q = line[1:].strip()
            if q.endswith(" ?"):
                q = q[:-2].strip()
            elif q.endswith("?") and len(q) > 1 and q[-2] == " ":
                q = q[:-1].strip()
            current_q = q
            correct = None
            opts = []

        elif line.startswith("+ ") or line == "+":
            ans = line[1:].strip()
            correct = ans
            if ans not in opts:
                opts.append(ans)

        elif line.startswith("- ") or line == "-":
            ans = line[1:].strip()
            if ans and ans not in opts:
                opts.append(ans)

    flush()
    return questions


def _parse_numbered(lines):
    """
    Raqamli A/B/C/D formatini parse qiladi:

        1. Savol matni
        A) To'g'ri javob
        B) Noto'g'ri 1
        C) Noto'g'ri 2
        D) Noto'g'ri 3
        Javob: A
    """
    questions = []
    current_q = None
    opts_dict = {}
    correct_letter = None

    def flush():
        nonlocal current_q, opts_dict, correct_letter
        if current_q and opts_dict and correct_letter:
            ul = correct_letter.upper()
            if ul in opts_dict:
                correct_ans = opts_dict[ul]
                options = list(opts_dict.values())[:4]
                while len(options) < 4:
                    options.append(f"Variant {len(options) + 1}")
                if correct_ans not in options:
                    options[0] = correct_ans
                questions.append({
                    "question": current_q,
                    "options": options,
                    "answer": correct_ans,
                })

    for raw in lines:
        line = raw.strip()
        if not line:
            continue

        m = re.match(r"^(\d+)[.)]\s+(.+)$", line)
        if m:
            flush()
            current_q = m.group(2).strip()
            opts_dict = {}
            correct_letter = None
            continue

        m = re.match(r"^([A-Da-d])[.)]\s+(.+)$", line)
        if m:
            opts_dict[m.group(1).upper()] = m.group(2).strip()
            continue

        m = re.match(
            r"^(?:Javob|To'g'ri\s+javob|Answer|Ans)[:\s]+([A-Da-d])",
            line, re.IGNORECASE
        )
        if m:
            correct_letter = m.group(1).upper()

    flush()
    return questions


def _parse_pipe(lines):
    """
    Pipe | ajratuvchi formatni parse qiladi:

        Savol|Javob A|Javob B|Javob C|Javob D
        (birinchi variant to'g'ri)
    """
    questions = []
    for line in lines:
        parts = [p.strip() for p in line.split("|")]
        if len(parts) == 5 and parts[0]:
            questions.append({
                "question": parts[0],
                "options": parts[1:5],
                "answer": parts[1],
            })
    return questions


def parse_txt(file_path):
    """
    TXT faylni avtomatik format aniqlab parse qiladi.

    Qo'llab-quvvatlanadigan formatlar:
    ┌─────────────────────────────────────────────────────┐
    │  FORMAT 1  (? + - belgilari)                        │
    │  ? Savol matni ?                                    │
    │  + To'g'ri javob                                    │
    │  - Noto'g'ri javob 1                                │
    │  - Noto'g'ri javob 2                                │
    │  - Noto'g'ri javob 3                                │
    ├─────────────────────────────────────────────────────┤
    │  FORMAT 2  (raqamli A/B/C/D)                        │
    │  1. Savol matni                                     │
    │  A) To'g'ri javob                                   │
    │  B) Noto'g'ri 1                                     │
    │  C) Noto'g'ri 2                                     │
    │  D) Noto'g'ri 3                                     │
    │  Javob: A                                           │
    ├─────────────────────────────────────────────────────┤
    │  FORMAT 3  (pipe | ajratuvchi)                      │
    │  Savol|Javob A|Javob B|Javob C|Javob D              │
    └─────────────────────────────────────────────────────┘
    """
    content = _read_txt(file_path)
    if not content:
        return []

    lines = content.splitlines()
    non_empty = [l.strip() for l in lines if l.strip()]

    if not non_empty:
        return []

    # Format aniqlanishi
    has_question_marker = any(l.startswith("?") for l in non_empty)
    has_plus = any(l.startswith("+") for l in non_empty)
    has_pipe = any("|" in l for l in non_empty)
    has_numbered = any(re.match(r"^\d+[.)]\s+", l) for l in non_empty)
    has_abcd = any(re.match(r"^[A-Da-d][.)]\s+", l) for l in non_empty)

    if has_question_marker or has_plus:
        result = _parse_plus_minus(lines)
        if result:
            return result

    if has_numbered and has_abcd:
        result = _parse_numbered(lines)
        if result:
            return result

    if has_pipe:
        result = _parse_pipe(non_empty)
        if result:
            return result

    return []


# ─────────────────────────── FSM & GLOBAL STATE ──────────────────────────────

class TestStates(StatesGroup):
    setting_count = State()
    testing = State()


users: dict = {}
user_messages: dict = {}


def clear_user_test_session(user_id):
    if user_id not in users:
        return
    for key in [
        "questions", "total_questions", "file_name", "uploaded_file",
        "selected_questions", "total_test", "current_index", "score",
        "answers", "poll_ids", "waiting_for_skip", "current_poll_message_id",
        "current_poll_id", "current_question_index", "current_answer_recorded",
        "selected_doc_index",
    ]:
        users[user_id].pop(key, None)


async def safe_delete_message(chat_id, message_id):
    try:
        await bot.delete_message(chat_id, message_id)
    except Exception:
        pass


async def clean_chat(user_id, chat_id, keep_last=0):
    msgs = user_messages.get(user_id, [])
    to_del = msgs[:-keep_last] if keep_last else msgs
    for mid in to_del:
        await safe_delete_message(chat_id, mid)
    user_messages[user_id] = msgs[-keep_last:] if keep_last else []


async def add_message(user_id, message_id):
    user_messages.setdefault(user_id, [])
    if message_id not in user_messages[user_id]:
        user_messages[user_id].append(message_id)


# ──────────────────────────── KLAVIATURALAR ───────────────────────────────────

def main_menu_keyboard(user_id=None):
    row1 = [KeyboardButton(text="📊 Test natijam"), KeyboardButton(text="🆘 Yordam")]
    row2 = [KeyboardButton(text="📄 Yangi test")]
    if user_id and users.get(user_id, {}).get("uploaded_docs"):
        row2.append(KeyboardButton(text="🔁 Qayta boshlash"))
    return ReplyKeyboardMarkup(keyboard=[row1, row2], resize_keyboard=True)


def get_count_keyboard(total):
    opts = [n for n in (5, 10, 15, 20, 25, 30, 35) if total >= n]
    if total not in opts:
        opts.append(total)
    buttons, row = [], []
    for i, c in enumerate(opts):
        label = f"📚 Hammasi ({c} ta)" if c == total else f"📝 {c} ta"
        row.append(InlineKeyboardButton(text=label, callback_data=f"count_{c}"))
        if len(row) == 2 or i == len(opts) - 1:
            buttons.append(row)
            row = []
    buttons.append([InlineKeyboardButton(text="✍️ O'zim kiritaman", callback_data="custom_count")])
    buttons.append([InlineKeyboardButton(text="❌ Bekor qilish", callback_data="cancel_test")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def poll_keyboard():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⏭ O'tkazib yuborish ➡️", callback_data="skip_question")],
        [InlineKeyboardButton(text="🔴 Testni yakunlash", callback_data="stop_test")],
    ])


# ─────────────────────────── HANDLER: START ───────────────────────────────────

@dp.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    await clean_chat(message.from_user.id, message.chat.id)
    msg = await message.answer(
        "🎯 <b>TEST MASTER BOT</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        "🤖 <i>Oraliq va sesiya uchun yordam!</i>\n\n"
        "✨ <b>Imkoniyatlar:</b>\n"
        "• 🎲 Savollar va variantlar aralash\n"
        "• 📝 Test sonini o'zingiz belgilaysiz\n"
        "• ⏱ Vaqt chegarasisiz\n"
        "• 🧑‍💻 @Rustamov_v1\n\n"
        "📎 <b>Qabul qilinadigan fayl turlari:</b>\n"
        "• <code>.docx</code> — Word jadval\n"
        "• <code>.doc</code>  — Eski Word\n"
        "• <code>.txt</code>  — Matnli fayl (? + - yoki A/B/C/D)\n\n"
        "📎 <b>Boshlash uchun fayl yuboring!</b>",
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(message.from_user.id),
    )
    await add_message(message.from_user.id, msg.message_id)


# ─────────────────────────── HANDLER: MENYU ───────────────────────────────────

@dp.message(F.text == "📄 Yangi test")
async def new_test(message: Message, state: FSMContext):
    await state.clear()
    await clean_chat(message.from_user.id, message.chat.id)
    msg = await message.answer(
        "📄 <b>YANGI TEST</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        "Fayl yuboring 📎\n\n"
        "📘 <b>DOCX:</b> 5 ustunli jadval\n"
        "📝 <b>TXT format 1</b> (? + - belgilari):\n"
        "<code>? Savol matni ?\n"
        "+ To'g'ri javob\n"
        "- Noto'g'ri 1\n"
        "- Noto'g'ri 2\n"
        "- Noto'g'ri 3</code>\n\n"
        "📝 <b>TXT format 2</b> (A/B/C/D):\n"
        "<code>1. Savol matni\n"
        "A) To'g'ri javob\n"
        "B) Noto'g'ri 1\n"
        "C) Noto'g'ri 2\n"
        "D) Noto'g'ri 3\n"
        "Javob: A</code>",
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(message.from_user.id),
    )
    await add_message(message.from_user.id, msg.message_id)


@dp.message(F.text == "🆘 Yordam")
async def cmd_help(message: Message):
    await clean_chat(message.from_user.id, message.chat.id)
    msg = await message.answer(
        "🆘 <b>YORDAM</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        "📌 <b>Tartibi:</b>\n"
        "1️⃣ DOCX yoki TXT fayl yuboring\n"
        "2️⃣ Test sonini tanlang\n"
        "3️⃣ Javob berib → 'O'tkazib yuborish'ni bosing\n\n"
        "─────────────────────\n"
        "📘 <b>DOCX formati:</b>\n"
        "• 5 ustunli jadval\n"
        "• 1-ustun: Savol\n"
        "• 2–5-ustun: Variantlar (2-ustun to'g'ri)\n\n"
        "─────────────────────\n"
        "📝 <b>TXT format 1</b> (? + -):\n"
        "<code>? Savol ?\n+ To'g'ri\n- Noto'g'ri 1\n- Noto'g'ri 2\n- Noto'g'ri 3</code>\n\n"
        "📝 <b>TXT format 2</b> (A/B/C/D):\n"
        "<code>1. Savol\nA) To'g'ri\nB) Noto'g'ri 1\nC) Noto'g'ri 2\nD) Noto'g'ri 3\nJavob: A</code>\n\n"
        "─────────────────────\n"
        "⚠️ Har javobdan keyin ⏭ tugmasini bosing!\n\n"
        "💡 Taklif: 🧑‍💻@Rustamov_v1",
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(message.from_user.id),
    )
    await add_message(message.from_user.id, msg.message_id)


@dp.message(F.text == "📊 Test natijam")
async def test_result(message: Message):
    await clean_chat(message.from_user.id, message.chat.id)
    user_id = message.from_user.id
    results = users.get(user_id, {}).get("results")

    if not results:
        msg = await message.answer(
            "📊 <b>TEST NATIJALARI</b>\n"
            "━━━━━━━━━━━━━━━━━━━\n\n"
            "❌ Hali natija yo'q.\n\n"
            "📎 Yangi test boshlang!",
            parse_mode="HTML",
            reply_markup=main_menu_keyboard(user_id),
        )
        await add_message(user_id, msg.message_id)
        return

    text = "📊 <b>TEST NATIJALARI (so'nggi 5)</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
    for i, r in enumerate(results[-5:], 1):
        text += (
            f"<b>{i}.</b> {r['date']}\n"
            f"   📝 {r['total']} ta | ✅ {r['score']} ta\n"
            f"   📈 {r['percentage']:.1f}% | {r['grade']}\n\n"
        )
    text += "<i>📎 Yangi test uchun fayl yuboring</i>"
    msg = await message.answer(text, parse_mode="HTML",
                               reply_markup=main_menu_keyboard(user_id))
    await add_message(user_id, msg.message_id)


@dp.message(F.text == "🔁 Qayta boshlash")
async def restart_list(message: Message):
    user_id = message.from_user.id
    docs = users.get(user_id, {}).get("uploaded_docs", [])
    if not docs:
        msg = await message.answer(
            "❌ Oldingi fayl topilmadi. Avval fayl yuboring!",
            reply_markup=main_menu_keyboard(user_id),
        )
        await add_message(user_id, msg.message_id)
        return

    buttons = [
        [InlineKeyboardButton(
            text=f"{i+1}. {d['file_name']} ({len(d['questions'])} ta)",
            callback_data=f"restart_doc_{i}",
        )]
        for i, d in enumerate(docs)
    ]
    buttons.append([InlineKeyboardButton(text="❌ Bekor", callback_data="cancel_test")])
    await clean_chat(user_id, message.chat.id)
    msg = await message.answer(
        "🔁 <b>OLDINGI FAYLLAR</b>\n━━━━━━━━━━━━━━━━━━━\n\nBirini tanlang:",
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
    )
    await add_message(user_id, msg.message_id)


@dp.callback_query(F.data.startswith("restart_doc_"))
async def restart_doc(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    docs = users.get(user_id, {}).get("uploaded_docs", [])
    try:
        idx = int(callback.data.rsplit("_", 1)[-1])
        doc = docs[idx]
    except (ValueError, IndexError):
        await callback.answer("❌ Noto'g'ri tanlov!", show_alert=True)
        return

    users[user_id].update({
        "questions": doc["questions"],
        "total_questions": len(doc["questions"]),
        "file_name": doc["file_name"],
        "selected_doc_index": idx,
    })
    total = len(doc["questions"])
    await state.set_state(TestStates.setting_count)
    txt = (
        f"📝 <b>TEST SONINI TANLANG</b>\n"
        f"━━━━━━━━━━━━━━━━━━━\n\n"
        f"📚 Mavjud savollar: <b>{total} ta</b>\n\n"
        f"<i>Variantni tanlang</i> 👇"
    )
    try:
        await callback.message.edit_text(txt, parse_mode="HTML",
                                         reply_markup=get_count_keyboard(total))
    except Exception:
        msg = await callback.message.answer(txt, parse_mode="HTML",
                                            reply_markup=get_count_keyboard(total))
        await add_message(user_id, msg.message_id)
    await callback.answer()


# ─────────────────────── HANDLER: FAYL QABUL QILISH ─────────────────────────

SUPPORTED_EXT = (".docx", ".doc", ".txt")


@dp.message(F.document)
async def handle_document(message: Message, state: FSMContext):
    await clean_chat(message.from_user.id, message.chat.id)
    doc = message.document
    fname = (doc.file_name or "fayl").lower()

    # Kengaytmani aniqla
    ext = next((e for e in SUPPORTED_EXT if fname.endswith(e)), None)
    if ext is None:
        msg = await message.answer(
            "❌ <b>Qo'llab-quvvatlanmaydigan fayl!</b>\n\n"
            "✅ Qabul qilinadi:\n"
            "• <code>.docx</code> — Word\n"
            "• <code>.doc</code>  — Eski Word\n"
            "• <code>.txt</code>  — Matnli fayl",
            parse_mode="HTML",
            reply_markup=main_menu_keyboard(message.from_user.id),
        )
        await add_message(message.from_user.id, msg.message_id)
        return

    loading = await message.answer(
        "⏳ <b>Fayl yuklanmoqda...</b>\n━━━━━━━━━━━━━━━━━━━\n🔄 Tahlil qilinmoqda...",
        parse_mode="HTML",
    )
    await add_message(message.from_user.id, loading.message_id)

    try:
        cleanup_old_files()
        tg_file = await bot.get_file(doc.file_id)
        downloaded = await bot.download_file(tg_file.file_path)

        save_path = os.path.join("temp", f"test_{message.from_user.id}_{int(time.time())}{ext}")
        with open(save_path, "wb") as f:
            f.write(downloaded.read())

        parse_path = save_path

        # DOC → DOCX konversiya
        if ext == ".doc":
            converted = os.path.splitext(save_path)[0] + ".docx"
            parse_path = convert_doc_to_docx(save_path, converted)
            ext = ".docx"

        # Savollarni parse qilish
        if ext == ".txt":
            questions = parse_txt(parse_path)
            file_icon = "📝"
            fmt_hint = (
                "📝 <b>TXT format namunasi:</b>\n"
                "<code>? Savol matni ?\n"
                "+ To'g'ri javob\n"
                "- Noto'g'ri 1\n"
                "- Noto'g'ri 2\n"
                "- Noto'g'ri 3</code>\n\n"
                "<i>Yoki A/B/C/D format ham qabul qilinadi.</i>"
            )
        else:
            questions = parse_docx(parse_path)
            file_icon = "📘"
            fmt_hint = (
                "📘 <b>DOCX format:</b>\n"
                "• 5 ustunli jadval bo'lishi kerak\n"
                "• 1-ustun: Savol, 2–5: Variantlar"
            )

        await clean_chat(message.from_user.id, message.chat.id)

        if not questions:
            msg = await message.answer(
                f"❌ <b>Test topilmadi!</b>\n\n"
                f"Fayl formatini tekshiring:\n\n{fmt_hint}",
                parse_mode="HTML",
                reply_markup=main_menu_keyboard(message.from_user.id),
            )
            await add_message(message.from_user.id, msg.message_id)
            return

        # Ma'lumotlarni saqlash
        existing = users.get(message.from_user.id, {})
        uploaded_docs = existing.get("uploaded_docs", [])
        uploaded_docs.append({
            "file_name": doc.file_name,
            "file_path": parse_path,
            "questions": questions,
            "uploaded_at": datetime.now().strftime("%d.%m.%Y %H:%M"),
        })
        existing.update({
            "questions": questions,
            "total_questions": len(questions),
            "file_name": doc.file_name,
            "uploaded_file": parse_path,
            "uploaded_docs": uploaded_docs,
        })
        users[message.from_user.id] = existing

        msg = await message.answer(
            f"✅ <b>Fayl muvaffaqiyatli yuklandi!</b>\n"
            f"━━━━━━━━━━━━━━━━━━━\n\n"
            f"📚 Jami savollar: <b>{len(questions)} ta</b>\n"
            f"{file_icon} Fayl: <code>{doc.file_name}</code>\n\n"
            f"<i>Test sonini tanlang...</i>",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="📝 Test sonini tanlash", callback_data="set_count")],
                [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="cancel_doc")],
            ]),
        )
        await add_message(message.from_user.id, msg.message_id)

    except Exception as e:
        await clean_chat(message.from_user.id, message.chat.id)
        msg = await message.answer(
            f"❌ <b>Xatolik yuz berdi:</b>\n<code>{e}</code>",
            parse_mode="HTML",
            reply_markup=main_menu_keyboard(message.from_user.id),
        )
        await add_message(message.from_user.id, msg.message_id)


# ─────────────────────── HANDLER: TEST SOZLASH ────────────────────────────────

@dp.callback_query(F.data == "cancel_doc")
async def cancel_doc(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await clean_chat(callback.from_user.id, callback.message.chat.id)
    try:
        await callback.message.edit_text("❌ Yuklash bekor qilindi")
    except Exception:
        pass
    clear_user_test_session(callback.from_user.id)
    msg = await callback.message.answer(
        "📎 Yangi test uchun fayl yuboring",
        reply_markup=main_menu_keyboard(callback.from_user.id),
    )
    await add_message(callback.from_user.id, msg.message_id)
    await callback.answer()


@dp.callback_query(F.data == "set_count")
async def set_count(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    if user_id not in users:
        await callback.answer("❌ Avval fayl yuklang!", show_alert=True)
        return
    total = users[user_id]["total_questions"]
    await state.set_state(TestStates.setting_count)
    txt = (
        f"📝 <b>TEST SONINI TANLANG</b>\n"
        f"━━━━━━━━━━━━━━━━━━━\n\n"
        f"📚 Mavjud savollar: <b>{total} ta</b>\n\n"
        f"<i>Variantni tanlang</i> 👇"
    )
    try:
        await callback.message.edit_text(txt, parse_mode="HTML",
                                         reply_markup=get_count_keyboard(total))
    except Exception:
        msg = await callback.message.answer(txt, parse_mode="HTML",
                                            reply_markup=get_count_keyboard(total))
        await add_message(user_id, msg.message_id)
    await callback.answer()


@dp.callback_query(F.data.startswith("count_"))
async def select_count(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    if user_id not in users:
        await callback.answer("❌ Xatolik!", show_alert=True)
        return
    count = int(callback.data.split("_")[1])
    await safe_delete_message(callback.message.chat.id, callback.message.message_id)
    await start_test(callback.message, user_id, count, state)
    await callback.answer(f"✅ {count} ta test boshlandi!")


@dp.callback_query(F.data == "custom_count")
async def custom_count_prompt(callback: CallbackQuery):
    try:
        await callback.message.edit_text(
            "✍️ <b>TEST SONINI KIRITING</b>\n"
            "━━━━━━━━━━━━━━━━━━━\n\n"
            "Nechta savol bo'lishini raqam bilan yozing:\n"
            "<i>Masalan: 15, 25, 30...</i>\n\n"
            "Bekor qilish: /cancel",
            parse_mode="HTML",
        )
    except Exception:
        pass
    await callback.answer()


@dp.message(TestStates.setting_count)
async def process_custom_count(message: Message, state: FSMContext):
    user_id = message.from_user.id
    if message.text and message.text.startswith("/"):
        return
    await clean_chat(user_id, message.chat.id)
    if user_id not in users:
        msg = await message.answer("❌ Avval fayl yuklang!",
                                   reply_markup=main_menu_keyboard(user_id))
        await add_message(user_id, msg.message_id)
        await state.clear()
        return
    try:
        count = int(message.text)
        total = users[user_id]["total_questions"]
        if count < 1 or count > total:
            msg = await message.answer(
                f"❌ 1 dan {total} gacha raqam kiriting",
                reply_markup=main_menu_keyboard(user_id),
            )
            await add_message(user_id, msg.message_id)
            return
        await start_test(message, user_id, count, state)
    except (ValueError, TypeError):
        msg = await message.answer("❌ Faqat raqam kiriting! (Masalan: 10)",
                                   reply_markup=main_menu_keyboard(user_id))
        await add_message(user_id, msg.message_id)


# ─────────────────────── TEST JARAYONI ───────────────────────────────────────

async def start_test(message, user_id, count, state: FSMContext):
    await clean_chat(user_id, message.chat.id)

    pool = copy.deepcopy(users[user_id]["questions"])
    random.shuffle(pool)
    selected = pool[:count]
    for q in selected:
        random.shuffle(q["options"])

    users[user_id].update({
        "selected_questions": selected,
        "total_test": count,
        "current_index": 0,
        "score": 0,
        "answers": [],
        "poll_ids": [],
        "waiting_for_skip": False,
        "current_answer_recorded": False,
        "current_poll_message_id": None,
    })
    await state.set_state(TestStates.testing)

    msg = await message.answer(
        f"🚀 <b>TEST BOSHLANDI!</b>\n"
        f"━━━━━━━━━━━━━━━━━━━\n\n"
        f"📝 Jami: <b>{count} ta savol</b>\n"
        f"⏱ Vaqt chegarasi yo'q\n"
        f"🎲 Savollar aralash holda\n\n"
        f"⚠️ <b>Muhim:</b> Javobdan keyin ⏭ tugmasini bosing!\n\n"
        f"<i>Omad!</i> 🍀",
        parse_mode="HTML",
    )
    await add_message(user_id, msg.message_id)
    await asyncio.sleep(2)
    await clean_chat(user_id, message.chat.id)
    await send_poll_question(message.chat.id, user_id)


def _norm(text, limit=100):
    t = str(text).strip()
    return (t[:limit - 1] + "…") if len(t) > limit else t


async def send_poll_question(chat_id, user_id):
    data = users[user_id]
    idx = data["current_index"]
    qd = data["selected_questions"][idx]

    opts = [_norm(o) for o in qd["options"]]
    ans = _norm(qd["answer"])

    try:
        correct_id = opts.index(ans)
    except ValueError:
        correct_id = 0
        opts[0] = ans

    q_text = str(qd["question"]).strip()
    if len(q_text) > 300:
        q_text = q_text[:297] + "..."

    if data.get("current_poll_message_id"):
        await safe_delete_message(chat_id, data["current_poll_message_id"])
    await clean_chat(user_id, chat_id)

    poll_msg = await bot.send_poll(
        chat_id=chat_id,
        question=f"📝 {idx + 1}/{data['total_test']}\n\n{q_text}",
        options=opts,
        type="quiz",
        correct_option_id=correct_id,
        explanation=f"✅ To'g'ri javob: {ans}",
        is_anonymous=False,
        reply_markup=poll_keyboard(),
    )

    data.update({
        "poll_ids": data["poll_ids"] + [poll_msg.poll.id],
        "current_poll_id": poll_msg.poll.id,
        "current_question_index": idx,
        "current_poll_message_id": poll_msg.message_id,
        "waiting_for_skip": True,
        "current_answer_recorded": False,
    })


@dp.poll_answer()
async def on_poll_answer(poll_answer):
    user_id = poll_answer.user.id
    data = users.get(user_id)
    if not data or "selected_questions" not in data:
        return
    if poll_answer.poll_id != data.get("current_poll_id"):
        return
    if data.get("current_answer_recorded") or not poll_answer.option_ids:
        return

    cidx = data.get("current_question_index", 0)
    if cidx >= len(data["selected_questions"]):
        return

    qd = data["selected_questions"][cidx]
    chosen = qd["options"][poll_answer.option_ids[0]]
    correct = chosen == qd["answer"]

    data["answers"].append({
        "question": qd["question"],
        "user_answer": chosen,
        "correct_answer": qd["answer"],
        "is_correct": correct,
    })
    if correct:
        data["score"] += 1
    data["current_answer_recorded"] = True


@dp.callback_query(F.data == "skip_question")
async def skip_question(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    data = users.get(user_id)
    if not data:
        await callback.answer("❌ Test topilmadi!", show_alert=True)
        return
    if not data.get("waiting_for_skip"):
        await callback.answer("⏳ Avval javob bering!", show_alert=True)
        return

    cidx = data.get("current_question_index", 0)
    if not data.get("current_answer_recorded"):
        qd = data["selected_questions"][cidx]
        data["answers"].append({
            "question": qd["question"],
            "user_answer": "Javob berilmadi",
            "correct_answer": qd["answer"],
            "is_correct": False,
        })

    data["waiting_for_skip"] = False
    data["current_answer_recorded"] = False
    data["current_index"] += 1

    if data["current_index"] >= data["total_test"]:
        await clean_chat(user_id, callback.message.chat.id)
        await safe_delete_message(callback.message.chat.id, callback.message.message_id)
        await show_results(callback.message.chat.id, user_id)
        await state.clear()
        await callback.answer("✅ Test yakunlandi!")
        return

    await callback.answer("⏭ Keyingi savol...")
    await clean_chat(user_id, callback.message.chat.id)
    await send_poll_question(callback.message.chat.id, user_id)


@dp.callback_query(F.data == "stop_test")
async def stop_test(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    data = users.get(user_id)
    if not data:
        await callback.answer("❌ Test topilmadi!", show_alert=True)
        return

    cidx = data.get("current_question_index", 0)
    if len(data["answers"]) <= cidx < data["total_test"]:
        qd = data["selected_questions"][cidx]
        data["answers"].append({
            "question": qd["question"],
            "user_answer": "Test yakunlandi",
            "correct_answer": qd["answer"],
            "is_correct": False,
        })

    await clean_chat(user_id, callback.message.chat.id)
    await safe_delete_message(callback.message.chat.id, callback.message.message_id)
    await show_results(callback.message.chat.id, user_id, stopped=True)
    await state.clear()
    await callback.answer("🔴 Test yakunlandi")


# ─────────────────────── NATIJALAR ────────────────────────────────────────────

async def show_results(chat_id, user_id, stopped=False):
    data = users[user_id]
    score = data["score"]
    total = data["total_test"]
    answered = len(data["answers"])
    percentage = (score / total * 100) if total else 0

    if percentage >= 90:
        grade, emoji = "🏆 A'lo", "🌟"
    elif percentage >= 75:
        grade, emoji = "🎉 Yaxshi", "👏"
    elif percentage >= 60:
        grade, emoji = "👍 Qoniqarli", "💪"
    else:
        grade, emoji = "📚 O'qish kerak", "📖"

    text = (
        f"{emoji} <b>TEST NATIJASI</b>\n"
        f"━━━━━━━━━━━━━━━━━━━\n\n"
        f"📊 <b>Statistika:</b>\n"
        f"• Jami savollar: <b>{total} ta</b>\n"
        f"• Javob berilgan: <b>{answered} ta</b>\n"
        f"• O'tkazib yuborilgan: <b>{max(total - answered, 0)} ta</b>\n\n"
        f"✅ To'g'ri: <b>{score} ta</b>\n"
        f"❌ Noto'g'ri: <b>{max(total - score, 0)} ta</b>\n"
        f"📈 Foiz: <b>{percentage:.1f}%</b>\n"
        f"🏆 Baho: <b>{grade}</b>\n\n"
        f"<i>{'⚠️ Test vaqtidan oldin yakunlandi' if stopped else '🎊 Test muvaffaqiyatli yakunlandi!'}</i>"
    )

    data.setdefault("results", []).append({
        "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
        "total": total, "score": score,
        "percentage": percentage, "grade": grade,
    })

    await clean_chat(user_id, chat_id)
    msg = await bot.send_message(
        chat_id, text, parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [
                InlineKeyboardButton(text="🔄 Qayta urinish", callback_data="retry_poll"),
                InlineKeyboardButton(text="📋 Batafsil", callback_data="poll_details"),
            ],
            [
                InlineKeyboardButton(text="📊 Barcha natijalar", callback_data="all_results"),
                InlineKeyboardButton(text="🏠 Menyu", callback_data="main_menu"),
            ],
        ]),
    )
    await add_message(user_id, msg.message_id)


@dp.callback_query(F.data == "retry_poll")
async def retry_poll(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    data = users.get(user_id)
    if not data:
        await callback.answer("❌ Ma'lumot topilmadi", show_alert=True)
        return

    selected = copy.deepcopy(data["selected_questions"])
    random.shuffle(selected)
    for q in selected:
        random.shuffle(q["options"])

    data.update({
        "selected_questions": selected,
        "current_index": 0,
        "score": 0,
        "answers": [],
        "poll_ids": [],
        "waiting_for_skip": False,
        "current_answer_recorded": False,
    })
    await state.set_state(TestStates.testing)
    await clean_chat(user_id, callback.message.chat.id)
    await safe_delete_message(callback.message.chat.id, callback.message.message_id)

    msg = await callback.message.answer(
        f"🔄 <b>TEST QAYTA BOSHLANDI!</b>\n"
        f"━━━━━━━━━━━━━━━━━━━\n\n"
        f"📝 Jami: <b>{data['total_test']} ta savol</b>\n"
        f"🎲 Savollar qayta aralashtirildi\n\n"
        f"⚠️ Javobdan keyin ⏭ tugmasini bosing!\n\n"
        f"<i>Omad!</i> 🍀",
        parse_mode="HTML",
    )
    await add_message(user_id, msg.message_id)
    await asyncio.sleep(2)
    await clean_chat(user_id, callback.message.chat.id)
    await send_poll_question(callback.message.chat.id, user_id)
    await callback.answer("✅ Qayta boshlandi")


@dp.callback_query(F.data == "poll_details")
async def poll_details(callback: CallbackQuery):
    user_id = callback.from_user.id
    answers = users.get(user_id, {}).get("answers", [])
    if not answers:
        await callback.answer("❌ Javoblar yo'q", show_alert=True)
        return

    text = "📋 <b>BATAFSIL NATIJALAR</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
    for i, a in enumerate(answers, 1):
        icon = "✅" if a["is_correct"] else "❌"
        q = a["question"][:60] + ("..." if len(a["question"]) > 60 else "")
        text += f"<b>{i}.</b> {q}\n   {icon} {a['user_answer']}\n"
        if not a["is_correct"]:
            text += f"   ✅ To'g'ri: {a['correct_answer']}\n"
        text += "\n"
        if len(text) > 3500:
            msg = await callback.message.answer(text, parse_mode="HTML")
            await add_message(user_id, msg.message_id)
            text = ""

    if text:
        msg = await callback.message.answer(text, parse_mode="HTML")
        await add_message(user_id, msg.message_id)
    await callback.answer()


@dp.callback_query(F.data == "all_results")
async def all_results(callback: CallbackQuery):
    user_id = callback.from_user.id
    results = users.get(user_id, {}).get("results", [])
    if not results:
        await callback.answer("❌ Natijalar yo'q", show_alert=True)
        return

    await clean_chat(user_id, callback.message.chat.id)
    await safe_delete_message(callback.message.chat.id, callback.message.message_id)

    text = "📊 <b>BARCHA NATIJALAR</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
    for i, r in enumerate(results[-10:], 1):
        text += (
            f"<b>{i}.</b> {r['date']}\n"
            f"   📝 {r['total']} ta | ✅ {r['score']} ta\n"
            f"   📈 {r['percentage']:.1f}% | {r['grade']}\n\n"
        )
    msg = await callback.message.answer(text, parse_mode="HTML",
                                        reply_markup=main_menu_keyboard(user_id))
    await add_message(user_id, msg.message_id)
    await callback.answer()


@dp.callback_query(F.data == "main_menu")
async def back_to_menu(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id
    clear_user_test_session(user_id)
    await safe_delete_message(chat_id, callback.message.message_id)
    await clean_chat(user_id, chat_id)
    msg = await bot.send_message(
        chat_id,
        "🏠 <b>BOSH MENYU</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
        "📎 Yangi test uchun fayl yuboring\n\n"
        "<i>Menu tugmalaridan foydalaning</i> 👇",
        parse_mode="HTML",
        reply_markup=main_menu_keyboard(user_id),
    )
    await add_message(user_id, msg.message_id)
    await callback.answer("🏠 Bosh menyu")


@dp.callback_query(F.data == "cancel_test")
async def cancel_test(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    user_id = callback.from_user.id
    chat_id = callback.message.chat.id
    clear_user_test_session(user_id)
    await clean_chat(user_id, chat_id)
    await safe_delete_message(chat_id, callback.message.message_id)
    msg = await bot.send_message(
        chat_id,
        "❌ Bekor qilindi\n\n📎 Yangi test uchun fayl yuboring",
        reply_markup=main_menu_keyboard(user_id),
    )
    await add_message(user_id, msg.message_id)
    await callback.answer()


# ─────────────────────────── MAIN ────────────────────────────────────────────

async def main():
    print("🚀 Bot ishga tushdi...")
    cleanup_old_files()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
