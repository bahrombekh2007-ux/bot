import asyncio
import copy
import os
import random
import re
import time
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart, Command
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
from datetime import datetime, timezone, timedelta

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

TOKEN = os.getenv("BOT_TOKEN", "8964353995:AAFgRTesY5nYBku_fyuFMNLQ2VW_hPMzOrg")
bot = Bot(token=TOKEN)
dp = Dispatcher()

os.makedirs("temp", exist_ok=True)

UZBEKISTAN_TZ = timezone(timedelta(hours=5))


def get_uz_time():
    return datetime.now(UZBEKISTAN_TZ)


def cleanup_old_files(directory="temp", max_age_days=3):
    cutoff = time.time() - max_age_days * 86400
    for name in os.listdir(directory):
        path = os.path.join(directory, name)
        try:
            if os.path.isfile(path) and os.path.getmtime(path) < cutoff:
                os.remove(path)
        except Exception:
            pass


# ─────────────────────────── PARSERLAR ───────────────────────────────────────

def _read_txt(file_path):
    """TXT faylni turli kodlashlarda o'qiydi."""
    for enc in ("utf-8-sig", "utf-8", "cp1251", "windows-1251", "latin-1"):
        try:
            with open(file_path, "r", encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, LookupError):
            pass
    return ""


def _parse_hash_format(lines):
    """
    Ikki xil formatni qabul qiladi:

    FORMAT A — # va savol bir qatorda:
        # Savol matni
        + To'g'ri javob
        - Noto'g'ri 1

    FORMAT B — # yolg'iz qatorda, savol keyingi qatorda:
        #
        Savol matni
        +
        To'g'ri javob
        -
        Noto'g'ri 1
    """
    questions = []
    current_q = None
    correct = None
    opts = []
    # holat: 'idle' | 'need_q' | 'need_correct' | 'need_wrong'
    state = 'idle'

    def flush():
        nonlocal current_q, correct, opts, state
        if current_q and correct and len(opts) >= 2:
            all_opts = opts[:4]
            while len(all_opts) < 4:
                all_opts.append(f"Variant {len(all_opts) + 1}")
            if correct not in all_opts:
                all_opts.insert(0, correct)
                all_opts = all_opts[:4]
            questions.append({
                "question": current_q,
                "options": all_opts,
                "answer": correct,
            })
        current_q, correct, opts, state = None, None, [], 'idle'

    for raw in lines:
        line = raw.strip()
        if not line:
            continue

        # # belgisi — savol boshlanishi
        if line.startswith("#") or line.startswith("?"):
            flush()
            q = line[1:].strip()
            q = re.sub(r"\s*\?$", "", q).strip()
            if q:
                # FORMAT A: # Savol matni (bir qatorda)
                current_q = q
                state = 'idle'
            else:
                # FORMAT B: # yolg'iz, savol keyingi qatorda
                state = 'need_q'

        elif state == 'need_q':
            # Keyingi qator — savol matni
            current_q = re.sub(r"\s*\?$", "", line).strip()
            state = 'idle'

        elif line.startswith("+"):
            ans = line[1:].strip()
            if not ans:
                # FORMAT B: + yolg'iz, javob keyingi qatorda
                state = 'need_correct'
            else:
                if current_q is not None:
                    correct = ans
                    if ans not in opts:
                        opts.append(ans)
                    state = 'idle'

        elif state == 'need_correct':
            correct = line
            if line not in opts:
                opts.append(line)
            state = 'idle'

        elif line.startswith("-"):
            ans = line[1:].strip()
            if not ans:
                # FORMAT B: - yolg'iz, javob keyingi qatorda
                state = 'need_wrong'
            else:
                if current_q is not None and ans not in opts:
                    opts.append(ans)

        elif state == 'need_wrong':
            if current_q is not None and line not in opts:
                opts.append(line)
            state = 'idle'

        # Agar hech bir belgi bilan boshlanmasa va state 'idle' bo'lsa — o'tkazib yuborish
        elif state == 'idle':
            pass

        # Qolgan holatlar — o'tkazib yuborish
        else:
            pass

    flush()
    return questions


def _parse_numbered_abcd(lines):
    """
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
                questions.append({
                    "question": current_q,
                    "options": options,
                    "answer": correct_ans,
                })
        current_q, opts_dict, correct_letter = None, {}, None

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
            r"^(?:Javob|To'g'ri\s*javob|Answer|Ans|Togri\s*javob|Javobi)[:\s]*([A-Da-d])",
            line, re.IGNORECASE
        )
        if m:
            correct_letter = m.group(1).upper()

    flush()
    return questions


def _parse_pipe(lines):
    """Savol|Javob A|Javob B|Javob C|Javob D"""
    questions = []
    for line in lines:
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 5 and parts[0]:
            questions.append({
                "question": parts[0],
                "options": parts[1:5],
                "answer": parts[1],
            })
    return questions


def _parse_qa_format(lines):
    """
    Savol: Savol matni?
    Javob: To'g'ri javob
    """
    questions = []
    current_q = None
    current_ans = None

    def flush():
        nonlocal current_q, current_ans
        if current_q and current_ans:
            questions.append({
                "question": current_q,
                "options": [current_ans, "Variant B", "Variant C", "Variant D"],
                "answer": current_ans,
            })
        current_q, current_ans = None, None

    for line in lines:
        line = line.strip()
        if not line:
            flush()
            continue

        if re.match(r"^(savol|question)\s*:", line, re.IGNORECASE):
            flush()
            current_q = re.split(r":\s*", line, maxsplit=1, flags=re.IGNORECASE)[-1].strip()
        elif re.match(r"^(javob|answer)\s*:", line, re.IGNORECASE):
            current_ans = re.split(r":\s*", line, maxsplit=1, flags=re.IGNORECASE)[-1].strip()

    flush()
    return questions


def parse_txt(file_path):
    """TXT faylni avtomatik format aniqlab parse qiladi."""
    content = _read_txt(file_path)
    if not content:
        return []

    lines = content.splitlines()
    non_empty = [l.strip() for l in lines if l.strip()]
    if not non_empty:
        return []

    # Format aniqlanishi
    has_hash = any(l.startswith("#") for l in non_empty)
    has_question_mark = any(l.startswith("?") for l in non_empty)
    has_plus = any(l.startswith("+") for l in non_empty)
    has_pipe = any("|" in l and l.count("|") >= 4 for l in non_empty)
    has_numbered = any(re.match(r"^\d+[.)]\s+", l) for l in non_empty)
    has_abcd = any(re.match(r"^[A-Da-d][.)]\s+", l) for l in non_empty)
    has_javob_key = any(re.match(r"^(javob|answer)\s*:", l, re.IGNORECASE) for l in non_empty)
    has_savol_key = any(re.match(r"^(savol|question)\s*:", l, re.IGNORECASE) for l in non_empty)

    # 1. # yoki ? + - format (eng keng tarqalgan)
    if (has_hash or has_question_mark) and has_plus:
        result = _parse_hash_format(lines)
        if result:
            return result

    # 2. Raqamli A/B/C/D format
    if has_numbered and has_abcd:
        result = _parse_numbered_abcd(lines)
        if result:
            return result

    # 3. Pipe format
    if has_pipe:
        result = _parse_pipe(non_empty)
        if result:
            return result

    # 4. Savol: Javob: format
    if has_savol_key or has_javob_key:
        result = _parse_qa_format(lines)
        if result:
            return result

    # 5. Agar hech narsa topilmasa — barcha parserlarni sinab ko'r
    for parser in [_parse_hash_format, _parse_numbered_abcd, _parse_pipe, _parse_qa_format]:
        result = parser(lines if parser != _parse_pipe else non_empty)
        if result:
            return result

    return []


def parse_docx(file_path):
    """DOCX fayldan savollarni o'qiydi (jadval va paragraf)."""
    from docx import Document
    doc = Document(file_path)
    questions = []

    def add_q(question, options, answer):
        if question and len(options) >= 2 and answer in options:
            opts = options[:4]
            while len(opts) < 4:
                opts.append(f"Variant {len(opts) + 1}")
            questions.append({"question": question, "options": opts, "answer": answer})

    for table in doc.tables:
        rows = []
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            cells = list(dict.fromkeys(cells))  # duplikat celllarni olib tashla
            cells = [c for c in cells if c]
            if cells:
                rows.append(cells)

        if not rows:
            continue

        # 1 ustunli jadval (har 5 qator = 1 savol)
        if all(len(r) == 1 for r in rows):
            flat = [r[0] for r in rows]
            for i in range(0, len(flat), 5):
                b = flat[i:i + 5]
                if len(b) == 5:
                    add_q(b[0], b[1:], b[1])
            continue

        # Ko'p ustunli jadval
        for r in rows:
            if len(r) >= 5:
                add_q(r[0], r[1:5], r[1])
            elif len(r) == 3:
                # Savol | To'g'ri | Noto'g'ri format
                add_q(r[0], [r[1], r[2], "Variant C", "Variant D"], r[1])

    # Jadvaldan savol topilmasa — paragraflardan o'qi
    if not questions:
        lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        # Avval # format sinab ko'r
        result = _parse_hash_format(lines)
        if result:
            return result
        # Keyin raqamli format
        result = _parse_numbered_abcd(lines)
        if result:
            return result
        # Oddiy 5-qatorli blok
        for i in range(0, len(lines), 5):
            b = lines[i:i + 5]
            if len(b) == 5:
                add_q(b[0], b[1:], b[1])

    return questions


def parse_xlsx(file_path):
    """Excel fayldan savollarni o'qiydi."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        questions = []
        for sheet in wb.worksheets:
            rows = []
            for row in sheet.iter_rows(values_only=True):
                cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
                if cells:
                    rows.append(cells)
            for r in rows:
                if len(r) >= 5:
                    q, *opts = r[:5]
                    questions.append({"question": q, "options": opts, "answer": opts[0]})
                elif len(r) == 5:
                    questions.append({"question": r[0], "options": r[1:5], "answer": r[1]})
        return questions
    except ImportError:
        return []
    except Exception:
        return []


def parse_pdf(file_path):
    """PDF fayldan savollarni o'qiydi."""
    try:
        import pdfplumber
        lines = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines.extend(text.splitlines())
        lines = [l.strip() for l in lines if l.strip()]
        if not lines:
            return []
        # Barcha formatlarni sinab ko'r
        for parser in [_parse_hash_format, _parse_numbered_abcd, _parse_pipe, _parse_qa_format]:
            result = parser(lines)
            if result:
                return result
        return []
    except ImportError:
        return []
    except Exception:
        return []


# ─────────────────────────── FSM & GLOBAL STATE ──────────────────────────────

class TestStates(StatesGroup):
    setting_count = State()
    testing = State()


users: dict = {}
user_messages: dict = {}

SUPPORTED_EXT = (".docx", ".doc", ".txt", ".xlsx", ".pdf")


def convert_doc_to_docx(doc_path, docx_path):
    """DOC faylni DOCX ga aylantiradi (LibreOffice orqali)."""
    import shutil, subprocess
    for soffice in ["soffice", "libreoffice",
                    r"C:\Program Files\LibreOffice\program\soffice.exe",
                    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
        found = (shutil.which(soffice) or
                 (os.path.isabs(soffice) and os.path.exists(soffice)))
        if found:
            out_dir = os.path.dirname(docx_path)
            subprocess.run(
                [soffice, "--headless", "--convert-to", "docx",
                 "--outdir", out_dir, doc_path],
                check=True, capture_output=True, timeout=60,
            )
            auto_out = os.path.join(
                out_dir,
                os.path.splitext(os.path.basename(doc_path))[0] + ".docx"
            )
            if os.path.exists(auto_out):
                if auto_out != docx_path:
                    os.rename(auto_out, docx_path)
                return docx_path
    raise RuntimeError(
        "LibreOffice topilmadi! .doc faylni ochish uchun o'rnating:\n"
        "Linux: sudo apt install libreoffice\n"
        "Windows: https://www.libreoffice.org"
    )


def clear_user_test_session(user_id):
    if user_id not in users:
        return
    for key in [
        "questions", "total_questions", "file_name", "uploaded_file",
        "selected_questions", "total_test", "current_index", "score",
        "answers", "poll_ids", "waiting_for_skip", "current_poll_message_id",
        "current_poll_id", "current_question_index", "current_answer_recorded",
    ]:
        users[user_id].pop(key, None)


async def safe_delete(chat_id, message_id):
    try:
        await bot.delete_message(chat_id, message_id)
    except Exception:
        pass


async def clean_chat(user_id, chat_id, keep_last=0):
    msgs = user_messages.get(user_id, [])
    to_del = msgs[:-keep_last] if keep_last else msgs
    for mid in to_del:
        await safe_delete(chat_id, mid)
    user_messages[user_id] = msgs[-keep_last:] if keep_last else []


async def add_msg(user_id, message_id):
    user_messages.setdefault(user_id, [])
    if message_id not in user_messages[user_id]:
        user_messages[user_id].append(message_id)


# ──────────────────────────── KLAVIATURALAR ───────────────────────────────────

def main_kb(user_id=None):
    row1 = [KeyboardButton(text="📊 Test natijam"), KeyboardButton(text="🆘 Yordam")]
    row2 = [KeyboardButton(text="📄 Yangi test"), KeyboardButton(text="📁 Fayllarim")]
    row3 = [KeyboardButton(text="⭐ Statistika"), KeyboardButton(text="⚙️ Sozlamalar")]
    if user_id and users.get(user_id, {}).get("uploaded_docs"):
        row2.append(KeyboardButton(text="🔁 Qayta boshlash"))
    return ReplyKeyboardMarkup(keyboard=[row1, row2, row3], resize_keyboard=True)


def count_kb(total):
    opts = [n for n in (5, 10, 15, 20, 25, 30, 40, 50) if total >= n]
    if total not in opts:
        opts.append(total)
    opts.sort()
    buttons, row = [], []
    for i, c in enumerate(opts):
        label = f"📚 Hammasi ({c})" if c == total else f"📝 {c} ta"
        row.append(InlineKeyboardButton(text=label, callback_data=f"count_{c}"))
        if len(row) == 2 or i == len(opts) - 1:
            buttons.append(row)
            row = []
    buttons.append([InlineKeyboardButton(text="✍️ O'zim kiritaman", callback_data="custom_count")])
    buttons.append([InlineKeyboardButton(text="🎲 Tasodifiy", callback_data="random_count")])
    buttons.append([InlineKeyboardButton(text="❌ Bekor qilish", callback_data="cancel_test")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def poll_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⏭ Keyingi savol", callback_data="skip_question")],
        [
            InlineKeyboardButton(text="💡 Javobni ko'rsat", callback_data="show_answer"),
            InlineKeyboardButton(text="🔴 Yakunlash", callback_data="stop_test"),
        ],
    ])


# ─────────────────────────── HANDLER: START ───────────────────────────────────

@dp.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    uid = message.from_user.id
    await clean_chat(uid, message.chat.id)

    if uid not in users:
        users[uid] = {
            "first_visit": get_uz_time().strftime("%d.%m.%Y %H:%M"),
            "total_tests": 0, "total_questions": 0, "total_correct": 0,
            "results": [], "uploaded_docs": [],
        }

    msg = await message.answer(
        "🎯 <b>TEST MASTER BOT</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        f"👋 Assalomu alaykum, <b>{message.from_user.first_name}</b>!\n\n"
        "✨ <b>Imkoniyatlar:</b>\n"
        "• 🎲 Savollar va variantlar aralash\n"
        "• 📝 Test sonini o'zingiz belgilaysiz\n"
        "• 📊 Batafsil statistika\n"
        "• 📁 Fayllar tarixi\n\n"
        "📎 Fayl yuboring\n\n"
        ✅ Qabul qilinadi:
        • .txt — Matnli fayl
        • .docx — Word
        • .xlsx — Excel
        • .pdf — PDF
        "💬 Murojat uchun? @Rustamov_v1",
        parse_mode="HTML",
        reply_markup=main_kb(uid),
    )
    await add_msg(uid, msg.message_id)


# ─────────────────────────── HANDLER: MENYU ───────────────────────────────────

@dp.message(F.text == "📄 Yangi test")
async def new_test(message: Message, state: FSMContext):
    await state.clear()
    uid = message.from_user.id
    await clean_chat(uid, message.chat.id)
    msg = await message.answer(
        "📄 <b>YANGI TEST</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        "📎 Fayl yuboring\n\n"
        "📝 <b>TXT Format 1</b> (<code>#</code> belgisi):\n"
        "<code># Savol matni\n+ To'g'ri javob\n- Noto'g'ri 1\n- Noto'g'ri 2\n- Noto'g'ri 3</code>\n\n"
        "📝 <b>TXT Format 2</b> (A/B/C/D):\n"
        "<code>1. Savol matni\nA) To'g'ri\nB) Noto'g'ri 1\nC) Noto'g'ri 2\nD) Noto'g'ri 3\nJavob: A</code>\n\n"
        "📝 <b>TXT Format 3</b> (Pipe):\n"
        "<code>Savol|A variant|B variant|C variant|D variant</code>\n\n"
        "📘 <b>DOCX:</b> Jadval (1-ustun: Savol, 2–5: Variantlar)\n"
        "📊 <b>XLSX:</b> Jadval (A: Savol, B–E: Variantlar)\n"
        "📄 <b>PDF:</b> Yuqoridagi TXT formatlardan biri",
        parse_mode="HTML",
        reply_markup=main_kb(uid),
    )
    await add_msg(uid, msg.message_id)


@dp.message(F.text == "🆘 Yordam")
async def cmd_help(message: Message):
    uid = message.from_user.id
    await clean_chat(uid, message.chat.id)
    msg = await message.answer(
        "🆘 <b>YORDAM</b>\n"
        "━━━━━━━━━━━━━━━━━━━\n\n"
        "📌 <b>Foydalanish tartibi:</b>\n"
        "1️⃣ Fayl yuboring (.txt / .docx / .xlsx / .pdf)\n"
        "2️⃣ Test sonini tanlang\n"
        "3️⃣ Javob bering → ⏭ tugmasini bosing\n\n"
        "─────────────────────\n"
        "📝 <b>TXT Format (<code>#</code> belgisi):</b>\n"
        "<code># Savol\n+ To'g'ri\n- Noto'g'ri 1\n- Noto'g'ri 2\n- Noto'g'ri 3</code>\n\n"
        "📝 <b>TXT Format (A/B/C/D):</b>\n"
        "<code>1. Savol\nA) To'g'ri\nB) Noto'g'ri 1\nC) Noto'g'ri 2\nD) Noto'g'ri 3\nJavob: A</code>\n\n"
        "📘 <b>DOCX:</b> 5 ustunli jadval\n"
        "📊 <b>XLSX:</b> 5 ustunli jadval\n\n"
        "⚠️ Har javobdan keyin ⏭ tugmasini bosing!\n\n"
        "💬 Muammo? @Rustamov_v1",
        parse_mode="HTML",
        reply_markup=main_kb(uid),
    )
    await add_msg(uid, msg.message_id)


@dp.message(F.text == "📊 Test natijam")
async def test_result(message: Message):
    uid = message.from_user.id
    await clean_chat(uid, message.chat.id)
    results = users.get(uid, {}).get("results", [])

    if not results:
        msg = await message.answer(
            "📊 <b>TEST NATIJALARI</b>\n"
            "━━━━━━━━━━━━━━━━━━━\n\n"
            "❌ Hali natija yo'q.\n\n📎 Yangi test boshlang!",
            parse_mode="HTML", reply_markup=main_kb(uid),
        )
        await add_msg(uid, msg.message_id)
        return

    text = "📊 <b>TEST NATIJALARI (so'nggi 10)</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
    for i, r in enumerate(results[-10:], 1):
        text += (
            f"<b>{i}.</b> {r['date']}\n"
            f"   📝 {r['total']} ta | ✅ {r['score']} ta\n"
            f"   📈 {r['percentage']:.1f}% | {r['grade']}\n\n"
        )
    msg = await message.answer(text, parse_mode="HTML", reply_markup=main_kb(uid))
    await add_msg(uid, msg.message_id)


@dp.message(F.text == "⭐ Statistika")
async def my_stats(message: Message):
    uid = message.from_user.id
    d = users.get(uid, {})
    results = d.get("results", [])
    avg = sum(r["percentage"] for r in results) / len(results) if results else 0
    best = max(results, key=lambda x: x["percentage"]) if results else None

    text = (
        "⭐ <b>UMUMIY STATISTIKA</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
        f"📅 Birinchi foydalanish: <b>{d.get('first_visit', 'Noma\'lum')}</b>\n\n"
        f"📊 Jami testlar: <b>{d.get('total_tests', 0)} ta</b>\n"
        f"📝 Jami savollar: <b>{d.get('total_questions', 0)} ta</b>\n"
        f"✅ To'g'ri javoblar: <b>{d.get('total_correct', 0)} ta</b>\n"
        f"📈 O'rtacha natija: <b>{avg:.1f}%</b>\n"
    )
    if best:
        text += f"\n🏆 <b>Eng yaxshi:</b> {best['percentage']:.1f}% — {best['date']}\n"

    msg = await message.answer(text, parse_mode="HTML", reply_markup=main_kb(uid))
    await add_msg(uid, msg.message_id)


@dp.message(F.text == "📁 Fayllarim")
async def my_files(message: Message):
    uid = message.from_user.id
    docs = users.get(uid, {}).get("uploaded_docs", [])

    if not docs:
        msg = await message.answer(
            "📁 <b>Fayllarim</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
            "❌ Hali hech qanday fayl yuklanmagan.\n\n📎 Fayl yuboring!",
            parse_mode="HTML", reply_markup=main_kb(uid),
        )
        await add_msg(uid, msg.message_id)
        return

    text = "📁 <b>Mening fayllarim</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
    for i, doc in enumerate(docs[-10:], 1):
        text += f"<b>{i}.</b> {doc['file_name']} — <b>{len(doc['questions'])} savol</b>\n   📅 {doc['uploaded_at']}\n\n"

    buttons = [
        [InlineKeyboardButton(
            text=f"📄 {d['file_name'][:35]}",
            callback_data=f"selfile_{i}"
        )]
        for i, d in enumerate(docs[-5:])
    ]
    buttons.append([InlineKeyboardButton(text="🏠 Menyu", callback_data="main_menu")])

    msg = await message.answer(text, parse_mode="HTML",
                               reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await add_msg(uid, msg.message_id)


@dp.message(F.text == "⚙️ Sozlamalar")
async def settings(message: Message):
    uid = message.from_user.id
    msg = await message.answer(
        "⚙️ <b>SOZLAMALAR</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
        "🔄 Savollar aralashtirish: ✅\n"
        "🔄 Variantlar aralashtirish: ✅\n\n"
        "<i>Tez orada qo'shimcha sozlamalar...</i>",
        parse_mode="HTML", reply_markup=main_kb(uid)
    )
    await add_msg(uid, msg.message_id)


@dp.message(F.text == "🔁 Qayta boshlash")
async def restart_list(message: Message):
    uid = message.from_user.id
    docs = users.get(uid, {}).get("uploaded_docs", [])
    if not docs:
        msg = await message.answer("❌ Oldingi fayl topilmadi. Avval fayl yuboring!",
                                   reply_markup=main_kb(uid))
        await add_msg(uid, msg.message_id)
        return

    buttons = [
        [InlineKeyboardButton(
            text=f"{i + 1}. {d['file_name']} ({len(d['questions'])} ta)",
            callback_data=f"restart_{i}",
        )]
        for i, d in enumerate(docs[-5:])
    ]
    buttons.append([InlineKeyboardButton(text="❌ Bekor", callback_data="cancel_test")])
    await clean_chat(uid, message.chat.id)
    msg = await message.answer(
        "🔁 <b>OLDINGI FAYLLAR</b>\n━━━━━━━━━━━━━━━━━━━\n\nBirini tanlang:",
        parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
    )
    await add_msg(uid, msg.message_id)


# ─────────────────────── CALLBACK: FAYL TANLASH ──────────────────────────────

@dp.callback_query(F.data.startswith("selfile_"))
async def select_file_cb(callback: CallbackQuery, state: FSMContext):
    uid = callback.from_user.id
    docs = users.get(uid, {}).get("uploaded_docs", [])
    try:
        idx = int(callback.data.split("_")[1])
        doc = docs[idx]
    except (ValueError, IndexError):
        await callback.answer("❌ Xatolik!", show_alert=True)
        return
    users[uid].update({
        "questions": doc["questions"],
        "total_questions": len(doc["questions"]),
        "file_name": doc["file_name"],
    })
    await state.set_state(TestStates.setting_count)
    await _show_count_selection(callback, uid, len(doc["questions"]))


@dp.callback_query(F.data.startswith("restart_"))
async def restart_doc_cb(callback: CallbackQuery, state: FSMContext):
    uid = callback.from_user.id
    docs = users.get(uid, {}).get("uploaded_docs", [])
    try:
        idx = int(callback.data.split("_")[1])
        doc = docs[idx]
    except (ValueError, IndexError):
        await callback.answer("❌ Noto'g'ri tanlov!", show_alert=True)
        return
    users[uid].update({
        "questions": doc["questions"],
        "total_questions": len(doc["questions"]),
        "file_name": doc["file_name"],
    })
    await state.set_state(TestStates.setting_count)
    await _show_count_selection(callback, uid, len(doc["questions"]))


async def _show_count_selection(callback, uid, total):
    txt = (
        f"📝 <b>TEST SONINI TANLANG</b>\n"
        f"━━━━━━━━━━━━━━━━━━━\n\n"
        f"📚 Mavjud savollar: <b>{total} ta</b>\n\n"
        f"<i>Variantni tanlang</i> 👇"
    )
    try:
        await callback.message.edit_text(txt, parse_mode="HTML", reply_markup=count_kb(total))
    except Exception:
        msg = await callback.message.answer(txt, parse_mode="HTML", reply_markup=count_kb(total))
        await add_msg(uid, msg.message_id)
    await callback.answer()


# ─────────────────────── HANDLER: FAYL QABUL QILISH ─────────────────────────

@dp.message(F.document)
async def handle_document(message: Message, state: FSMContext):
    uid = message.from_user.id
    await clean_chat(uid, message.chat.id)
    doc = message.document
    fname = (doc.file_name or "fayl").lower()

    ext = next((e for e in SUPPORTED_EXT if fname.endswith(e)), None)
    if ext is None:
        msg = await message.answer(
            "❌ <b>Qo'llab-quvvatlanmaydigan fayl!</b>\n\n"
            "✅ Qabul qilinadi:\n"
            "• <code>.txt</code>  — Matnli fayl\n"
            "• <code>.docx</code> — Word\n"
            "• <code>.doc</code>  — Eski Word\n"
            "• <code>.xlsx</code> — Excel\n"
            "• <code>.pdf</code>  — PDF",
            parse_mode="HTML", reply_markup=main_kb(uid),
        )
        await add_msg(uid, msg.message_id)
        return

    loading = await message.answer(
        "⏳ <b>Fayl yuklanmoqda...</b>\n━━━━━━━━━━━━━━━━━━━\n🔄 Tahlil qilinmoqda...",
        parse_mode="HTML",
    )
    await add_msg(uid, loading.message_id)

    try:
        cleanup_old_files()
        tg_file = await bot.get_file(doc.file_id)
        downloaded = await bot.download_file(tg_file.file_path)
        save_path = os.path.join("temp", f"u{uid}_{int(time.time())}{ext}")
        with open(save_path, "wb") as f:
            f.write(downloaded.read())

        # .doc → .docx konversiya
        if ext == ".doc":
            converted = save_path.replace(".doc", "_conv.docx")
            save_path = convert_doc_to_docx(save_path, converted)
            ext = ".docx"

        # Parse qilish
        if ext == ".txt":
            questions = parse_txt(save_path)
            fmt_name = "TXT"
        elif ext == ".docx":
            questions = parse_docx(save_path)
            fmt_name = "DOCX"
        elif ext == ".xlsx":
            questions = parse_xlsx(save_path)
            fmt_name = "XLSX"
        elif ext == ".pdf":
            questions = parse_pdf(save_path)
            fmt_name = "PDF"
        else:
            questions = []
            fmt_name = "?"

        await clean_chat(uid, message.chat.id)

        if not questions:
            hint = (
                "📌 <b>TXT uchun to'g'ri format:</b>\n"
                "<code># Savol matni\n+ To'g'ri javob\n- Noto'g'ri 1\n- Noto'g'ri 2\n- Noto'g'ri 3</code>\n\n"
                "yoki\n\n"
                "<code>1. Savol\nA) To'g'ri\nB) Noto'g'ri 1\nC) Noto'g'ri 2\nD) Noto'g'ri 3\nJavob: A</code>"
                if ext == ".txt" else
                "📘 <b>DOCX/XLSX uchun:</b> 5 ustunli jadval kerak\n"
                "(1-ustun: Savol, 2-5-ustunlar: Variantlar)"
            )
            msg = await message.answer(
                f"❌ <b>Savol topilmadi!</b>\n\n"
                f"<i>{fmt_name} fayldan savol o'qib bo'lmadi.</i>\n\n{hint}",
                parse_mode="HTML", reply_markup=main_kb(uid),
            )
            await add_msg(uid, msg.message_id)
            return

        # Saqlash
        existing = users.setdefault(uid, {
            "first_visit": get_uz_time().strftime("%d.%m.%Y %H:%M"),
            "total_tests": 0, "total_questions": 0, "total_correct": 0,
            "results": [], "uploaded_docs": [],
        })
        uploaded_docs = existing.get("uploaded_docs", [])
        uploaded_docs.append({
            "file_name": doc.file_name,
            "file_path": save_path,
            "questions": questions,
            "uploaded_at": get_uz_time().strftime("%d.%m.%Y %H:%M"),
        })
        existing.update({
            "questions": questions,
            "total_questions": len(questions),
            "file_name": doc.file_name,
            "uploaded_file": save_path,
            "uploaded_docs": uploaded_docs,
        })
        await state.set_state(TestStates.setting_count)

        msg = await message.answer(
            f"✅ <b>Fayl muvaffaqiyatli yuklandi!</b>\n"
            f"━━━━━━━━━━━━━━━━━━━\n\n"
            f"📋 Format: <b>{fmt_name}</b>\n"
            f"📚 Jami savollar: <b>{len(questions)} ta</b>\n"
            f"📄 Fayl: <code>{doc.file_name}</code>\n\n"
            f"<i>Test sonini tanlang 👇</i>",
            parse_mode="HTML",
            reply_markup=count_kb(len(questions)),
        )
        await add_msg(uid, msg.message_id)

    except Exception as e:
        await clean_chat(uid, message.chat.id)
        msg = await message.answer(
            f"❌ <b>Xatolik yuz berdi:</b>\n<code>{str(e)[:300]}</code>",
            parse_mode="HTML", reply_markup=main_kb(uid),
        )
        await add_msg(uid, msg.message_id)


# ─────────────────────── HANDLER: TEST SOZLASH ────────────────────────────────

@dp.callback_query(F.data.startswith("count_"))
async def select_count(callback: CallbackQuery, state: FSMContext):
    uid = callback.from_user.id
    if uid not in users:
        await callback.answer("❌ Avval fayl yuklang!", show_alert=True)
        return
    count = int(callback.data.split("_")[1])
    await safe_delete(callback.message.chat.id, callback.message.message_id)
    await start_test(callback.message, uid, count, state)
    await callback.answer(f"✅ {count} ta test boshlandi!")


@dp.callback_query(F.data == "random_count")
async def random_count(callback: CallbackQuery, state: FSMContext):
    uid = callback.from_user.id
    if uid not in users:
        await callback.answer("❌ Avval fayl yuklang!", show_alert=True)
        return
    total = users[uid]["total_questions"]
    count = random.randint(5, min(50, total))
    await safe_delete(callback.message.chat.id, callback.message.message_id)
    await start_test(callback.message, uid, count, state)
    await callback.answer(f"🎲 {count} ta test boshlandi!")


@dp.callback_query(F.data == "custom_count")
async def custom_count_prompt(callback: CallbackQuery, state: FSMContext):
    await state.set_state(TestStates.setting_count)
    try:
        await callback.message.edit_text(
            "✍️ <b>TEST SONINI KIRITING</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
            "Nechta savol bo'lishini raqam bilan yozing:\n"
            "<i>Masalan: 15, 25, 30...</i>",
            parse_mode="HTML",
        )
    except Exception:
        pass
    await callback.answer()


@dp.message(TestStates.setting_count)
async def process_custom_count(message: Message, state: FSMContext):
    uid = message.from_user.id
    if message.text and message.text.startswith("/"):
        return
    await clean_chat(uid, message.chat.id)
    if uid not in users or "total_questions" not in users[uid]:
        msg = await message.answer("❌ Avval fayl yuklang!", reply_markup=main_kb(uid))
        await add_msg(uid, msg.message_id)
        await state.clear()
        return
    try:
        count = int(message.text.strip())
        total = users[uid]["total_questions"]
        if count < 1 or count > total:
            msg = await message.answer(f"❌ 1 dan {total} gacha raqam kiriting",
                                       reply_markup=main_kb(uid))
            await add_msg(uid, msg.message_id)
            return
        await start_test(message, uid, count, state)
    except (ValueError, TypeError):
        msg = await message.answer("❌ Faqat raqam kiriting! (Masalan: 10)",
                                   reply_markup=main_kb(uid))
        await add_msg(uid, msg.message_id)


# ─────────────────────── TEST JARAYONI ───────────────────────────────────────

async def start_test(message, uid, count, state: FSMContext):
    await clean_chat(uid, message.chat.id)
    pool = copy.deepcopy(users[uid]["questions"])
    random.shuffle(pool)
    selected = pool[:count]
    for q in selected:
        random.shuffle(q["options"])

    users[uid].update({
        "selected_questions": selected,
        "total_test": count,
        "current_index": 0,
        "score": 0,
        "answers": [],
        "poll_ids": [],
        "waiting_for_skip": False,
        "current_answer_recorded": False,
        "current_poll_message_id": None,
        "test_start_time": get_uz_time(),
    })
    await state.set_state(TestStates.testing)

    msg = await message.answer(
        f"🚀 <b>TEST BOSHLANDI!</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
        f"📝 Jami: <b>{count} ta savol</b>\n"
        f"🎲 Savollar aralash holda\n\n"
        f"⚠️ <b>Muhim:</b> Javobdan keyin ⏭ tugmasini bosing!\n\n"
        f"<i>Omad! 🍀</i>",
        parse_mode="HTML",
    )
    await add_msg(uid, msg.message_id)
    await asyncio.sleep(2)
    await clean_chat(uid, message.chat.id)
    await send_poll(message.chat.id, uid)


def _norm(text, limit=100):
    t = str(text).strip()
    return (t[:limit - 1] + "…") if len(t) > limit else t


async def send_poll(chat_id, uid):
    data = users[uid]
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
        await safe_delete(chat_id, data["current_poll_message_id"])
    await clean_chat(uid, chat_id)

    poll_msg = await bot.send_poll(
        chat_id=chat_id,
        question=f"📝 {idx + 1}/{data['total_test']}\n\n{q_text}",
        options=opts,
        type="quiz",
        correct_option_id=correct_id,
        explanation=f"✅ To'g'ri: {ans}",
        is_anonymous=False,
        reply_markup=poll_kb(),
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
    uid = poll_answer.user.id
    data = users.get(uid)
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


@dp.callback_query(F.data == "show_answer")
async def show_answer(callback: CallbackQuery):
    uid = callback.from_user.id
    data = users.get(uid)
    if not data or "selected_questions" not in data:
        await callback.answer("❌ Test topilmadi!", show_alert=True)
        return
    cidx = data.get("current_question_index", 0)
    qd = data["selected_questions"][cidx]
    await callback.answer(f"✅ To'g'ri javob: {qd['answer']}", show_alert=True)


@dp.callback_query(F.data == "skip_question")
async def skip_question(callback: CallbackQuery, state: FSMContext):
    uid = callback.from_user.id
    data = users.get(uid)
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
        await clean_chat(uid, callback.message.chat.id)
        await safe_delete(callback.message.chat.id, callback.message.message_id)
        await show_results(callback.message.chat.id, uid)
        await state.clear()
        await callback.answer("✅ Test yakunlandi!")
        return

    await callback.answer("⏭ Keyingi...")
    await clean_chat(uid, callback.message.chat.id)
    await send_poll(callback.message.chat.id, uid)


@dp.callback_query(F.data == "stop_test")
async def stop_test(callback: CallbackQuery, state: FSMContext):
    uid = callback.from_user.id
    data = users.get(uid)
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

    await clean_chat(uid, callback.message.chat.id)
    await safe_delete(callback.message.chat.id, callback.message.message_id)
    await show_results(callback.message.chat.id, uid, stopped=True)
    await state.clear()
    await callback.answer("🔴 Test yakunlandi")


# ─────────────────────── NATIJALAR ────────────────────────────────────────────

async def show_results(chat_id, uid, stopped=False):
    data = users[uid]
    score = data["score"]
    total = data["total_test"]
    answered = len(data["answers"])
    pct = (score / total * 100) if total else 0

    if pct >= 90:
        grade, emoji = "A'lo", "🏆"
    elif pct >= 75:
        grade, emoji = "Yaxshi", "🎉"
    elif pct >= 60:
        grade, emoji = "Qoniqarli", "👍"
    else:
        grade, emoji = "O'qish kerak", "📚"

    start = data.get("test_start_time")
    time_str = ""
    if start:
        diff = get_uz_time() - start
        m, s = diff.seconds // 60, diff.seconds % 60
        time_str = f"\n⏱ Vaqt: {m} min {s} sek"

    text = (
        f"{emoji} <b>TEST NATIJASI</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
        f"📊 <b>Statistika:</b>\n"
        f"• Jami: <b>{total} ta</b>\n"
        f"• Javob berilgan: <b>{answered} ta</b>\n"
        f"• O'tkazilgan: <b>{max(total - answered, 0)} ta</b>{time_str}\n\n"
        f"✅ To'g'ri: <b>{score} ta</b>\n"
        f"❌ Noto'g'ri: <b>{max(total - score, 0)} ta</b>\n"
        f"📈 Foiz: <b>{pct:.1f}%</b>\n"
        f"🏆 Baho: <b>{grade}</b>\n\n"
        f"<i>{'⚠️ Test vaqtidan oldin yakunlandi' if stopped else '🎊 Test muvaffaqiyatli yakunlandi!'}</i>"
    )

    data.setdefault("total_tests", 0)
    data.setdefault("total_questions", 0)
    data.setdefault("total_correct", 0)
    data["total_tests"] += 1
    data["total_questions"] += total
    data["total_correct"] += score
    data.setdefault("results", []).append({
        "date": get_uz_time().strftime("%d.%m.%Y %H:%M"),
        "total": total, "score": score, "percentage": pct, "grade": grade,
    })

    await clean_chat(uid, chat_id)
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
    await add_msg(uid, msg.message_id)


@dp.callback_query(F.data == "retry_poll")
async def retry_poll(callback: CallbackQuery, state: FSMContext):
    uid = callback.from_user.id
    data = users.get(uid)
    if not data or "selected_questions" not in data:
        await callback.answer("❌ Ma'lumot topilmadi", show_alert=True)
        return

    selected = copy.deepcopy(data["selected_questions"])
    random.shuffle(selected)
    for q in selected:
        random.shuffle(q["options"])

    data.update({
        "selected_questions": selected,
        "current_index": 0, "score": 0, "answers": [],
        "poll_ids": [], "waiting_for_skip": False,
        "current_answer_recorded": False,
        "test_start_time": get_uz_time(),
    })
    await state.set_state(TestStates.testing)
    await clean_chat(uid, callback.message.chat.id)
    await safe_delete(callback.message.chat.id, callback.message.message_id)

    msg = await callback.message.answer(
        f"🔄 <b>TEST QAYTA BOSHLANDI!</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
        f"📝 {data['total_test']} ta savol | 🎲 Aralashtirildi\n\n<i>Omad! 🍀</i>",
        parse_mode="HTML",
    )
    await add_msg(uid, msg.message_id)
    await asyncio.sleep(2)
    await clean_chat(uid, callback.message.chat.id)
    await send_poll(callback.message.chat.id, uid)
    await callback.answer("✅ Boshlandi")


@dp.callback_query(F.data == "poll_details")
async def poll_details(callback: CallbackQuery):
    uid = callback.from_user.id
    answers = users.get(uid, {}).get("answers", [])
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
            await add_msg(uid, msg.message_id)
            text = ""
    if text:
        msg = await callback.message.answer(text, parse_mode="HTML")
        await add_msg(uid, msg.message_id)
    await callback.answer()


@dp.callback_query(F.data == "all_results")
async def all_results(callback: CallbackQuery):
    uid = callback.from_user.id
    results = users.get(uid, {}).get("results", [])
    if not results:
        await callback.answer("❌ Natijalar yo'q", show_alert=True)
        return

    await clean_chat(uid, callback.message.chat.id)
    await safe_delete(callback.message.chat.id, callback.message.message_id)

    text = "📊 <b>BARCHA NATIJALAR</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
    for i, r in enumerate(results[-10:], 1):
        text += (
            f"<b>{i}.</b> {r['date']}\n"
            f"   📝 {r['total']} ta | ✅ {r['score']} ta\n"
            f"   📈 {r['percentage']:.1f}% | {r['grade']}\n\n"
        )
    msg = await callback.message.answer(text, parse_mode="HTML", reply_markup=main_kb(uid))
    await add_msg(uid, msg.message_id)
    await callback.answer()


@dp.callback_query(F.data == "main_menu")
async def back_to_menu(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    uid = callback.from_user.id
    clear_user_test_session(uid)
    await safe_delete(callback.message.chat.id, callback.message.message_id)
    await clean_chat(uid, callback.message.chat.id)
    msg = await bot.send_message(
        callback.message.chat.id,
        "🏠 <b>BOSH MENYU</b>\n━━━━━━━━━━━━━━━━━━━\n\n"
        "📎 Yangi test uchun fayl yuboring 👇",
        parse_mode="HTML",
        reply_markup=main_kb(uid),
    )
    await add_msg(uid, msg.message_id)
    await callback.answer("🏠 Bosh menyu")


@dp.callback_query(F.data == "cancel_test")
async def cancel_test(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    uid = callback.from_user.id
    clear_user_test_session(uid)
    await clean_chat(uid, callback.message.chat.id)
    await safe_delete(callback.message.chat.id, callback.message.message_id)
    msg = await bot.send_message(
        callback.message.chat.id,
        "❌ Bekor qilindi\n\n📎 Yangi test uchun fayl yuboring",
        reply_markup=main_kb(uid),
    )
    await add_msg(uid, msg.message_id)
    await callback.answer()


# ─────────────────────────── MAIN ────────────────────────────────────────────

async def main():
    print("🚀 Bot ishga tushdi...")
    print(f"📍 Vaqt: UTC+5 (O'zbekiston)")
    cleanup_old_files()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
