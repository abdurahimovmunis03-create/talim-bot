#!/usr/bin/env python3
"""
O'zbekiston OTM Ta'lim Boti
Mustaqil ish va Kurs ishini OTM standartlarida tayyorlab beruvchi Telegram bot
"""

import os
import io
import re
import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler,
    filters, ContextTypes
)
import anthropic
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ── Sozlamalar ────────────────────────────────────────────────────
BOT_TOKEN = os.environ["BOT_TOKEN"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
claude = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

# Suhbat bosqichlari
LANG, WORK_TYPE, UNIVERSITY, FACULTY, DIRECTION, SUBJECT, STUDENT_NAME, COURSE, STUDY_TYPE, TOPIC = range(10)

# ── Matnlar (3 tilda) ─────────────────────────────────────────────
T = {
    "uz": {
        "welcome": (
            "👋 *Assalomu alaykum!*\n\n"
            "Men O'zbekiston OTM standartlarida yozma ishlar tayyorlab beruvchi botman:\n\n"
            "📝 *Mustaqil ish* — 15–20 sahifa\n"
            "📚 *Kurs ishi* — 25–35 sahifa\n\n"
            "Tayyor bo'lgan ish Word (.docx) formatida yuboriladi.\n\n"
            "🌐 *Tilni tanlang:*"
        ),
        "work_type": "📋 Qanday ish kerak?",
        "mustaqil": "📝 Mustaqil ish",
        "kurs": "📚 Kurs ishi",
        "university": "🏛 *Universitetingiz* nomini to'liq kiriting:\n_(masalan: Mirzo Ulug'bek nomidagi O'zbekiston Milliy Universiteti)_",
        "faculty": "🏫 *Fakultetingiz* nomini kiriting:\n_(masalan: Ijtimoiy fanlar fakulteti)_",
        "direction": "📖 *Yo'nalishingizni* kiriting:\n_(masalan: Amaliy psixologiya yo'nalishi)_",
        "subject": "📚 *Fan nomini* kiriting:\n_(masalan: Sotsiologiya)_",
        "student": "👤 *Ism va familiyangizni* kiriting:\n_(masalan: Abdullayev Jasur Baxtiyorovich)_",
        "course": "🎓 *Nechanchi kursda* o'qiysiz?\n_(1 dan 4 gacha raqam kiriting)_",
        "study_type": "📅 O'qish shaklini tanlang:",
        "kunduzgi": "☀️ Kunduzgi",
        "sirtqi": "🌙 Sirtqi",
        "kechki": "🌆 Kechki",
        "topic": (
            "✏️ *Ish mavzusini* kiriting:\n\n"
            "_Masalan: Ijtimoiy munosabatlar tushunchasi va uning turlari_\n\n"
            "⚡ Mavzu qanchalik aniq bo'lsa, ish shunchalik sifatli bo'ladi!"
        ),
        "generating": (
            "⏳ *Ish tayyorlanmoqda...*\n\n"
            "Bu 1–3 daqiqa olishi mumkin.\n"
            "Iltimos, kuting 🙏"
        ),
        "done": "✅ *Ish tayyor!* Word formatida yuborilmoqda...",
        "error": "❌ Xatolik yuz berdi. /start bilan qaytadan urinib ko'ring.",
        "invalid_course": "⚠️ Iltimos, 1 dan 4 gacha raqam kiriting.",
        "restart": "\n\n🔄 Yangi ish uchun: /start",
    },
    "ru": {
        "welcome": (
            "👋 *Здравствуйте!*\n\n"
            "Я бот для написания учебных работ по стандартам узбекских вузов:\n\n"
            "📝 *Самостоятельная работа* — 15–20 страниц\n"
            "📚 *Курсовая работа* — 25–35 страниц\n\n"
            "Готовая работа отправляется в формате Word (.docx).\n\n"
            "🌐 *Выберите язык:*"
        ),
        "work_type": "📋 Какая работа нужна?",
        "mustaqil": "📝 Самостоятельная работа",
        "kurs": "📚 Курсовая работа",
        "university": "🏛 Введите *полное название университета*:\n_(например: Национальный университет Узбекистана им. Мирзо Улугбека)_",
        "faculty": "🏫 Введите *название факультета*:\n_(например: Факультет социальных наук)_",
        "direction": "📖 Введите *направление обучения*:\n_(например: Прикладная психология)_",
        "subject": "📚 Введите *название предмета*:\n_(например: Социология)_",
        "student": "👤 Введите *ваше имя и фамилию*:\n_(например: Иванов Иван Иванович)_",
        "course": "🎓 На *каком курсе* вы учитесь?\n_(введите цифру от 1 до 4)_",
        "study_type": "📅 Выберите форму обучения:",
        "kunduzgi": "☀️ Дневное",
        "sirtqi": "🌙 Заочное",
        "kechki": "🌆 Вечернее",
        "topic": (
            "✏️ Введите *тему работы*:\n\n"
            "_Например: Понятие социальных отношений и их виды_\n\n"
            "⚡ Чем точнее тема, тем качественнее работа!"
        ),
        "generating": (
            "⏳ *Работа готовится...*\n\n"
            "Это займёт 1–3 минуты.\n"
            "Пожалуйста, подождите 🙏"
        ),
        "done": "✅ *Работа готова!* Отправляю в формате Word...",
        "error": "❌ Произошла ошибка. Попробуйте /start снова.",
        "invalid_course": "⚠️ Введите цифру от 1 до 4.",
        "restart": "\n\n🔄 Для новой работы: /start",
    },
    "en": {
        "welcome": (
            "👋 *Hello!*\n\n"
            "I'm a bot for writing academic papers per Uzbekistan HEI standards:\n\n"
            "📝 *Independent work* — 15–20 pages\n"
            "📚 *Course work* — 25–35 pages\n\n"
            "The finished work is sent in Word (.docx) format.\n\n"
            "🌐 *Choose language:*"
        ),
        "work_type": "📋 What type of work do you need?",
        "mustaqil": "📝 Independent work",
        "kurs": "📚 Course work",
        "university": "🏛 Enter the *full university name*:\n_(e.g. National University of Uzbekistan named after Mirzo Ulugbek)_",
        "faculty": "🏫 Enter your *faculty name*:\n_(e.g. Faculty of Social Sciences)_",
        "direction": "📖 Enter your *study direction*:\n_(e.g. Applied Psychology)_",
        "subject": "📚 Enter the *subject name*:\n_(e.g. Sociology)_",
        "student": "👤 Enter your *full name*:\n_(e.g. John Smith)_",
        "course": "🎓 Which *year* are you in?\n_(enter a number from 1 to 4)_",
        "study_type": "📅 Select study type:",
        "kunduzgi": "☀️ Full-time",
        "sirtqi": "🌙 Part-time",
        "kechki": "🌆 Evening",
        "topic": (
            "✏️ Enter the *work topic*:\n\n"
            "_E.g. The concept of social relations and their types_\n\n"
            "⚡ The more specific the topic, the better the work!"
        ),
        "generating": (
            "⏳ *Preparing the work...*\n\n"
            "This may take 1–3 minutes.\n"
            "Please wait 🙏"
        ),
        "done": "✅ *Work is ready!* Sending in Word format...",
        "error": "❌ An error occurred. Try /start again.",
        "invalid_course": "⚠️ Please enter a number from 1 to 4.",
        "restart": "\n\n🔄 For a new work: /start",
    },
}

STUDY_TYPES = {
    "uz": {"kunduzgi": "kunduzgi", "sirtqi": "sirtqi", "kechki": "kechki"},
    "ru": {"kunduzgi": "дневное", "sirtqi": "заочное", "kechki": "вечернее"},
    "en": {"kunduzgi": "full-time", "sirtqi": "part-time", "kechki": "evening"},
}

# ── Klaviaturalar ─────────────────────────────────────────────────
def lang_kb():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("🇺🇿 O'zbek", callback_data="lang_uz"),
        InlineKeyboardButton("🇷🇺 Русский", callback_data="lang_ru"),
        InlineKeyboardButton("🇬🇧 English", callback_data="lang_en"),
    ]])

def work_kb(lang):
    t = T[lang]
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(t["mustaqil"], callback_data="work_mustaqil")],
        [InlineKeyboardButton(t["kurs"], callback_data="work_kurs")],
    ])

def study_kb(lang):
    t = T[lang]
    return InlineKeyboardMarkup([[
        InlineKeyboardButton(t["kunduzgi"], callback_data="study_kunduzgi"),
        InlineKeyboardButton(t["sirtqi"], callback_data="study_sirtqi"),
        InlineKeyboardButton(t["kechki"], callback_data="study_kechki"),
    ]])

# ── Claude prompts ────────────────────────────────────────────────
def prompt_mustaqil(d: dict) -> str:
    lang, topic, subject = d["lang"], d["topic"], d["subject"]
    course, study = d["course"], d["study_type"]

    if lang == "uz":
        return f"""Siz O'zbekiston universiteti {course}-kurs {study} talabasi uchun "{subject}" fanidan "{topic}" mavzusida MUSTAQIL ISH yozmoqdasiz.

TALABLAR:
- O'zbek tilida, ilmiy uslubda
- Har bir bo'lim KAMIDA 700 so'z
- Jami matn 15-20 sahifaga mos bo'lsin
- Faktlar, tahlil, ilmiy asoslar bo'lsin

Quyidagi ANIQ belgilar bilan yozing (belgilarni o'zgartirmang!):

===KIRISH===
Mavzuning dolzarbligi, o'rganilganlik darajasi, maqsad va vazifalari haqida 4-5 paragraf. Kamida 450 so'z.

===BOLIM_1_NOMI===
(Birinchi bo'lim sarlavhasi - mavzuning birinchi jihatiga oid)

===BOLIM_1_MATN===
Birinchi bo'lim to'liq matni. Kamida 750 so'z. Ilmiy tahlil, misollar, nazariya.

===BOLIM_2_NOMI===
(Ikkinchi bo'lim sarlavhasi)

===BOLIM_2_MATN===
Ikkinchi bo'lim to'liq matni. Kamida 750 so'z.

===BOLIM_3_NOMI===
(Uchinchi bo'lim sarlavhasi)

===BOLIM_3_MATN===
Uchinchi bo'lim to'liq matni. Kamida 750 so'z.

===XULOSA===
Asosiy xulosalar, 3-4 paragraf. Kamida 350 so'z.

===ADABIYOTLAR===
Kamida 12 ta manba, raqamlangan. Format:
1. Muallif I.O. Kitob nomi. – Shahar: Nashriyot, Yil.
"""
    elif lang == "ru":
        return f"""Напишите самостоятельную работу для студента {course}-го курса ({study}) по предмету "{subject}" на тему "{topic}".

ТРЕБОВАНИЯ:
- Научный стиль, на русском языке
- Каждый раздел МИНИМУМ 700 слов
- Факты, анализ, научная база

Используйте ТОЧНЫЕ маркеры:

===KIRISH===
Введение: актуальность, степень изученности, цели и задачи. Минимум 450 слов.

===BOLIM_1_NOMI===
(Название первого раздела)

===BOLIM_1_MATN===
Текст первого раздела. Минимум 750 слов.

===BOLIM_2_NOMI===
(Название второго раздела)

===BOLIM_2_MATN===
Текст второго раздела. Минимум 750 слов.

===BOLIM_3_NOMI===
(Название третьего раздела)

===BOLIM_3_MATN===
Текст третьего раздела. Минимум 750 слов.

===XULOSA===
Заключение, 3-4 абзаца. Минимум 350 слов.

===ADABIYOTLAR===
Минимум 12 источников, пронумерованных.
"""
    else:
        return f"""Write an independent academic work for a {course}-year ({study}) student on "{subject}", topic: "{topic}".

REQUIREMENTS:
- Academic style, in English
- Each section MINIMUM 700 words
- Facts, analysis, scientific basis

Use EXACT markers:

===KIRISH===
Introduction: relevance, literature review, goals and objectives. Minimum 450 words.

===BOLIM_1_NOMI===
(First section title)

===BOLIM_1_MATN===
First section full text. Minimum 750 words.

===BOLIM_2_NOMI===
(Second section title)

===BOLIM_2_MATN===
Second section full text. Minimum 750 words.

===BOLIM_3_NOMI===
(Third section title)

===BOLIM_3_MATN===
Third section full text. Minimum 750 words.

===XULOSA===
Conclusion, 3-4 paragraphs. Minimum 350 words.

===ADABIYOTLAR===
Minimum 12 numbered sources.
"""


def prompt_kurs(d: dict) -> str:
    lang, topic, subject = d["lang"], d["topic"], d["subject"]
    course, study = d["course"], d["study_type"]

    if lang == "uz":
        return f"""Siz O'zbekiston universiteti {course}-kurs {study} talabasi uchun "{subject}" fanidan "{topic}" mavzusida KURS ISHI yozmoqdasiz.

TALABLAR:
- O'zbek tilida, ilmiy uslubda
- Har bir § paragraf KAMIDA 700 so'z
- Jami 25-35 sahifaga mos
- Chuqur ilmiy tahlil, faktlar, misollar

Quyidagi ANIQ belgilar bilan yozing:

===KIRISH===
Kirish bo'limi (kamida 600 so'z):
- Mavzuning dolzarbligi (2-3 paragraf)
- Mavzuning nazariy o'rganilganlik darajasi (1-2 paragraf)
- Kurs ishining maqsadi
- Kurs ishining vazifalari (5-6 ta, ro'yxat bilan)
- Kurs ishining ob'ekti va predmeti
- Nazariy va amaliy ahamiyati
- Kurs ishining tuzilishi va hajmi

===I_BOB_NOMI===
(I Bobning to'liq sarlavhasi — nazariy asoslar)

===PAR_1_1_NOMI===
(1.1 § sarlavhasi)

===PAR_1_1_MATN===
1.1 § to'liq matni. Kamida 750 so'z. Ilmiy tahlil.

===PAR_1_2_NOMI===
(1.2 § sarlavhasi)

===PAR_1_2_MATN===
1.2 § to'liq matni. Kamida 750 so'z.

===PAR_1_3_NOMI===
(1.3 § sarlavhasi)

===PAR_1_3_MATN===
1.3 § to'liq matni. Kamida 750 so'z.

===I_BOB_XULOSA===
I Bob bo'yicha xulosalar. Kamida 280 so'z.

===II_BOB_NOMI===
(II Bobning to'liq sarlavhasi — amaliy qism)

===PAR_2_1_NOMI===
(2.1 § sarlavhasi)

===PAR_2_1_MATN===
2.1 § to'liq matni. Kamida 750 so'z.

===PAR_2_2_NOMI===
(2.2 § sarlavhasi)

===PAR_2_2_MATN===
2.2 § to'liq matni. Kamida 750 so'z.

===PAR_2_3_NOMI===
(2.3 § sarlavhasi)

===PAR_2_3_MATN===
2.3 § to'liq matni. Kamida 750 so'z.

===II_BOB_XULOSA===
II Bob bo'yicha xulosalar. Kamida 280 so'z.

===XULOSA===
Umumiy xulosa. Kamida 420 so'z.

===GLOSSARIY===
Kamida 20 ta atama. Format:
**Atama** — ta'rifi (1-2 jumlada aniq izoh)

===ADABIYOTLAR===
Asosiy adabiyotlar (kamida 12 ta):
1. ...

Internet manbalar (kamida 5 ta):
1. ...
"""
    elif lang == "ru":
        return f"""Напишите курсовую работу для студента {course}-го курса ({study}) по предмету "{subject}" на тему "{topic}".

ТРЕБОВАНИЯ:
- Научный стиль, на русском языке
- Каждый параграф МИНИМУМ 700 слов
- Объём 25-35 страниц

Точные маркеры:

===KIRISH===
Введение (минимум 600 слов): актуальность, степень изученности, цель, задачи (5-6 штук), объект и предмет, значимость, структура работы.

===I_BOB_NOMI===
(Название I главы — теоретические основы)

===PAR_1_1_NOMI===
(Название параграфа 1.1)

===PAR_1_1_MATN===
Текст параграфа 1.1. Минимум 750 слов.

===PAR_1_2_NOMI===
(Название параграфа 1.2)

===PAR_1_2_MATN===
Текст параграфа 1.2. Минимум 750 слов.

===PAR_1_3_NOMI===
(Название параграфа 1.3)

===PAR_1_3_MATN===
Текст параграфа 1.3. Минимум 750 слов.

===I_BOB_XULOSA===
Выводы по I главе. Минимум 280 слов.

===II_BOB_NOMI===
(Название II главы — практическая часть)

===PAR_2_1_NOMI===
(Название параграфа 2.1)

===PAR_2_1_MATN===
Текст. Минимум 750 слов.

===PAR_2_2_NOMI===
(Название параграфа 2.2)

===PAR_2_2_MATN===
Текст. Минимум 750 слов.

===PAR_2_3_NOMI===
(Название параграфа 2.3)

===PAR_2_3_MATN===
Текст. Минимум 750 слов.

===II_BOB_XULOSA===
Выводы по II главе. Минимум 280 слов.

===XULOSA===
Общее заключение. Минимум 420 слов.

===GLOSSARIY===
Минимум 20 терминов. Формат:
**Термин** — определение

===ADABIYOTLAR===
Основная литература (минимум 12):
1. ...

Интернет-источники (минимум 5):
1. ...
"""
    else:
        return f"""Write a course work for a {course}-year ({study}) student on "{subject}", topic: "{topic}".

REQUIREMENTS:
- Academic style, in English
- Each paragraph MINIMUM 700 words
- Total 25-35 pages

Exact markers:

===KIRISH===
Introduction (minimum 600 words): relevance, literature review, goal, tasks (5-6 items), object & subject, significance, structure.

===I_BOB_NOMI===
(Chapter I title — theoretical foundations)

===PAR_1_1_NOMI===
(Section 1.1 title)

===PAR_1_1_MATN===
Section 1.1 text. Minimum 750 words.

===PAR_1_2_NOMI===
(Section 1.2 title)

===PAR_1_2_MATN===
Section 1.2 text. Minimum 750 words.

===PAR_1_3_NOMI===
(Section 1.3 title)

===PAR_1_3_MATN===
Section 1.3 text. Minimum 750 words.

===I_BOB_XULOSA===
Chapter I conclusions. Minimum 280 words.

===II_BOB_NOMI===
(Chapter II title — practical part)

===PAR_2_1_NOMI===
(Section 2.1 title)

===PAR_2_1_MATN===
Section 2.1 text. Minimum 750 words.

===PAR_2_2_NOMI===
(Section 2.2 title)

===PAR_2_2_MATN===
Section 2.2 text. Minimum 750 words.

===PAR_2_3_NOMI===
(Section 2.3 title)

===PAR_2_3_MATN===
Section 2.3 text. Minimum 750 words.

===II_BOB_XULOSA===
Chapter II conclusions. Minimum 280 words.

===XULOSA===
General conclusion. Minimum 420 words.

===GLOSSARIY===
Minimum 20 terms. Format:
**Term** — definition

===ADABIYOTLAR===
Main sources (minimum 12):
1. ...

Internet sources (minimum 5):
1. ...
"""


# ── Parsing ───────────────────────────────────────────────────────
def parse_sections(text: str) -> dict:
    parts = re.split(r"===([A-Z0-9_']+)===", text)
    result = {}
    for i in range(1, len(parts), 2):
        key = parts[i].strip()
        val = parts[i + 1].strip() if i + 1 < len(parts) else ""
        result[key] = val
    return result


# ── Word hujjat yaratish ──────────────────────────────────────────
def doc_setup(doc: Document):
    """OTM standart formatlash"""
    sec = doc.sections[0]
    sec.left_margin = Cm(3)
    sec.right_margin = Cm(1.5)
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(14)
    pf = normal.paragraph_format
    pf.line_spacing = Pt(21)
    pf.first_line_indent = Cm(1.25)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


def add_footer_pagenum(doc: Document):
    """Pastda sahifa raqami"""
    for sec in doc.sections:
        footer = sec.footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText"); instr.text = "PAGE"
        fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "end")
        run._r.extend([fld1, instr, fld2])


def h(doc: Document, text: str, level: int = 1):
    """Sarlavha qo'shish"""
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = Pt(21)
    run = p.add_run(text.upper() if level == 1 else text)
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(16 if level == 1 else 14)


def body(doc: Document, text: str):
    """Oddiy matn qo'shish"""
    for chunk in text.split("\n\n"):
        chunk = chunk.strip()
        if not chunk:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(1.25)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = Pt(21)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(chunk)
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)


def ref_line(doc: Document, text: str):
    """Adabiyot qatori (chekinmasiz)"""
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = Pt(21)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text.strip())
    run.font.name = "Times New Roman"
    run.font.size = Pt(14)


def muqova(doc: Document, d: dict):
    """Muqova sahifasi"""
    import datetime
    lang = d["lang"]
    year = datetime.datetime.now().year

    def cp(text, bold=False, size=13, space=0):
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(space)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(19)
        r = p.add_run(text)
        r.bold = bold
        r.font.name = "Times New Roman"
        r.font.size = Pt(size)

    def rp(text, size=13):
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.right_indent = Cm(1.5)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(19)
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(size)

    if lang == "uz":
        cp("O'ZBEKISTON RESPUBLIKASI OLIY VA O'RTA MAXSUS TA'LIM VAZIRLIGI", bold=True, size=12)
    elif lang == "ru":
        cp("МИНИСТЕРСТВО ВЫСШЕГО И СРЕДНЕГО СПЕЦИАЛЬНОГО ОБРАЗОВАНИЯ\nРЕСПУБЛИКИ УЗБЕКИСТАН", bold=True, size=12)
    else:
        cp("MINISTRY OF HIGHER AND SECONDARY SPECIALISED EDUCATION\nOF THE REPUBLIC OF UZBEKISTAN", bold=True, size=12)

    cp("")
    cp(d["university"].upper(), bold=True, size=13, space=6)
    cp("")
    cp(d["faculty"], size=12)
    cp(d["direction"], size=12)

    cp(""); cp(""); cp("")

    if lang == "uz":
        cp(f'"{d["subject"]}"', size=12)
        cp("fanidan", size=12)
    elif lang == "ru":
        cp(f'По предмету: "{d["subject"]}"', size=12)
    else:
        cp(f'Subject: "{d["subject"]}"', size=12)

    cp("")
    cp(f'"{d["topic"].upper()}"', bold=True, size=14, space=4)
    cp("")

    if lang == "uz":
        wname = "MUSTAQIL ISH" if d["work_type"] == "mustaqil" else "KURS ISHI"
    elif lang == "ru":
        wname = "САМОСТОЯТЕЛЬНАЯ РАБОТА" if d["work_type"] == "mustaqil" else "КУРСОВАЯ РАБОТА"
    else:
        wname = "INDEPENDENT WORK" if d["work_type"] == "mustaqil" else "COURSE WORK"

    cp(wname, bold=True, size=16, space=6)

    cp(""); cp(""); cp("")

    if lang == "uz":
        rp(f"{d['course']}-kurs {d['study_type']} talabasi:")
        rp(d["student_name"], size=14)
    elif lang == "ru":
        rp(f"Студент {d['course']}-го курса ({d['study_type']}):")
        rp(d["student_name"], size=14)
    else:
        rp(f"{d['course']}-year student ({d['study_type']}):")
        rp(d["student_name"], size=14)

    if d["work_type"] == "kurs":
        cp("")
        if lang == "uz": rp("Ilmiy rahbar:")
        elif lang == "ru": rp("Научный руководитель:")
        else: rp("Scientific supervisor:")
        rp("______________________")

    cp(""); cp(""); cp(""); cp("")

    if lang == "uz":
        cp(f"Toshkent — {year}", size=13)
    elif lang == "ru":
        cp(f"Ташкент — {year}", size=13)
    else:
        cp(f"Tashkent — {year}", size=13)


def mundarija_mustaqil(doc: Document, secs: dict, lang: str):
    doc.add_page_break()
    if lang == "uz": h(doc, "MUNDARIJA")
    elif lang == "ru": h(doc, "ОГЛАВЛЕНИЕ")
    else: h(doc, "TABLE OF CONTENTS")
    doc.add_paragraph("")

    def row(text):
        ref_line(doc, text)

    if lang == "uz": row("Kirish ......................................................................................3")
    elif lang == "ru": row("Введение ......................................................................................3")
    else: row("Introduction ......................................................................................3")

    for i in range(1, 4):
        name = secs.get(f"BOLIM_{i}_NOMI", f"Bo'lim {i}")
        row(f"{i}. {name} .........................................................................{i*4+2}")

    if lang == "uz":
        row("Xulosa ......................................................................................17")
        row("Foydalanilgan adabiyotlar ..........................................................19")
    elif lang == "ru":
        row("Заключение ......................................................................................17")
        row("Список литературы ......................................................................19")
    else:
        row("Conclusion ......................................................................................17")
        row("References ......................................................................................19")


def mundarija_kurs(doc: Document, secs: dict, lang: str):
    doc.add_page_break()
    if lang == "uz": h(doc, "MUNDARIJA")
    elif lang == "ru": h(doc, "ОГЛАВЛЕНИЕ")
    else: h(doc, "TABLE OF CONTENTS")
    doc.add_paragraph("")

    def row(text): ref_line(doc, text)

    if lang == "uz": row("Kirish .................................................................................3")
    elif lang == "ru": row("Введение .................................................................................3")
    else: row("Introduction .................................................................................3")

    b1 = secs.get("I_BOB_NOMI", "NAZARIY ASOSLAR")
    row(f"I BOB. {b1.upper()} ..........................................5")
    for i in range(1, 4):
        nm = secs.get(f"PAR_1_{i}_NOMI", f"1.{i} § ")
        row(f"   1.{i} § {nm} .........................................................{i*3+5}")

    if lang == "uz": row("I Bob bo'yicha xulosalar ..........................................................16")
    elif lang == "ru": row("Выводы по I главе .......................................................................16")
    else: row("Chapter I conclusions ....................................................................16")

    b2 = secs.get("II_BOB_NOMI", "AMALIY QISM")
    row(f"II BOB. {b2.upper()} ........................................18")
    for i in range(1, 4):
        nm = secs.get(f"PAR_2_{i}_NOMI", f"2.{i} §")
        row(f"   2.{i} § {nm} .........................................................{i*3+18}")

    if lang == "uz":
        row("II Bob bo'yicha xulosalar .........................................................28")
        row("Xulosa .................................................................................30")
        row("Glossariy ...............................................................................32")
        row("Foydalanilgan adabiyotlar .........................................................34")
    elif lang == "ru":
        row("Выводы по II главе ......................................................................28")
        row("Заключение .................................................................................30")
        row("Глоссарий ...................................................................................32")
        row("Список литературы ......................................................................34")
    else:
        row("Chapter II conclusions ...................................................................28")
        row("Conclusion .................................................................................30")
        row("Glossary ...................................................................................32")
        row("References .................................................................................34")


# ── Hujjat yaratish (Mustaqil ish) ───────────────────────────────
def build_mustaqil(d: dict, secs: dict) -> bytes:
    doc = Document()
    doc_setup(doc)
    add_footer_pagenum(doc)
    lang = d["lang"]

    muqova(doc, d)
    mundarija_mustaqil(doc, secs, lang)

    # Kirish
    doc.add_page_break()
    if lang == "uz": h(doc, "KIRISH")
    elif lang == "ru": h(doc, "ВВЕДЕНИЕ")
    else: h(doc, "INTRODUCTION")
    body(doc, secs.get("KIRISH", ""))

    # 3 ta bo'lim
    for i in range(1, 4):
        doc.add_page_break()
        name = secs.get(f"BOLIM_{i}_NOMI", f"Bo'lim {i}")
        h(doc, f"{i}. {name}")
        body(doc, secs.get(f"BOLIM_{i}_MATN", ""))

    # Xulosa
    doc.add_page_break()
    if lang == "uz": h(doc, "XULOSA")
    elif lang == "ru": h(doc, "ЗАКЛЮЧЕНИЕ")
    else: h(doc, "CONCLUSION")
    body(doc, secs.get("XULOSA", ""))

    # Adabiyotlar
    doc.add_page_break()
    if lang == "uz": h(doc, "FOYDALANILGAN ADABIYOTLAR")
    elif lang == "ru": h(doc, "СПИСОК ЛИТЕРАТУРЫ")
    else: h(doc, "REFERENCES")
    for line in secs.get("ADABIYOTLAR", "").split("\n"):
        if line.strip():
            ref_line(doc, line)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ── Hujjat yaratish (Kurs ishi) ───────────────────────────────────
def build_kurs(d: dict, secs: dict) -> bytes:
    doc = Document()
    doc_setup(doc)
    add_footer_pagenum(doc)
    lang = d["lang"]

    muqova(doc, d)
    mundarija_kurs(doc, secs, lang)

    # Kirish
    doc.add_page_break()
    if lang == "uz": h(doc, "KIRISH")
    elif lang == "ru": h(doc, "ВВЕДЕНИЕ")
    else: h(doc, "INTRODUCTION")
    body(doc, secs.get("KIRISH", ""))

    # I Bob
    doc.add_page_break()
    b1 = secs.get("I_BOB_NOMI", "NAZARIY ASOSLAR")
    h(doc, f"I BOB. {b1}")
    for i in range(1, 4):
        nm = secs.get(f"PAR_1_{i}_NOMI", f"Paragraf 1.{i}")
        h(doc, f"1.{i} § {nm}", level=2)
        body(doc, secs.get(f"PAR_1_{i}_MATN", ""))

    # I Bob xulosalari
    doc.add_page_break()
    if lang == "uz": h(doc, "I BOB BO'YICHA XULOSALAR")
    elif lang == "ru": h(doc, "ВЫВОДЫ ПО I ГЛАВЕ")
    else: h(doc, "CHAPTER I CONCLUSIONS")
    body(doc, secs.get("I_BOB_XULOSA", ""))

    # II Bob
    doc.add_page_break()
    b2 = secs.get("II_BOB_NOMI", "AMALIY QISM")
    h(doc, f"II BOB. {b2}")
    for i in range(1, 4):
        nm = secs.get(f"PAR_2_{i}_NOMI", f"Paragraf 2.{i}")
        h(doc, f"2.{i} § {nm}", level=2)
        body(doc, secs.get(f"PAR_2_{i}_MATN", ""))

    # II Bob xulosalari
    doc.add_page_break()
    if lang == "uz": h(doc, "II BOB BO'YICHA XULOSALAR")
    elif lang == "ru": h(doc, "ВЫВОДЫ ПО II ГЛАВЕ")
    else: h(doc, "CHAPTER II CONCLUSIONS")
    body(doc, secs.get("II_BOB_XULOSA", ""))

    # Umumiy xulosa
    doc.add_page_break()
    if lang == "uz": h(doc, "XULOSA")
    elif lang == "ru": h(doc, "ЗАКЛЮЧЕНИЕ")
    else: h(doc, "CONCLUSION")
    body(doc, secs.get("XULOSA", ""))

    # Glossariy
    doc.add_page_break()
    if lang == "uz": h(doc, "GLOSSARIY")
    elif lang == "ru": h(doc, "ГЛОССАРИЙ")
    else: h(doc, "GLOSSARY")
    raw_gloss = secs.get("GLOSSARIY", "")
    for line in raw_gloss.split("\n"):
        line = line.strip()
        if not line:
            continue
        clean = re.sub(r"\*\*(.+?)\*\*", r"\1", line)
        ref_line(doc, clean)

    # Adabiyotlar
    doc.add_page_break()
    if lang == "uz": h(doc, "FOYDALANILGAN ADABIYOTLAR")
    elif lang == "ru": h(doc, "СПИСОК ЛИТЕРАТУРЫ")
    else: h(doc, "REFERENCES")
    for line in secs.get("ADABIYOTLAR", "").split("\n"):
        if line.strip():
            ref_line(doc, line)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ── Conversation handlers ─────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text(
        T["uz"]["welcome"], reply_markup=lang_kb(), parse_mode="Markdown"
    )
    return LANG


async def cb_lang(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    lang = q.data.replace("lang_", "")
    ctx.user_data["lang"] = lang
    await q.edit_message_text(T[lang]["work_type"], reply_markup=work_kb(lang), parse_mode="Markdown")
    return WORK_TYPE


async def cb_work(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    lang = ctx.user_data["lang"]
    ctx.user_data["work_type"] = q.data.replace("work_", "")
    await q.edit_message_text(T[lang]["university"], parse_mode="Markdown")
    return UNIVERSITY


async def msg_university(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    lang = ctx.user_data["lang"]
    ctx.user_data["university"] = update.message.text.strip()
    await update.message.reply_text(T[lang]["faculty"], parse_mode="Markdown")
    return FACULTY


async def msg_faculty(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    lang = ctx.user_data["lang"]
    ctx.user_data["faculty"] = update.message.text.strip()
    await update.message.reply_text(T[lang]["direction"], parse_mode="Markdown")
    return DIRECTION


async def msg_direction(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    lang = ctx.user_data["lang"]
    ctx.user_data["direction"] = update.message.text.strip()
    await update.message.reply_text(T[lang]["subject"], parse_mode="Markdown")
    return SUBJECT


async def msg_subject(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    lang = ctx.user_data["lang"]
    ctx.user_data["subject"] = update.message.text.strip()
    await update.message.reply_text(T[lang]["student"], parse_mode="Markdown")
    return STUDENT_NAME


async def msg_student(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    lang = ctx.user_data["lang"]
    ctx.user_data["student_name"] = update.message.text.strip()
    await update.message.reply_text(T[lang]["course"], parse_mode="Markdown")
    return COURSE


async def msg_course(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    lang = ctx.user_data["lang"]
    text = update.message.text.strip()
    if not text.isdigit() or int(text) not in range(1, 5):
        await update.message.reply_text(T[lang]["invalid_course"])
        return COURSE
    ctx.user_data["course"] = text
    await update.message.reply_text(T[lang]["study_type"], reply_markup=study_kb(lang))
    return STUDY_TYPE


async def cb_study(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    lang = ctx.user_data["lang"]
    study_key = q.data.replace("study_", "")
    ctx.user_data["study_type"] = STUDY_TYPES[lang][study_key]
    await q.edit_message_text(T[lang]["topic"], parse_mode="Markdown")
    return TOPIC


async def msg_topic(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    lang = ctx.user_data["lang"]
    ctx.user_data["topic"] = update.message.text.strip()
    d = ctx.user_data

    status = await update.message.reply_text(T[lang]["generating"], parse_mode="Markdown")

    try:
        prompt = prompt_mustaqil(d) if d["work_type"] == "mustaqil" else prompt_kurs(d)

        response = claude.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8000,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = response.content[0].text
        secs = parse_sections(raw)

        if d["work_type"] == "mustaqil":
            docx_bytes = build_mustaqil(d, secs)
            fname = f"mustaqil_ish_{d['topic'][:25].replace(' ', '_')}.docx"
        else:
            docx_bytes = build_kurs(d, secs)
            fname = f"kurs_ishi_{d['topic'][:25].replace(' ', '_')}.docx"

        await status.edit_text(T[lang]["done"], parse_mode="Markdown")
        await update.message.reply_document(
            document=io.BytesIO(docx_bytes),
            filename=fname,
            caption=f"📄 {d['work_type'].upper()} | {d['topic']}",
        )

    except Exception as e:
        logger.error(f"Xatolik: {e}")
        await status.edit_text(T[lang]["error"])

    await update.message.reply_text(T[lang]["restart"])
    return ConversationHandler.END


async def cmd_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    lang = ctx.user_data.get("lang", "uz")
    await update.message.reply_text(T[lang]["restart"])
    return ConversationHandler.END


# ── Main ──────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("start", cmd_start)],
        states={
            LANG:         [CallbackQueryHandler(cb_lang,       pattern="^lang_")],
            WORK_TYPE:    [CallbackQueryHandler(cb_work,       pattern="^work_")],
            UNIVERSITY:   [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_university)],
            FACULTY:      [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_faculty)],
            DIRECTION:    [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_direction)],
            SUBJECT:      [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_subject)],
            STUDENT_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_student)],
            COURSE:       [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_course)],
            STUDY_TYPE:   [CallbackQueryHandler(cb_study,      pattern="^study_")],
            TOPIC:        [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_topic)],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
        allow_reentry=True,
    )

    app.add_handler(conv)
    logger.info("✅ Bot ishga tushdi!")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
