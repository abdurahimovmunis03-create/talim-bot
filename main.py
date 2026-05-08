#!/usr/bin/env python3
"""
O'zbekiston OTM Ta'lim Boti v2
Reja → Bo'limlar → Word fayl arxitekturasi
"""

import os, io, re, logging, json
import google.generativeai as genai
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler,
    filters, ContextTypes
)
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ["BOT_TOKEN"]
GEMINI_KEY = os.environ["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_KEY)
gemini = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    generation_config={"max_output_tokens": 8192, "temperature": 0.7}
)

# Bosqichlar
LANG, WORK_TYPE, UNIVERSITY, FACULTY, DIRECTION, SUBJECT, STUDENT_NAME, COURSE, STUDY_TYPE, TOPIC = range(10)

# ── Matnlar ───────────────────────────────────────────────────────
T = {
    "uz": {
        "welcome": "👋 *Assalomu alaykum!*\n\nMen O'zbekiston OTM standartlarida yozma ishlar tayyorlab beruvchi botman.\n\n📝 *Mustaqil ish* — 15–20 sahifa\n📚 *Kurs ishi* — 25–35 sahifa\n\nTayyor ish Word (.docx) formatida yuboriladi.\n\n🌐 *Tilni tanlang:*",
        "work_type": "📋 Qanday ish kerak?",
        "mustaqil": "📝 Mustaqil ish",
        "kurs": "📚 Kurs ishi",
        "university": "🏛 *Universitetingiz* nomini to'liq kiriting:",
        "faculty": "🏫 *Fakultetingiz* nomini kiriting:",
        "direction": "📖 *Yo'nalishingizni* kiriting:",
        "subject": "📚 *Fan nomini* kiriting:",
        "student": "👤 *Ism va familiyangizni* kiriting:",
        "course": "🎓 *Nechanchi kursda* o'qiysiz? (1-4):",
        "study_type": "📅 O'qish shaklini tanlang:",
        "kunduzgi": "☀️ Kunduzgi", "sirtqi": "🌙 Sirtqi", "kechki": "🌆 Kechki",
        "topic": "✏️ *Ish mavzusini* kiriting:\n\n⚡ Mavzu qanchalik aniq bo'lsa, ish shunchalik sifatli bo'ladi!",
        "planning": "📋 *Reja tuzilmoqda...*\n\nBir daqiqa kuting ⏳",
        "generating_section": "✍️ *{section} yozilmoqda...*\n\n_{progress}_",
        "combining": "📎 *Fayl tayyorlanmoqda...*",
        "done": "✅ *Ish tayyor!* Word formatida yuborilmoqda...",
        "error": "❌ Xatolik yuz berdi. /start bilan qaytadan urinib ko'ring.",
        "invalid_course": "⚠️ 1 dan 4 gacha raqam kiriting.",
        "restart": "\n\n🔄 Yangi ish uchun: /start",
    },
    "ru": {
        "welcome": "👋 *Здравствуйте!*\n\nЯ бот для написания учебных работ по стандартам узбекских вузов.\n\n📝 *Самостоятельная работа* — 15–20 страниц\n📚 *Курсовая работа* — 25–35 страниц\n\nГотовая работа отправляется в формате Word (.docx).\n\n🌐 *Выберите язык:*",
        "work_type": "📋 Какая работа нужна?",
        "mustaqil": "📝 Самостоятельная работа",
        "kurs": "📚 Курсовая работа",
        "university": "🏛 Введите *полное название университета*:",
        "faculty": "🏫 Введите *название факультета*:",
        "direction": "📖 Введите *направление обучения*:",
        "subject": "📚 Введите *название предмета*:",
        "student": "👤 Введите *ваше имя и фамилию*:",
        "course": "🎓 На *каком курсе* вы учитесь? (1-4):",
        "study_type": "📅 Выберите форму обучения:",
        "kunduzgi": "☀️ Дневное", "sirtqi": "🌙 Заочное", "kechki": "🌆 Вечернее",
        "topic": "✏️ Введите *тему работы*:\n\n⚡ Чем точнее тема, тем качественнее работа!",
        "planning": "📋 *Составляется план...*\n\nОдну минуту ⏳",
        "generating_section": "✍️ *Пишется {section}...*\n\n_{progress}_",
        "combining": "📎 *Файл готовится...*",
        "done": "✅ *Работа готова!* Отправляю в формате Word...",
        "error": "❌ Произошла ошибка. Попробуйте /start снова.",
        "invalid_course": "⚠️ Введите цифру от 1 до 4.",
        "restart": "\n\n🔄 Для новой работы: /start",
    },
    "en": {
        "welcome": "👋 *Hello!*\n\nI'm a bot for writing academic papers per Uzbekistan HEI standards.\n\n📝 *Independent work* — 15–20 pages\n📚 *Course work* — 25–35 pages\n\nThe finished work is sent in Word (.docx) format.\n\n🌐 *Choose language:*",
        "work_type": "📋 What type of work do you need?",
        "mustaqil": "📝 Independent work",
        "kurs": "📚 Course work",
        "university": "🏛 Enter the *full university name*:",
        "faculty": "🏫 Enter your *faculty name*:",
        "direction": "📖 Enter your *study direction*:",
        "subject": "📚 Enter the *subject name*:",
        "student": "👤 Enter your *full name*:",
        "course": "🎓 Which *year* are you in? (1-4):",
        "study_type": "📅 Select study type:",
        "kunduzgi": "☀️ Full-time", "sirtqi": "🌙 Part-time", "kechki": "🌆 Evening",
        "topic": "✏️ Enter the *work topic*:\n\n⚡ The more specific the topic, the better the work!",
        "planning": "📋 *Creating outline...*\n\nOne moment ⏳",
        "generating_section": "✍️ *Writing {section}...*\n\n_{progress}_",
        "combining": "📎 *Preparing file...*",
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

# ── Gemini yordamchi funksiya ─────────────────────────────────────
def ask_gemini(prompt: str) -> str:
    response = gemini.generate_content(prompt)
    return response.text.strip()

# ── REJA TUZISH ───────────────────────────────────────────────────
def make_plan_prompt(d: dict) -> str:
    lang, topic, subject = d["lang"], d["topic"], d["subject"]
    work = d["work_type"]

    if lang == "uz":
        if work == "mustaqil":
            return f"""Sen O'zbekiston universiteti talabasi uchun "{subject}" fanidan "{topic}" mavzusida mustaqil ish rejasini tuzmoqdasiz.

Faqat JSON formatida javob ber, boshqa hech narsa yozma:

{{
  "kirish_rejasi": "Kirish bo'limi uchun 3-4 ta asosiy nuqta (vergul bilan ajratilgan)",
  "bolim1_nomi": "Birinchi bo'lim sarlavhasi",
  "bolim1_kichik": ["1.1 kichik sarlavha", "1.2 kichik sarlavha", "1.3 kichik sarlavha"],
  "bolim2_nomi": "Ikkinchi bo'lim sarlavhasi",
  "bolim2_kichik": ["2.1 kichik sarlavha", "2.2 kichik sarlavha", "2.3 kichik sarlavha"],
  "bolim3_nomi": "Uchinchi bo'lim sarlavhasi",
  "bolim3_kichik": ["3.1 kichik sarlavha", "3.2 kichik sarlavha", "3.3 kichik sarlavha"],
  "xulosa_rejasi": "Xulosa uchun 3-4 ta asosiy nuqta (vergul bilan ajratilgan)"
}}"""
        else:
            return f"""Sen O'zbekiston universiteti talabasi uchun "{subject}" fanidan "{topic}" mavzusida kurs ishi rejasini tuzmoqdasiz.

Faqat JSON formatida javob ber:

{{
  "kirish_rejasi": "Kirish elementlari: dolzarbligi, o'rganilganlik darajasi, maqsad, vazifalar, ob'ekt, predmet, ahamiyati",
  "bob1_nomi": "I Bobning sarlavhasi (nazariy asoslar)",
  "bob1_par1": "1.1 paragraf sarlavhasi",
  "bob1_par2": "1.2 paragraf sarlavhasi",
  "bob1_par3": "1.3 paragraf sarlavhasi",
  "bob2_nomi": "II Bobning sarlavhasi (amaliy qism)",
  "bob2_par1": "2.1 paragraf sarlavhasi",
  "bob2_par2": "2.2 paragraf sarlavhasi",
  "bob2_par3": "2.3 paragraf sarlavhasi",
  "xulosa_rejasi": "Xulosa uchun asosiy fikrlar"
}}"""
    elif lang == "ru":
        if work == "mustaqil":
            return f"""Составь план самостоятельной работы по предмету "{subject}" на тему "{topic}".

Ответь только в формате JSON:

{{
  "kirish_rejasi": "Основные пункты введения через запятую",
  "bolim1_nomi": "Название первого раздела",
  "bolim1_kichik": ["1.1 подзаголовок", "1.2 подзаголовок", "1.3 подзаголовок"],
  "bolim2_nomi": "Название второго раздела",
  "bolim2_kichik": ["2.1 подзаголовок", "2.2 подзаголовок", "2.3 подзаголовок"],
  "bolim3_nomi": "Название третьего раздела",
  "bolim3_kichik": ["3.1 подзаголовок", "3.2 подзаголовок", "3.3 подзаголовок"],
  "xulosa_rejasi": "Основные пункты заключения через запятую"
}}"""
        else:
            return f"""Составь план курсовой работы по предмету "{subject}" на тему "{topic}".

Ответь только в формате JSON:

{{
  "kirish_rejasi": "Элементы введения: актуальность, изученность, цель, задачи, объект, предмет, значимость",
  "bob1_nomi": "Название I главы (теоретические основы)",
  "bob1_par1": "Название параграфа 1.1",
  "bob1_par2": "Название параграфа 1.2",
  "bob1_par3": "Название параграфа 1.3",
  "bob2_nomi": "Название II главы (практическая часть)",
  "bob2_par1": "Название параграфа 2.1",
  "bob2_par2": "Название параграфа 2.2",
  "bob2_par3": "Название параграфа 2.3",
  "xulosa_rejasi": "Основные мысли заключения"
}}"""
    else:
        if work == "mustaqil":
            return f"""Create an outline for an independent work on "{subject}", topic: "{topic}".

Reply only in JSON format:

{{
  "kirish_rejasi": "Main introduction points separated by comma",
  "bolim1_nomi": "First section title",
  "bolim1_kichik": ["1.1 subtitle", "1.2 subtitle", "1.3 subtitle"],
  "bolim2_nomi": "Second section title",
  "bolim2_kichik": ["2.1 subtitle", "2.2 subtitle", "2.3 subtitle"],
  "bolim3_nomi": "Third section title",
  "bolim3_kichik": ["3.1 subtitle", "3.2 subtitle", "3.3 subtitle"],
  "xulosa_rejasi": "Main conclusion points separated by comma"
}}"""
        else:
            return f"""Create an outline for a course work on "{subject}", topic: "{topic}".

Reply only in JSON format:

{{
  "kirish_rejasi": "Introduction elements: relevance, literature, goal, tasks, object, subject, significance",
  "bob1_nomi": "Chapter I title (theoretical foundations)",
  "bob1_par1": "Section 1.1 title",
  "bob1_par2": "Section 1.2 title",
  "bob1_par3": "Section 1.3 title",
  "bob2_nomi": "Chapter II title (practical part)",
  "bob2_par1": "Section 2.1 title",
  "bob2_par2": "Section 2.2 title",
  "bob2_par3": "Section 2.3 title",
  "xulosa_rejasi": "Main conclusion thoughts"
}}"""

# ── BO'LIM YOZISH ─────────────────────────────────────────────────
def write_section(d: dict, section_name: str, section_type: str, plan: dict, extra: str = "") -> str:
    lang, topic, subject = d["lang"], d["topic"], d["subject"]
    course, study = d["course"], d["study_type"]

    if lang == "uz":
        base = f'"{subject}" fanidan "{topic}" mavzusidagi yozma ish uchun'
        inst = f"O'zbek tilida, ilmiy uslubda yoz. Kamida 600 so'z. Paragraflar bilan. Sarlavha yozma, faqat matn."
    elif lang == "ru":
        base = f'для работы по предмету "{subject}" на тему "{topic}"'
        inst = f"Пиши на русском языке, научным стилем. Минимум 600 слов. С абзацами. Без заголовка, только текст."
    else:
        base = f'for the work on "{subject}", topic: "{topic}"'
        inst = f"Write in English, academic style. Minimum 600 words. With paragraphs. No heading, just text."

    if section_type == "kirish":
        if lang == "uz":
            prompt = f"""{base} KIRISH bo'limini yoz.
Quyidagilarni qamrab ol: mavzuning dolzarbligi, o'rganilganlik darajasi, maqsad, vazifalar, ob'ekt va predmet, ahamiyati.
{inst}"""
        elif lang == "ru":
            prompt = f"""{base} напиши ВВЕДЕНИЕ.
Включи: актуальность, степень изученности, цель, задачи, объект и предмет, значимость.
{inst}"""
        else:
            prompt = f"""{base} write the INTRODUCTION.
Include: relevance, literature review, goal, tasks, object and subject, significance.
{inst}"""

    elif section_type == "bolim":
        if lang == "uz":
            prompt = f"""{base} "{section_name}" bo'limini yoz.
Bu bo'lim mavzuning muhim jihatlarini chuqur tahlil qilishi kerak.
Ilmiy dalillar, misollar va tahlillar bilan boyit.
{inst}"""
        elif lang == "ru":
            prompt = f"""{base} напиши раздел "{section_name}".
Раздел должен глубоко анализировать важные аспекты темы.
Обогати научными доказательствами, примерами и анализом.
{inst}"""
        else:
            prompt = f"""{base} write the section "{section_name}".
The section should deeply analyze important aspects of the topic.
Enrich with scientific evidence, examples and analysis.
{inst}"""

    elif section_type == "bob_xulosa":
        if lang == "uz":
            prompt = f"""{base} "{section_name}" bo'yicha XULOSALAR yoz.
Bu bobda ko'rib chiqilgan asosiy fikrlarni qisqacha jamla.
{inst}"""
        elif lang == "ru":
            prompt = f"""{base} напиши ВЫВОДЫ по "{section_name}".
Кратко обобщи основные мысли рассмотренные в этой главе.
{inst}"""
        else:
            prompt = f"""{base} write CONCLUSIONS for "{section_name}".
Briefly summarize the main ideas covered in this chapter.
{inst}"""

    elif section_type == "xulosa":
        if lang == "uz":
            prompt = f"""{base} umumiy XULOSA yoz.
Butun ishning asosiy natijalarini, topilmalarini va tavsiyalarni yoz.
Kamida 400 so'z. {inst}"""
        elif lang == "ru":
            prompt = f"""{base} напиши общее ЗАКЛЮЧЕНИЕ.
Напиши основные результаты, выводы и рекомендации всей работы.
Минимум 400 слов. {inst}"""
        else:
            prompt = f"""{base} write the general CONCLUSION.
Write the main results, findings and recommendations of the entire work.
Minimum 400 words. {inst}"""

    elif section_type == "glossariy":
        if lang == "uz":
            prompt = f"""{base} GLOSSARIY tuz.
Mavzu bilan bog'liq 15-20 ta muhim atamani izohlanglar.
Har bir atama: "Atama nomi — ta'rifi (2-3 jumlada)"
Faqat atamalar ro'yxati, boshqa narsa yozma."""
        elif lang == "ru":
            prompt = f"""{base} составь ГЛОССАРИЙ.
15-20 важных терминов связанных с темой.
Каждый термин: "Название термина — определение (2-3 предложения)"
Только список терминов, ничего больше."""
        else:
            prompt = f"""{base} create a GLOSSARY.
15-20 important terms related to the topic.
Each term: "Term name — definition (2-3 sentences)"
Only list of terms, nothing else."""

    elif section_type == "adabiyotlar":
        if lang == "uz":
            prompt = f"""{base} FOYDALANILGAN ADABIYOTLAR ro'yxatini tuz.
Kamida 12 ta manba. Format:
1. Muallif I.O. Kitob nomi. – Shahar: Nashriyot, Yil.
Oxirida 4-5 ta internet manba ham qo'sh:
13. Muallif. Maqola nomi // Sayt nomi. URL: https://... (Murojaat: sana)
Faqat ro'yxat, boshqa narsa yozma."""
        elif lang == "ru":
            prompt = f"""{base} составь СПИСОК ЛИТЕРАТУРЫ.
Минимум 12 источников. Формат:
1. Автор И.О. Название книги. – Город: Издательство, Год.
В конце добавь 4-5 интернет-источника.
Только список, ничего больше."""
        else:
            prompt = f"""{base} create a REFERENCES list.
Minimum 12 sources. Format:
1. Author. Book Title. – City: Publisher, Year.
Add 4-5 internet sources at the end.
Only the list, nothing else."""

    return ask_gemini(prompt)


# ── Word hujjat yaratish ──────────────────────────────────────────
def doc_setup(doc):
    sec = doc.sections[0]
    sec.left_margin = Cm(3); sec.right_margin = Cm(1.5)
    sec.top_margin = Cm(2); sec.bottom_margin = Cm(2)
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"; normal.font.size = Pt(14)
    pf = normal.paragraph_format
    pf.line_spacing = Pt(21); pf.first_line_indent = Cm(1.25)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_before = Pt(0); pf.space_after = Pt(0)

def add_footer(doc):
    for sec in doc.sections:
        p = sec.footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.font.name = "Times New Roman"; run.font.size = Pt(12)
        for tag, txt in [("begin",""), ("end","")]:
            el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), tag)
            run._r.append(el)
            if tag == "begin":
                ins = OxmlElement("w:instrText"); ins.text = "PAGE"
                run._r.append(ins)

def heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = Pt(21)
    r = p.add_run(text.upper() if level == 1 else text)
    r.bold = True; r.font.name = "Times New Roman"
    r.font.size = Pt(16 if level == 1 else 14)

def body_text(doc, text):
    for chunk in text.split("\n\n"):
        chunk = chunk.strip()
        if not chunk: continue
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(1.25)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = Pt(21)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(chunk)
        r.font.name = "Times New Roman"; r.font.size = Pt(14)

def ref_text(doc, text):
    for line in text.split("\n"):
        line = line.strip()
        if not line: continue
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = Pt(21)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(line)
        r.font.name = "Times New Roman"; r.font.size = Pt(14)

def muqova(doc, d):
    import datetime
    year = datetime.datetime.now().year
    lang = d["lang"]

    def cp(txt, bold=False, size=13, space=0):
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(space)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(19)
        r = p.add_run(txt); r.bold = bold
        r.font.name = "Times New Roman"; r.font.size = Pt(size)

    def rp(txt, size=13):
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.right_indent = Cm(1.5)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(19)
        r = p.add_run(txt)
        r.font.name = "Times New Roman"; r.font.size = Pt(size)

    if lang == "uz":
        cp("O'ZBEKISTON RESPUBLIKASI OLIY VA O'RTA MAXSUS TA'LIM VAZIRLIGI", bold=True, size=12)
    elif lang == "ru":
        cp("МИНИСТЕРСТВО ВЫСШЕГО И СРЕДНЕГО СПЕЦИАЛЬНОГО ОБРАЗОВАНИЯ\nРЕСПУБЛИКИ УЗБЕКИСТАН", bold=True, size=12)
    else:
        cp("MINISTRY OF HIGHER AND SECONDARY SPECIALISED EDUCATION\nOF THE REPUBLIC OF UZBEKISTAN", bold=True, size=12)

    cp(""); cp(d["university"].upper(), bold=True, size=13, space=6)
    cp(""); cp(d["faculty"], size=12); cp(d["direction"], size=12)
    cp(""); cp(""); cp("")

    if lang == "uz": cp(f'"{d["subject"]}"', size=12); cp("fanidan", size=12)
    elif lang == "ru": cp(f'По предмету: "{d["subject"]}"', size=12)
    else: cp(f'Subject: "{d["subject"]}"', size=12)

    cp(""); cp(f'"{d["topic"].upper()}"', bold=True, size=14, space=4); cp("")

    if lang == "uz": wname = "MUSTAQIL ISH" if d["work_type"] == "mustaqil" else "KURS ISHI"
    elif lang == "ru": wname = "САМОСТОЯТЕЛЬНАЯ РАБОТА" if d["work_type"] == "mustaqil" else "КУРСОВАЯ РАБОТА"
    else: wname = "INDEPENDENT WORK" if d["work_type"] == "mustaqil" else "COURSE WORK"
    cp(wname, bold=True, size=16, space=6)

    cp(""); cp(""); cp("")

    if lang == "uz": rp(f"{d['course']}-kurs {d['study_type']} talabasi:"); rp(d["student_name"], size=14)
    elif lang == "ru": rp(f"Студент {d['course']}-го курса ({d['study_type']}):"); rp(d["student_name"], size=14)
    else: rp(f"{d['course']}-year student ({d['study_type']}):"); rp(d["student_name"], size=14)

    if d["work_type"] == "kurs":
        cp("")
        if lang == "uz": rp("Ilmiy rahbar:")
        elif lang == "ru": rp("Научный руководитель:")
        else: rp("Scientific supervisor:")
        rp("______________________")

    cp(""); cp(""); cp(""); cp("")
    if lang == "uz": cp(f"Toshkent — {year}", size=13)
    elif lang == "ru": cp(f"Ташкент — {year}", size=13)
    else: cp(f"Tashkent — {year}", size=13)


# ── Hujjat yig'ish ────────────────────────────────────────────────
def build_doc(d: dict, sections: dict) -> bytes:
    doc = Document()
    doc_setup(doc)
    add_footer(doc)
    lang = d["lang"]
    plan = d["plan"]

    muqova(doc, d)

    if d["work_type"] == "mustaqil":
        # Kirish
        doc.add_page_break()
        if lang == "uz": heading(doc, "KIRISH")
        elif lang == "ru": heading(doc, "ВВЕДЕНИЕ")
        else: heading(doc, "INTRODUCTION")
        body_text(doc, sections.get("kirish", ""))

        # 3 bo'lim
        for i in range(1, 4):
            doc.add_page_break()
            nom = plan.get(f"bolim{i}_nomi", f"Bo'lim {i}")
            heading(doc, f"{i}. {nom}")
            body_text(doc, sections.get(f"bolim{i}", ""))

        # Xulosa
        doc.add_page_break()
        if lang == "uz": heading(doc, "XULOSA")
        elif lang == "ru": heading(doc, "ЗАКЛЮЧЕНИЕ")
        else: heading(doc, "CONCLUSION")
        body_text(doc, sections.get("xulosa", ""))

        # Adabiyotlar
        doc.add_page_break()
        if lang == "uz": heading(doc, "FOYDALANILGAN ADABIYOTLAR")
        elif lang == "ru": heading(doc, "СПИСОК ЛИТЕРАТУРЫ")
        else: heading(doc, "REFERENCES")
        ref_text(doc, sections.get("adabiyotlar", ""))

    else:  # kurs ishi
        # Kirish
        doc.add_page_break()
        if lang == "uz": heading(doc, "KIRISH")
        elif lang == "ru": heading(doc, "ВВЕДЕНИЕ")
        else: heading(doc, "INTRODUCTION")
        body_text(doc, sections.get("kirish", ""))

        # I Bob
        doc.add_page_break()
        b1 = plan.get("bob1_nomi", "NAZARIY ASOSLAR")
        heading(doc, f"I BOB. {b1}")
        for i in range(1, 4):
            nm = plan.get(f"bob1_par{i}", f"Paragraf 1.{i}")
            heading(doc, f"1.{i} § {nm}", level=2)
            body_text(doc, sections.get(f"bob1_par{i}", ""))

        doc.add_page_break()
        if lang == "uz": heading(doc, "I BOB BO'YICHA XULOSALAR")
        elif lang == "ru": heading(doc, "ВЫВОДЫ ПО I ГЛАВЕ")
        else: heading(doc, "CHAPTER I CONCLUSIONS")
        body_text(doc, sections.get("bob1_xulosa", ""))

        # II Bob
        doc.add_page_break()
        b2 = plan.get("bob2_nomi", "AMALIY QISM")
        heading(doc, f"II BOB. {b2}")
        for i in range(1, 4):
            nm = plan.get(f"bob2_par{i}", f"Paragraf 2.{i}")
            heading(doc, f"2.{i} § {nm}", level=2)
            body_text(doc, sections.get(f"bob2_par{i}", ""))

        doc.add_page_break()
        if lang == "uz": heading(doc, "II BOB BO'YICHA XULOSALAR")
        elif lang == "ru": heading(doc, "ВЫВОДЫ ПО II ГЛАВЕ")
        else: heading(doc, "CHAPTER II CONCLUSIONS")
        body_text(doc, sections.get("bob2_xulosa", ""))

        # Xulosa
        doc.add_page_break()
        if lang == "uz": heading(doc, "XULOSA")
        elif lang == "ru": heading(doc, "ЗАКЛЮЧЕНИЕ")
        else: heading(doc, "CONCLUSION")
        body_text(doc, sections.get("xulosa", ""))

        # Glossariy
        doc.add_page_break()
        if lang == "uz": heading(doc, "GLOSSARIY")
        elif lang == "ru": heading(doc, "ГЛОССАРИЙ")
        else: heading(doc, "GLOSSARY")
        ref_text(doc, sections.get("glossariy", ""))

        # Adabiyotlar
        doc.add_page_break()
        if lang == "uz": heading(doc, "FOYDALANILGAN ADABIYOTLAR")
        elif lang == "ru": heading(doc, "СПИСОК ЛИТЕРАТУРЫ")
        else: heading(doc, "REFERENCES")
        ref_text(doc, sections.get("adabiyotlar", ""))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ── Conversation handlers ─────────────────────────────────────────
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text(T["uz"]["welcome"], reply_markup=lang_kb(), parse_mode="Markdown")
    return LANG

async def cb_lang(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
    lang = q.data.replace("lang_", "")
    ctx.user_data["lang"] = lang
    await q.edit_message_text(T[lang]["work_type"], reply_markup=work_kb(lang), parse_mode="Markdown")
    return WORK_TYPE

async def cb_work(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query; await q.answer()
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
    q = update.callback_query; await q.answer()
    lang = ctx.user_data["lang"]
    study_key = q.data.replace("study_", "")
    ctx.user_data["study_type"] = STUDY_TYPES[lang][study_key]
    await q.edit_message_text(T[lang]["topic"], parse_mode="Markdown")
    return TOPIC

async def msg_topic(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    lang = ctx.user_data["lang"]
    ctx.user_data["topic"] = update.message.text.strip()
    d = ctx.user_data

    status = await update.message.reply_text(T[lang]["planning"], parse_mode="Markdown")

    try:
        # 1. REJA TUZISH
        logger.info("Reja tuzilmoqda...")
        plan_raw = ask_gemini(make_plan_prompt(d))
        clean = plan_raw.replace("```json","").replace("```","").strip()
        plan = json.loads(clean)
        d["plan"] = plan
        logger.info(f"Reja tayyor: {list(plan.keys())}")

        sections = {}

        if d["work_type"] == "mustaqil":
            steps = [
                ("kirish", "bolim", "Kirish / Введение", "1/6"),
                ("bolim1", "bolim", plan.get("bolim1_nomi","Bo'lim 1"), "2/6"),
                ("bolim2", "bolim", plan.get("bolim2_nomi","Bo'lim 2"), "3/6"),
                ("bolim3", "bolim", plan.get("bolim3_nomi","Bo'lim 3"), "4/6"),
                ("xulosa", "xulosa", "Xulosa / Заключение", "5/6"),
                ("adabiyotlar", "adabiyotlar", "Adabiyotlar / Литература", "6/6"),
            ]
        else:
            steps = [
                ("kirish", "kirish", "Kirish / Введение", "1/10"),
                ("bob1_par1", "bolim", plan.get("bob1_par1","1.1 §"), "2/10"),
                ("bob1_par2", "bolim", plan.get("bob1_par2","1.2 §"), "3/10"),
                ("bob1_par3", "bolim", plan.get("bob1_par3","1.3 §"), "4/10"),
                ("bob1_xulosa", "bob_xulosa", plan.get("bob1_nomi","I Bob"), "5/10"),
                ("bob2_par1", "bolim", plan.get("bob2_par1","2.1 §"), "6/10"),
                ("bob2_par2", "bolim", plan.get("bob2_par2","2.2 §"), "7/10"),
                ("bob2_par3", "bolim", plan.get("bob2_par3","2.3 §"), "8/10"),
                ("bob2_xulosa", "bob_xulosa", plan.get("bob2_nomi","II Bob"), "9/10"),
                ("xulosa", "xulosa", "Xulosa / Заключение", "9/10"),
                ("glossariy", "glossariy", "Glossariy / Глоссарий", "9/10"),
                ("adabiyotlar", "adabiyotlar", "Adabiyotlar / Литература", "10/10"),
            ]

        # 2. BO'LIMLARNI YOZISH
        for key, stype, sname, progress in steps:
            logger.info(f"Yozilmoqda: {key} - {sname}")
            msg = T[lang]["generating_section"].format(section=sname, progress=progress)
            await status.edit_text(msg, parse_mode="Markdown")
            sections[key] = write_section(d, sname, stype, plan)
            logger.info(f"Tayyor: {key} - {len(sections[key])} belgi")

        # 3. FAYLNI YIGISH
        await status.edit_text(T[lang]["combining"], parse_mode="Markdown")
        docx_bytes = build_doc(d, sections)

        fname = f"{'mustaqil_ish' if d['work_type']=='mustaqil' else 'kurs_ishi'}_{d['topic'][:20].replace(' ','_')}.docx"

        await status.edit_text(T[lang]["done"], parse_mode="Markdown")
        await update.message.reply_document(
            document=io.BytesIO(docx_bytes),
            filename=fname,
            caption=f"📄 {d['topic']}",
        )

    except Exception as e:
        logger.error(f"Xatolik: {type(e).__name__}: {str(e)}")
        await status.edit_text(T[lang]["error"])

    await update.message.reply_text(T[lang]["restart"])
    return ConversationHandler.END

async def cmd_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    lang = ctx.user_data.get("lang", "uz")
    await update.message.reply_text(T[lang]["restart"])
    return ConversationHandler.END

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", cmd_start)],
        states={
            LANG:         [CallbackQueryHandler(cb_lang,    pattern="^lang_")],
            WORK_TYPE:    [CallbackQueryHandler(cb_work,    pattern="^work_")],
            UNIVERSITY:   [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_university)],
            FACULTY:      [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_faculty)],
            DIRECTION:    [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_direction)],
            SUBJECT:      [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_subject)],
            STUDENT_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_student)],
            COURSE:       [MessageHandler(filters.TEXT & ~filters.COMMAND, msg_course)],
            STUDY_TYPE:   [CallbackQueryHandler(cb_study,   pattern="^study_")],
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
