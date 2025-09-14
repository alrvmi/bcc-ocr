# parser.py
import os
import re
import json
from datetime import datetime
from collections import OrderedDict

import pandas as pd
from docx import Document as DocxDocument  # для fallback чтения docx

# ----------------- Настройки -----------------
OUTPUT_BASE = os.path.join("data", "output")  # папка с результатами prod.py
RESULT_XLSX = os.path.join(OUTPUT_BASE, "results.xlsx")
CONF_THRESHOLD = None  # если хочешь фильтровать по confidence, ставь 0.5 и т.д.
# ---------------------------------------------

# ---- Полезные маппинги для русских месяцев ----
RU_MONTHS = {
    "янв": "01", "января": "01", "январь": "01",
    "фев": "02", "февраля": "02", "февраль": "02",
    "мар": "03", "марта": "03", "март": "03",
    "апр": "04", "апреля": "04", "апрель": "04",
    "май": "05", "мая": "05",
    "июн": "06", "июня": "06", "июнь": "06",
    "июл": "07", "июля": "07", "июль": "07",
    "авг": "08", "августа": "08", "август": "08",
    "сен": "09", "сентября": "09", "сентябрь": "09",
    "окт": "10", "октября": "10", "октябрь": "10",
    "ноя": "11", "ноября": "11", "ноябрь": "11",
    "дек": "12", "декабря": "12", "декабрь": "12"
}
# -----------------------------------------------

# ---------- Утилиты для чтения результата OCR ----------
def read_txt_if_exists(folder):
    path = os.path.join(folder, "result.txt")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return None

def read_docx_if_exists(folder):
    path = os.path.join(folder, "result.docx")
    if os.path.exists(path):
        doc = DocxDocument(path)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n".join(paragraphs)
    return None

def load_text_from_output_folder(folder):
    """
    Пытаемся получить текст результата OCR из папки (result.txt или result.docx).
    Возвращает строку с текстом (или "" если ничего не найдено)
    """
    t = read_txt_if_exists(folder)
    if t:
        return t
    t2 = read_docx_if_exists(folder)
    if t2:
        return t2
    return ""

# ---------- Нормализация и парсинг чисел/даты ----------
def normalize_whitespace(s):
    return re.sub(r'\s+', ' ', s).strip()

def normalize_amount_str(s):
    """Пытаемся привести строку суммы к float-представлению (в строковом виде)."""
    if not s:
        return None
    s = s.replace('\u00A0', '').replace('\u202F', '')  # NBSP
    # Оставим только цифры, пробелы, точку, запятую
    s = re.sub(r'[^\d\.,]', '', s)
    if not s:
        return None
    # Если есть и запятая и точка — считаем, что запятые это разделители тысяч
    if ',' in s and '.' in s:
        s = s.replace(',', '')
    elif ',' in s and '.' not in s:
        s = s.replace(',', '.')
    # теперь удалим лишние символы кроме цифр и точки
    s = re.sub(r'[^0-9\.]', '', s)
    if not s:
        return None
    try:
        val = float(s)
        return val
    except:
        return None

def try_parse_date(s):
    """Пробуем спарсить дату из строки, возвращаем ISO yyyy-mm-dd или None"""
    if not s or not isinstance(s, str):
        return None
    s = s.strip()
    # 1) ищем цифровые форматы dd.mm.yyyy / dd.mm.yy / dd/mm/yyyy etc
    m = re.search(r'(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})', s)
    if m:
        ds = m.group(1)
        for fmt in ("%d.%m.%Y", "%d.%m.%y", "%d-%m-%Y", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%y"):
            try:
                dt = datetime.strptime(ds, fmt)
                # нормализуем двухзначный год
                if dt.year < 100:
                    dt = dt.replace(year=2000 + dt.year)
                return dt.date().isoformat()
            except Exception:
                pass
    # 2) ищем формат "12 декабря 2023" (русские месяцы)
    m2 = re.search(r'(\d{1,2})\s+([А-Яа-яёЁ]+)\s+(\d{4})', s)
    if m2:
        d = m2.group(1)
        month_word = m2.group(2).lower()
        y = m2.group(3)
        mon = RU_MONTHS.get(month_word[:3]) or RU_MONTHS.get(month_word)
        if mon:
            try:
                dt = datetime.strptime(f"{int(d):02d}.{mon}.{y}", "%d.%m.%Y")
                return dt.date().isoformat()
            except:
                pass
    return None

# ---------- Извлечение конкретных полей ----------
def extract_contract_number(text):
    """
    Ищем номер контракта:
    - "№ SM-1712/22"
    - "Договор № 24022311"
    Возвращаем строку или None
    """
    if not text:
        return None
    # маленькая пред-очистка
    txt = text.replace('\n', ' ')
    # паттерны с символом №
    patterns = [
        r'№\s*([A-ZА-ЯЁ0-9\-\._\/]{3,})',
        r'Договор\s*№\s*([A-ZА-ЯЁ0-9\-\._\/]{3,})',
        r'ДОГОВОР\s*№\s*([A-ZА-ЯЁ0-9\-\._\/]{3,})',
        r'Contract\s*No\.?\s*([A-Z0-9\-\._\/]{3,})'
    ]
    for p in patterns:
        m = re.search(p, txt, flags=re.I)
        if m:
            val = m.group(1).strip().strip('.,;:')
            return val
    # fallback: искать короткие токены, похожие на номер (буквы-цифры с дефисом)
    m2 = re.search(r'\b([A-ZА-ЯЁ]{1,3}[-/]\d{2,6}[/]?\d{0,4})\b', txt)
    if m2:
        return m2.group(1)
    return None

def extract_dates(text):
    """
    Возвращает (date_start_iso, date_end_iso) или (None, None).
    Логика:
      - ищем контекстные маркеры "дата заключения", "срок действия" и т.д.
      - если маркеров нет — собираем все найденные даты и возвращаем первый/последний как догадку.
    """
    if not text:
        return (None, None)

    lines = [l.strip() for l in text.splitlines() if l.strip()]
    joined = "\n".join(lines)

    start_patterns = ['дата заключения', 'дата подписания', 'дата договора', 'дата составления', 'подписан']
    end_patterns = ['дата окончания', 'срок действия', 'действует до', 'по ', 'до ']

    # helper: find date in string or nearby lines
    def find_date_near(idx, window=1):
        # check this line and next `window` lines
        for j in range(idx, min(len(lines), idx + window + 1)):
            d = try_parse_date(lines[j])
            if d:
                return d
        # also check previous line
        for j in range(max(0, idx - window), idx + 1):
            d = try_parse_date(lines[j])
            if d:
                return d
        return None

    # scan lines for start / end patterns
    date_start = None
    date_end = None
    for i, ln in enumerate(lines):
        low = ln.lower()
        for pat in start_patterns:
            if pat in low and not date_start:
                d = find_date_near(i, window=2)
                if d:
                    date_start = d
        for pat in end_patterns:
            if pat in low and not date_end:
                d = find_date_near(i, window=2)
                if d:
                    date_end = d

    # если не нашли контекстно — просто найдём все даты в документе
    all_dates = []
    for ln in lines:
        # найти все подходящие числовые даты и русские
        for m in re.findall(r'\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}', ln):
            parsed = try_parse_date(m)
            if parsed:
                all_dates.append(parsed)
        # Russian month style
        m2 = re.findall(r'\d{1,2}\s+[А-Яа-яёЁ]+\s+\d{4}', ln)
        for mm in m2:
            parsed = try_parse_date(mm)
            if parsed:
                all_dates.append(parsed)

    # fallback assignment
    if not date_start and all_dates:
        date_start = all_dates[0]
    if not date_end and all_dates:
        # if only one date -> date_end = None, else last
        if len(all_dates) > 1:
            date_end = all_dates[-1]
        else:
            date_end = None

    return (date_start, date_end)

def extract_amount_and_currency(text):
    """
    Ищем сумму и валюту.
    Ищем строки с ключевыми словами 'сумма', 'стоимость', 'цена', 'итого'.
    Возвращаем (amount_float, currency_str) или (None, None)
    """
    if not text:
        return (None, None)

    text_lines = [l.strip() for l in text.splitlines() if l.strip()]

    # currency keywords
    currency_words = ['KZT', 'тенге', 'RUB', 'руб', 'руб.', 'рублей', 'USD', 'доллар', 'EUR', 'евро']
    # number pattern with possible spaces and comma/dot decimals, e.g. 3 209 315,71 or 3209315.71
    num_pat = re.compile(r'(\d{1,3}(?:[ \u00A0]\d{3})*(?:[.,]\d+)?|\d+(?:[.,]\d+)?)')

    # 1) Поиск по строкам с ключевыми словами
    for ln in text_lines:
        low = ln.lower()
        if any(k in low for k in ['сумма', 'стоимость', 'итого', 'цена', 'amount', 'total']):
            # попробуем найти число в этой строк
            m = num_pat.search(ln)
            if m:
                raw_num = m.group(1)
                amount = normalize_amount_str(raw_num)
                # найти валюту в той же строке
                cur = None
                for cw in currency_words:
                    if cw.lower() in ln.lower():
                        cur = cw
                        break
                return (amount, cur)

    # 2) Если не нашли — искать явные большие числа в документе (возможно общая сумма)
    all_nums = []
    for ln in text_lines:
        for m in num_pat.findall(ln):
            val = normalize_amount_str(m)
            if val:
                all_nums.append((val, ln))
    if all_nums:
        # возможно самая большая сумма — это сумма контракта
        all_nums_sorted = sorted(all_nums, key=lambda x: x[0], reverse=True)
        val, src_ln = all_nums_sorted[0]
        cur = None
        for cw in currency_words:
            if cw.lower() in src_ln.lower():
                cur = cw
                break
        return (val, cur)

    return (None, None)

def extract_counterparty(text):
    """
    Пытаемся извлечь название контрагента (Продавец/Покупатель/Контрагент/Поставщик)
    Используем несколько эвристик:
      - ищем 'именуемое в дальнейшем' и берём часть, предшествующую этой фразе
      - ищем строки с формой организации (ООО, ТОО, ОАО и т.п.)
      - ищем строки с ключевыми маркерами 'Покупатель', 'Продавец', 'Контрагент'
    """
    if not text:
        return None

    txt = text.replace('\r', '\n')
    lines = [l.strip() for l in txt.splitlines() if l.strip()]

    # 1) Ищем "именуемое в дальнейшем" (часто формат: "<Название>, именуемое в дальнейшем Продавец")
    for ln in lines:
        if 'именуем' in ln.lower():
            # возьмём часть до "именуем"
            m = re.search(r'(.{3,200}?)\s*,?\s*именуем', ln, flags=re.I)
            if m:
                cand = m.group(1).strip().strip(',:;')
                if len(cand) > 3:
                    return normalize_whitespace(cand)

    # 2) Ищем строки с формами организации
    forms = ['ООО', 'ОАО', 'ТОО', 'ПАО', 'ЗАО', 'ИП', 'LLP', 'LLC', 'TOO']
    for ln in lines:
        for f in forms:
            if f in ln:
                # берем всю строку или часть до запятой
                cand = ln
                # возможно название разнесено на несколько строк — соберём следующий кусок, если коротко
                if len(cand) < 6:
                    # попробуем собрать текущую и следующую
                    idx = lines.index(ln)
                    if idx + 1 < len(lines):
                        cand = (ln + " " + lines[idx + 1]).strip()
                return normalize_whitespace(cand.strip(',:;'))

    # 3) Ищем контекстные ключи "Покупатель", "Продавец", "Контрагент", "Поставщик"
    markers = ['покупатель', 'продавец', 'контрагент', 'поставщик', 'заказчик', 'исполнитель']
    for i, ln in enumerate(lines):
        low = ln.lower()
        for mk in markers:
            if mk in low:
                # попробуем взять часть строки после маркера
                # e.g. "Покупатель, в лице директора ...: ООО X"
                # если в строке есть запятая — берем остаток
                if ',' in ln:
                    # обычно имя компании позже, возьмём часть после запятой
                    after = ln.split(',', 1)[1].strip()
                    if after:
                        # возможно следующий токен — имя компании или "в лице" — тогда берем следующую строку
                        if any(tok in after.lower() for tok in ['в лице', 'действующ', 'директор']):
                            # берем следующую строку, если есть
                            if i + 1 < len(lines):
                                cand = lines[i + 1]
                                return normalize_whitespace(cand)
                        else:
                            return normalize_whitespace(after)
                # иначе — попробуем взять соседние строки (i-1..i+1)
                combined = " ".join(lines[max(0, i-1):min(len(lines), i+3)])
                # try to pull name-like substring from combined
                m = re.search(r'((?:ООО|ТОО|ОАО|ЗАО|ПАО|LLP|LLC)[\s\S]{0,80})', combined, flags=re.I)
                if m:
                    return normalize_whitespace(m.group(1))
                # fallback: return the whole combined context
                return normalize_whitespace(combined)

    return None

def extract_payment_currency(text):
    """
    Пытаемся найти валюта платежа по ключевой фразе 'валюта платежа' или 'валюта договора'
    """
    if not text:
        return None
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for ln in lines:
        low = ln.lower()
        if 'валюта платежа' in low or 'валюта договора' in low or 'валюта' == low.strip():
            # ищем код валюты
            m = re.search(r'\b(KZT|USD|EUR|RUB|руб|тенге|доллар|евро)\b', ln, flags=re.I)
            if m:
                return m.group(1)
            # возможно валюта в следующей строке
            idx = lines.index(ln)
            if idx + 1 < len(lines):
                m2 = re.search(r'\b(KZT|USD|EUR|RUB|руб|тенге|доллар|евро)\b', lines[idx+1], flags=re.I)
                if m2:
                    return m2.group(1)
    # fallback: искать частую валюту в документе
    m = re.search(r'\b(KZT|USD|EUR|RUB|руб|тенге|доллар|евро)\b', text, flags=re.I)
    if m:
        return m.group(1)
    return None

# ---------- Основной обработчик папок ----------
def process_all_outputs():
    """
    Проходим по всем подпапкам в data/output, в каждой ищем result.txt/result.docx,
    парсим и собираем итоговую таблицу.
    """
    if not os.path.exists(OUTPUT_BASE):
        print(f"[FATAL] Папка с результатами OCR не найдена: {OUTPUT_BASE}")
        return

    folders = [os.path.join(OUTPUT_BASE, d) for d in os.listdir(OUTPUT_BASE)
               if os.path.isdir(os.path.join(OUTPUT_BASE, d))]

    records = []
    for folder in sorted(folders):
        name = os.path.basename(folder)
        print(f"\n[PARSE] {name}")
        text = load_text_from_output_folder(folder)
        if not text:
            print("  [WARN] Текст не найден в папке (result.txt/result.docx). Пропускаем.")
            continue

        # извлечём основные поля
        contract_no = extract_contract_number(text)
        date_start, date_end = extract_dates(text)
        amount, currency = extract_amount_and_currency(text)
        counterparty = extract_counterparty(text)
        payment_currency = extract_payment_currency(text)

        # также соберём средний confidence если есть (парсим скобки "(conf=0.97)")
        confs = [float(m) for m in re.findall(r'\(conf=([0-9.]+)\)', text)]
        avg_conf = (sum(confs) / len(confs)) if confs else None

        rec = OrderedDict([
            ("file_folder", name),
            ("contract_number", contract_no),
            ("date_start", date_start),
            ("date_end", date_end),
            ("counterparty", counterparty),
            ("amount", amount),
            ("currency", currency),
            ("payment_currency", payment_currency),
            ("avg_confidence", avg_conf)
        ])
        records.append(rec)

        # Сохраняем подробный JSON с raw_text и найденными полями рядом в папке
        parsed_json_path = os.path.join(folder, "parsed.json")
        save_obj = {
            "file_folder": name,
            "fields": rec,
            "raw_text_preview": "\n".join(text.splitlines()[:40])
        }
        with open(parsed_json_path, "w", encoding="utf-8") as jf:
            json.dump(save_obj, jf, ensure_ascii=False, indent=2)

        print("  Найдено:")
        print(f"    contract_number: {contract_no}")
        print(f"    date_start: {date_start}, date_end: {date_end}")
        print(f"    counterparty: {counterparty}")
        print(f"    amount: {amount}  currency: {currency}")
        print(f"    payment_currency: {payment_currency}  avg_conf: {avg_conf}")

    # Сохраняем итоговую таблицу
    if records:
        df = pd.DataFrame(records)
        # записываем в Excel
        df.to_excel(RESULT_XLSX, index=False)
        print(f"\n[OK] Итог сохранён в {RESULT_XLSX} (строк: {len(records)})")
    else:
        print("\n[WARN] Нет данных для сохранения.")

if __name__ == "__main__":
    process_all_outputs()
