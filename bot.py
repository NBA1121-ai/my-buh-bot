"""
Telegram-бот для my-buh
Читает db.json из GitHub (ветка data) и отвечает на вопросы:
- остатки на складе
- деньги на счетах
- долги контрагентов
- цены товаров
- информация о сотрудниках, контрагентах, организации
- зарплатные настройки
"""

import json
import logging
import os
import re
import time
import requests
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

load_dotenv()

# ============ НАСТРОЙКИ (из .env) ============
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "")
GITHUB_REPO = os.getenv("GITHUB_REPO", "")
DATA_BRANCH = os.getenv("DATA_BRANCH", "data")
DATA_FILE = os.getenv("DATA_FILE", "db.json")
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")
# ==============================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)


# =============================================
# Загрузка данных
# =============================================

_db_cache = None
_db_cache_time = 0
DB_CACHE_TTL = 60  # кэш на 60 секунд

def fetch_db() -> dict:
    global _db_cache, _db_cache_time
    now = time.time()
    if _db_cache and (now - _db_cache_time) < DB_CACHE_TTL:
        return _db_cache
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{DATA_FILE}?ref={DATA_BRANCH}"
    headers = {"Accept": "application/vnd.github.v3.raw"}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    resp = requests.get(url, headers=headers, timeout=15)
    resp.raise_for_status()
    _db_cache = resp.json()
    _db_cache_time = now
    return _db_cache


# =============================================
# Вспомогательные: стемминг, поиск, форматирование
# =============================================

def stem(word: str) -> str:
    if len(word) <= 3:
        return word
    for suffix in ["ами", "ями", "ов", "ев", "ей", "ах", "ях", "ом", "ем",
                    "ой", "ей", "ам", "ям", "ий", "ый", "ая", "яя", "ое",
                    "ее", "ую", "юю", "ые", "ие", "ок", "ек", "ик",
                    "а", "я", "о", "е", "у", "ю", "ы", "и"]:
        if word.endswith(suffix) and len(word) - len(suffix) >= 3:
            return word[:-len(suffix)]
    return word


def fuzzy_match(query: str, name: str) -> bool:
    q_words = [stem(w) for w in query.lower().split() if len(w) > 1]
    if not q_words:
        return False
    name_lower = name.lower()
    name_stems = [stem(w) for w in name_lower.split()]
    for qw in q_words:
        if not any(qw in ns or ns in qw or qw in name_lower for ns in name_stems):
            return False
    return True


def fmt(n: float) -> str:
    if n == int(n):
        return f"{int(n):,}".replace(",", " ")
    return f"{n:,.2f}".replace(",", " ")


def clean_search(text: str, extra_stops: list = None) -> str:
    """Убирает пунктуацию и стоп-слова, возвращает чистый поисковый запрос."""
    text = re.sub(r'[?!.,;:\-—–()\"\'\«\»]', ' ', text.lower())
    stops = {
        "на", "у", "нас", "в", "по", "мне", "нам", "ещё", "еще", "мы", "я", "он", "они",
        "есть", "ли", "же", "бы", "не", "и", "а", "то", "от", "до", "за", "из", "об",
        "что", "как", "где", "кто", "чей", "это", "вот", "тот", "так", "там",
        "всё", "все", "весь", "общий", "общие", "итого", "только", "ещё", "уже",
        "скажи", "покажи", "подскажи", "расскажи", "напиши", "дай", "выведи",
        "пожалуйста", "можно", "нужно", "хочу", "знать", "узнать",
        "какой", "какая", "какие", "каков", "сколько", "много",
    }
    if extra_stops:
        for w in extra_stops:
            stops.add(w)
            stops.add(stem(w))
    words = [w for w in text.split() if w not in stops and stem(w) not in stops and len(w) > 1]
    return " ".join(words)


# =============================================
# Справочники: получение имён по ID
# =============================================

def get_account_name(db, acc_id):
    """acc_id может быть 'bank:_1' или 'cash:_1' или просто '_1'."""
    raw_id = acc_id.split(":", 1)[-1] if ":" in acc_id else acc_id
    prefix = acc_id.split(":", 1)[0] if ":" in acc_id else ""
    # Для кассовых документов — ищем сначала в кассах
    if prefix == "cash":
        for c in db.get("cashs", []):
            if c["id"] == raw_id:
                return c.get("name", raw_id)
    # Для банковских — в счетах
    for a in db.get("accounts", []):
        if a["id"] == raw_id:
            return a.get("name", raw_id)
    return raw_id or "Без счёта"

def get_warehouse_name(db, wh_id):
    for w in db.get("trade", {}).get("warehouses", []):
        if w["id"] == wh_id:
            return w.get("name", wh_id)
    return wh_id or "Без склада"

def get_contractor_name(db, c_id):
    for c in db.get("trade", {}).get("contractors", []):
        if c["id"] == c_id:
            return c.get("name", c_id)
    return c_id

def get_position_name(db, pos_id):
    for p in db.get("trade", {}).get("positions", []):
        if p["id"] == pos_id:
            return p.get("name", pos_id)
    return pos_id

def get_nom_name(db, nom_id):
    for n in db.get("trade", {}).get("nomenclature", []):
        if n["id"] == nom_id:
            return n.get("name", nom_id)
    return nom_id


# =============================================
# Вычисления
# =============================================

def calc_stock(db):
    stock = {}
    for doc in db.get("trade", {}).get("docs", []):
        dtype = doc.get("type", "")
        is_in = dtype in ("prihod", "postupleniye", "receipt", "purchase", "vozvrat_pokup")
        is_out = dtype in ("rashod", "realizaciya", "sale", "shipment", "vozvrat_post", "spisaniye")
        if not is_in and not is_out:
            continue
        wh = doc.get("warehouse", doc.get("warehouseFrom", ""))
        for row in doc.get("rows", []):
            nom_id = row.get("nomenclature", row.get("nom", ""))
            qty = float(row.get("qty", row.get("quantity", 0)))
            if is_out:
                qty = -qty
            if nom_id:
                stock.setdefault(nom_id, {})
                stock[nom_id][wh] = stock[nom_id].get(wh, 0) + qty
    return stock


def calc_balances(db):
    """Возвращает { 'bank:_1': сумма, 'cash:_1': сумма } — ключи с префиксом чтобы не путать банк и кассу."""
    balances = {}
    for doc in db.get("bankDocuments", []):
        key = "bank:" + doc.get("account", "")
        s = float(doc.get("sum", 0))
        if doc.get("type") == "payment_in":
            balances[key] = balances.get(key, 0) + s
        elif doc.get("type") == "payment_out":
            balances[key] = balances.get(key, 0) - s
    for doc in db.get("cashDocuments", []):
        key = "cash:" + doc.get("cash", doc.get("account", ""))
        s = float(doc.get("sum", 0))
        if doc.get("type") in ("cash_in", "pko"):
            balances[key] = balances.get(key, 0) + s
        elif doc.get("type") in ("cash_out", "rko"):
            balances[key] = balances.get(key, 0) - s
    return balances


def calc_contractor_debts(db):
    """
    Считает долги контрагентов по торговым + банковским документам.
    Реализация (sale/realizaciya) -> контрагент должен нам
    Оплата от покупателя (payment_in с contractor) -> контрагент заплатил
    Поступление (purchase/prihod) -> мы должны контрагенту
    Оплата поставщику (payment_out с contractor) -> мы заплатили
    Возвращает: { contractor_id: сумма } (+ = нам должны, - = мы должны)
    """
    debts = {}
    # Торговые документы
    for doc in db.get("trade", {}).get("docs", []):
        dtype = doc.get("type", "")
        c_id = doc.get("contractor", "")
        if not c_id:
            continue
        total = sum(float(r.get("total", r.get("sum", 0))) for r in doc.get("rows", []))
        if not total:
            total = float(doc.get("total", doc.get("sum", 0)))
        if dtype in ("realizaciya", "sale", "shipment"):
            debts[c_id] = debts.get(c_id, 0) + total  # нам должны
        elif dtype in ("postupleniye", "prihod", "purchase", "receipt"):
            debts[c_id] = debts.get(c_id, 0) - total  # мы должны
    # Банковские документы с контрагентом
    for doc in db.get("bankDocuments", []) + db.get("cashDocuments", []):
        c_id = doc.get("contractor", "")
        if not c_id:
            continue
        s = float(doc.get("sum", 0))
        dtype = doc.get("type", "")
        if dtype in ("payment_in", "cash_in", "pko"):
            debts[c_id] = debts.get(c_id, 0) - s  # нам заплатили
        elif dtype in ("payment_out", "cash_out", "rko"):
            debts[c_id] = debts.get(c_id, 0) + s  # мы заплатили (наш долг уменьшился)
    return debts


# =============================================
# Поиск
# =============================================

def find_nomenclature(db, query):
    return [n for n in db.get("trade", {}).get("nomenclature", [])
            if fuzzy_match(query, n.get("name", ""))]

def find_contractor(db, query):
    return [c for c in db.get("trade", {}).get("contractors", [])
            if fuzzy_match(query, c.get("name", "") + " " + c.get("full", ""))]

def find_employee(db, query):
    return [e for e in db.get("trade", {}).get("employees", [])
            if fuzzy_match(query, e.get("name", ""))]


# =============================================
# Обработчики вопросов (каждый возвращает str или None)
# =============================================

def handle_money(db, text):
    """Сколько денег на счетах / в кассе"""
    keywords = ["деньг", "денег", "денежн", "баланс", "счёт", "счет", "касс", "финанс", "бюджет"]
    if "взаиморасч" in text:
        return None  # это про долги, не про деньги
    if not any(kw in text for kw in keywords):
        return None
    balances = calc_balances(db)
    if not balances:
        return "Нет данных по движению денег."
    lines = ["Остатки по счетам:"]
    total = 0
    for acc_id, amount in balances.items():
        lines.append(f"  {get_account_name(db, acc_id)}: {fmt(amount)} сом")
        total += amount
    lines.append(f"\nИтого: {fmt(total)} сом")
    return "\n".join(lines)


def handle_debts(db, text):
    """Сколько должны контрагенты / кому мы должны"""
    kw_owe_us = ["должн", "должен", "долг", "долж", "дебитор", "задолжен", "задолженн", "взаиморасчёт", "взаиморасчет"]
    kw_we_owe = ["мы должн", "мы должен", "наш долг", "кредитор", "наша задолженн", "кому мы должн", "кому должны мы"]
    if not any(kw in text for kw in kw_owe_us):
        return None

    debts = calc_contractor_debts(db)
    we_ask_our_debt = any(kw in text for kw in kw_we_owe)

    if not debts:
        return "Нет данных по взаиморасчётам с контрагентами.\nДолги появятся после оформления реализаций/поступлений и оплат."

    # Ищем конкретного контрагента в вопросе
    search = clean_search(text, ["должн", "долг", "дебитор", "кредитор", "задолженн",
                                  "контрагент", "поставщик", "покупател", "клиент", "нам", "мы"])
    if search:
        found = find_contractor(db, search)
        if found:
            lines = []
            for c in found:
                d = debts.get(c["id"], 0)
                if d > 0:
                    lines.append(f"{c['name']}: должен нам {fmt(d)} сом")
                elif d < 0:
                    lines.append(f"{c['name']}: мы должны {fmt(-d)} сом")
                else:
                    lines.append(f"{c['name']}: взаиморасчёты закрыты (0 сом)")
            return "\n".join(lines)

    # Общая сводка
    owe_us = {}  # нам должны
    we_owe = {}  # мы должны
    for c_id, amount in debts.items():
        if amount > 0:
            owe_us[c_id] = amount
        elif amount < 0:
            we_owe[c_id] = -amount

    lines = []
    if not we_ask_our_debt and owe_us:
        lines.append("Нам должны:")
        for c_id, amount in owe_us.items():
            lines.append(f"  {get_contractor_name(db, c_id)}: {fmt(amount)} сом")
        lines.append(f"  Итого: {fmt(sum(owe_us.values()))} сом\n")
    if we_owe:
        lines.append("Мы должны:")
        for c_id, amount in we_owe.items():
            lines.append(f"  {get_contractor_name(db, c_id)}: {fmt(amount)} сом")
        lines.append(f"  Итого: {fmt(sum(we_owe.values()))} сом")
    if not we_ask_our_debt and not owe_us and not we_owe:
        return "Все взаиморасчёты закрыты, долгов нет."
    if not lines:
        return "По данному направлению долгов нет."
    return "\n".join(lines)


def handle_price(db, text):
    """Сколько стоит товар / цена товара"""
    keywords = ["стоит", "стоимость", "цен", "прайс", "расценк"]
    if not any(kw in text for kw in keywords):
        return None
    search = clean_search(text, ["стоит", "стоимость", "цена", "прайс", "расценка", "товар"])
    if not search:
        # Показываем прайс-лист
        noms = db.get("trade", {}).get("nomenclature", [])
        lines = [f"Прайс-лист ({len(noms)} позиций):"]
        for n in noms:
            if n.get("price", 0) > 0:
                lines.append(f"  {n['name']}: {fmt(n['price'])} сом/{n.get('unit', 'шт')}")
        return "\n".join(lines)
    noms = find_nomenclature(db, search)
    if not noms:
        return f"Товар '{search}' не найден."
    lines = []
    for n in noms:
        price = n.get("price", 0)
        cost = n.get("cost", 0)
        line = f"{n['name']}: {fmt(price)} сом/{n.get('unit', 'шт')}"
        if cost > 0:
            line += f" (себестоимость: {fmt(cost)} сом)"
        lines.append(line)
    return "\n".join(lines)


def handle_warehouses(db, text):
    """Список складов"""
    if not re.search(r'\bсклад[ыа]?\b', text):
        return None
    if any(kw in text for kw in ["сколько", "остат", "наличи", "есть", "что на"]):
        return None  # это вопрос про остатки, не про список складов
    warehouses = db.get("trade", {}).get("warehouses", [])
    if not warehouses:
        return "Склады не найдены."
    lines = ["Склады:"]
    for w in warehouses:
        lines.append(f"  {w['name']} ({w.get('kind', '')}), ответственный: {w.get('responsible', '-')}")
    return "\n".join(lines)


def handle_stock(db, text):
    """Остатки на складе"""
    keywords = ["склад", "остат", "наличи", "сколько", "есть ли", "имеется"]
    if not any(kw in text for kw in keywords):
        return None
    search = clean_search(text, [
        "склад", "складе", "складу", "складах", "остатки", "остаток", "остатков",
        "наличие", "наличии", "имеется", "имеются", "товар", "штук", "штуки", "кг", "шт",
    ])

    if search:
        noms = find_nomenclature(db, search)
        if not noms:
            all_noms = db.get("trade", {}).get("nomenclature", [])
            suggestions = [n["name"] for n in all_noms
                           if any(w in n["name"].lower() for w in search.split() if len(w) > 2)]
            if suggestions:
                return f"Товар '{search}' не найден. Может вы имели в виду:\n" + "\n".join(f"  - {s}" for s in suggestions[:10])
            return f"Товар '{search}' не найден в номенклатуре."
    else:
        noms = db.get("trade", {}).get("nomenclature", [])

    stock = calc_stock(db)
    has_any = any(stock.get(n["id"]) for n in noms)

    if not has_any:
        lines = ["На складе пока нет движений (приход/расход не оформлялся).\n"]
        show = [n for n in noms if n.get("kind") in ("Товар", "Материал", "Продукция", "Тара")]
        lines.append(f"Номенклатура ({len(show)} позиций):")
        for n in show:
            lines.append(f"  {n['name']} [{n.get('kind','')}] — 0 {n.get('unit', 'шт')}")
        lines.append("\nОформите поступление товара на сайте, тогда появятся остатки.")
        return "\n".join(lines)

    lines = ["Остатки на складе:"]
    for n in noms:
        ns = stock.get(n["id"], {})
        total = sum(ns.values())
        if total != 0:
            lines.append(f"  {n['name']}: {fmt(total)} {n.get('unit', 'шт')}")
            for wh_id, qty in ns.items():
                lines.append(f"      {get_warehouse_name(db, wh_id)}: {fmt(qty)} {n.get('unit', 'шт')}")
        elif search:
            lines.append(f"  {n['name']}: 0 {n.get('unit', 'шт')} (нет движений)")
    return "\n".join(lines)


def handle_goods(db, text):
    """Список товаров / номенклатура"""
    keywords = ["товар", "номенклатур", "ассортимент", "что продаём", "что продаем", "каталог"]
    if not any(kw in text for kw in keywords):
        return None
    noms = db.get("trade", {}).get("nomenclature", [])
    if not noms:
        return "Номенклатура пуста."
    lines = [f"Номенклатура ({len(noms)} позиций):"]
    for n in noms:
        price = fmt(n.get("price", 0))
        lines.append(f"  {n['name']} [{n.get('kind', '')}] - {price} сом/{n.get('unit', 'шт')}")
    return "\n".join(lines)


def handle_contractors(db, text):
    """Контрагенты"""
    keywords = ["контрагент", "поставщик", "покупател", "клиент", "партнёр", "партнер"]
    if not any(kw in text for kw in keywords):
        return None
    search = clean_search(text, keywords + ["список", "информац", "данные", "инн"])
    contractors = db.get("trade", {}).get("contractors", [])
    if search:
        contractors = find_contractor(db, search)
    if not contractors:
        return "Контрагенты не найдены."

    # Если нашли конкретного — подробная инфо
    if len(contractors) <= 3 and search:
        lines = []
        for c in contractors:
            lines.append(f"{c.get('name', '')}")
            if c.get("full"): lines.append(f"  Полное: {c['full']}")
            if c.get("inn"): lines.append(f"  ИНН: {c['inn']}")
            if c.get("kind"): lines.append(f"  Тип: {c['kind']}")
            if c.get("address"): lines.append(f"  Адрес: {c['address']}")
            if c.get("phone"): lines.append(f"  Тел: {c['phone']}")
            if c.get("email"): lines.append(f"  Email: {c['email']}")
            # Показать договоры
            contracts = [d for d in db.get("trade", {}).get("contracts", []) if d.get("contractor") == c["id"]]
            if contracts:
                lines.append("  Договоры:")
                for d in contracts:
                    lines.append(f"    №{d.get('number', '?')} от {d.get('date', '?')} - {d.get('name', '')} ({d.get('kind', '')})")
        return "\n".join(lines)

    # Список
    lines = [f"Контрагенты ({len(contractors)}):"]
    for c in contractors:
        lines.append(f"  {c['name']} (ИНН: {c.get('inn', '-')})")
    return "\n".join(lines)


def handle_employees(db, text):
    """Сотрудники"""
    keywords = ["сотрудник", "работник", "персонал", "кадр", "штат", "табельн"]
    # Также ловим "инфо о [имя]" — проверяем совпадение с именами сотрудников
    direct_match = False
    if not any(kw in text for kw in keywords):
        # Проверяем, не спрашивают ли про конкретного сотрудника по имени/фамилии
        employees = db.get("trade", {}).get("employees", [])
        search = clean_search(text, ["информация", "инфо", "данные", "расскажи", "про", "кто", "такой", "такая"])
        if search and any(fuzzy_match(search, e.get("name", "")) for e in employees):
            direct_match = True
        else:
            return None
    search = clean_search(text, keywords + ["список", "информац", "данные"])
    employees = db.get("trade", {}).get("employees", [])

    if search:
        found = find_employee(db, search)
        if found:
            employees = found
        # Если не нашли — показываем всех

    if not employees:
        return "Список сотрудников пуст."

    # Если нашли конкретного — подробно
    if len(employees) <= 2 and search:
        lines = []
        for e in employees:
            lines.append(f"{e.get('name', '')}")
            lines.append(f"  Таб. №: {e.get('tabNo', '-')}")
            lines.append(f"  Должность: {get_position_name(db, e.get('position', ''))}")
            if e.get("phone"): lines.append(f"  Тел: {e['phone']}")
            if e.get("birth"): lines.append(f"  Дата рождения: {e['birth']}")
            if e.get("inn"): lines.append(f"  ИНН: {e['inn']}")
            if e.get("address"): lines.append(f"  Адрес: {e['address']}")
            salary = e.get("salary", 0)
            if salary: lines.append(f"  Оклад: {fmt(salary)} сом")
        return "\n".join(lines)

    # Общий список
    lines = [f"Сотрудники ({len(employees)}):"]
    for e in employees:
        pos = get_position_name(db, e.get("position", ""))
        lines.append(f"  {e['name']} - {pos}")
    return "\n".join(lines)


def handle_org(db, text):
    """Информация об организации"""
    keywords = ["организац", "компани", "фирм", "реквизит", "наша компан", "юрлиц"]
    if not any(kw in text for kw in keywords):
        return None
    org = db.get("trade", {}).get("org", {})
    if not org or not org.get("name"):
        return "Данные организации не заполнены."
    lines = ["Организация:"]
    if org.get("name"): lines.append(f"  Название: {org['name']}")
    if org.get("inn"): lines.append(f"  ИНН: {org['inn']}")
    if org.get("address"): lines.append(f"  Адрес: {org['address']}")
    if org.get("phone"): lines.append(f"  Тел: {org['phone']}")
    if org.get("director"): lines.append(f"  Руководитель: {org['director']}")
    if org.get("accountant"): lines.append(f"  Бухгалтер: {org['accountant']}")
    if org.get("okpo"): lines.append(f"  ОКПО: {org['okpo']}")
    return "\n".join(lines)


def handle_payroll_info(db, text):
    """Зарплатные ставки и налоги"""
    keywords = ["зарплат", "оклад", "налог", "отчислен", "ставк", "соцфонд", "подоходн"]
    if not any(kw in text for kw in keywords):
        return None
    payroll = db.get("trade", {}).get("payroll", {})
    if not payroll:
        return "Зарплатные настройки не заданы."
    lines = ["Зарплатные настройки:"]
    lines.append(f"  Подоходный налог: {payroll.get('incomeTax', 0)}%")
    lines.append(f"  Соцфонд (сотрудник): {payroll.get('sfEmployee', 0)}%")
    lines.append(f"    в т.ч. ПФ: {payroll.get('sfEmployeePF', 0)}%, ГНПФ: {payroll.get('sfEmployeeGNPF', 0)}%")
    lines.append(f"  Соцфонд (работодатель): {payroll.get('sfEmployer', 0)}%")
    lines.append(f"    в т.ч. ПФ: {payroll.get('sfEmployerPF', 0)}%, ФОМС: {payroll.get('sfEmployerFOMS', 0)}%, ФОТ: {payroll.get('sfEmployerFOT', 0)}%")
    lines.append(f"  Стандартный вычет: {fmt(payroll.get('stdDeduction', 0))} сом")
    lines.append(f"  Вычет на иждивенца: {fmt(payroll.get('dependentDeduction', 0))} сом")
    lines.append(f"  Мин. зарплата: {fmt(payroll.get('minSalary', 0))} сом")
    return "\n".join(lines)


def handle_contracts(db, text):
    """Договоры"""
    keywords = ["договор", "контракт"]
    if not any(kw in text for kw in keywords):
        return None
    contracts = db.get("trade", {}).get("contracts", [])
    if not contracts:
        return "Договоры не найдены."
    lines = [f"Договоры ({len(contracts)}):"]
    for d in contracts:
        c_name = get_contractor_name(db, d.get("contractor", ""))
        lines.append(f"  №{d.get('number', '?')} от {d.get('date', '?')} - {d.get('name', '')}")
        lines.append(f"    Контрагент: {c_name}, тип: {d.get('kind', '')}")
    return "\n".join(lines)


def handle_summary(db, text):
    """Общая сводка / отчёт"""
    keywords = ["сводк", "отчёт", "отчет", "итог", "обзор", "дашборд", "dashboard", "статус"]
    if not any(kw in text for kw in keywords):
        return None
    lines = ["Сводка по базе:\n"]

    # Деньги
    balances = calc_balances(db)
    total_money = sum(balances.values()) if balances else 0
    lines.append(f"Деньги: {fmt(total_money)} сом")
    for acc_id, amount in balances.items():
        lines.append(f"  {get_account_name(db, acc_id)}: {fmt(amount)} сом")

    # Склад
    stock = calc_stock(db)
    stock_count = sum(1 for nid in stock if sum(stock[nid].values()) > 0)
    lines.append(f"\nТовары на складе: {stock_count} позиций с остатками")

    # Контрагенты
    contractors = db.get("trade", {}).get("contractors", [])
    lines.append(f"Контрагентов: {len(contractors)}")

    # Долги
    debts = calc_contractor_debts(db)
    owe_us = sum(v for v in debts.values() if v > 0)
    we_owe = sum(-v for v in debts.values() if v < 0)
    if owe_us: lines.append(f"Нам должны: {fmt(owe_us)} сом")
    if we_owe: lines.append(f"Мы должны: {fmt(we_owe)} сом")

    # Сотрудники
    employees = db.get("trade", {}).get("employees", [])
    lines.append(f"Сотрудников: {len(employees)}")

    # Номенклатура
    noms = db.get("trade", {}).get("nomenclature", [])
    lines.append(f"Номенклатура: {len(noms)} позиций")

    return "\n".join(lines)


# =============================================
# Главная маршрутизация
# =============================================

HANDLERS = [
    handle_summary,      # "сводка", "отчёт" — первый, чтобы не перехватывался
    handle_money,        # "деньги", "счёт", "баланс"
    handle_debts,        # "должны", "долг"
    handle_price,        # "стоит", "цена"
    handle_payroll_info, # "зарплата", "налог"
    handle_org,          # "организация", "реквизиты"
    handle_contracts,    # "договор"
    handle_warehouses,   # "склады" (без вопроса про остатки)
    handle_stock,        # "сколько", "склад", "остатки"
    handle_goods,        # "товары", "номенклатура"
    handle_contractors,  # "контрагенты", "поставщики"
    handle_employees,    # "сотрудники", "кадры"
]

HELP_TEXT = (
    "Я могу ответить на:\n\n"
    "  Сколько [товар] на складе?\n"
    "  Сколько денег на счетах?\n"
    "  Сколько стоит [товар]?\n"
    "  Кто нам должен? / Наши долги\n"
    "  Сводка / Отчёт\n"
    "  Список товаров\n"
    "  Контрагенты / инфо о [контрагент]\n"
    "  Сотрудники / инфо о [сотрудник]\n"
    "  Договоры\n"
    "  Реквизиты организации\n"
    "  Зарплатные ставки\n"
    "  Склады\n\n"
    "Просто напишите вопрос!"
)


def process_question(text: str) -> str:
    text_lower = text.lower().strip()
    try:
        db = fetch_db()
    except Exception as e:
        log.error(f"Ошибка загрузки данных: {e}")
        return "Не удалось загрузить данные с сервера. Попробуйте позже."

    for handler in HANDLERS:
        result = handler(db, text_lower)
        if result is not None:
            return result

    return "Не совсем понял вопрос.\n\n" + HELP_TEXT


# =============================================
# Telegram
# =============================================

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Я бот бухгалтерии.\n\n" + HELP_TEXT)

async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP_TEXT)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if not text:
        return
    log.info(f"Вопрос от {update.effective_user.first_name}: {text}")
    answer = process_question(text)
    # Telegram лимит 4096 символов
    if len(answer) > 4000:
        answer = answer[:4000] + "\n\n... (обрезано)"
    await update.message.reply_text(answer)
    log.info(f"Ответ: {answer[:100]}...")


def main():
    if not TELEGRAM_TOKEN:
        log.error("TELEGRAM_TOKEN не задан в .env!")
        return
    if not GITHUB_REPO:
        log.error("GITHUB_REPO не задан в .env!")
        return
    log.info("Запуск бота...")
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    log.info("Бот запущен! Ожидаю сообщения...")
    app.run_polling()


if __name__ == "__main__":
    main()
