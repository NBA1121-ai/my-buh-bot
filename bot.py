"""
Telegram-бот для my-buh
Читает db.json из GitHub (ветка data) и отвечает на вопросы
о складе, финансах, контрагентах, сотрудниках и т.д.
"""

import json
import logging
import os
import re
import time
import requests
from dotenv import load_dotenv
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    filters, ContextTypes,
)
from telegram.constants import ParseMode, ChatAction

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "")
GITHUB_REPO = os.getenv("GITHUB_REPO", "")
DATA_BRANCH = os.getenv("DATA_BRANCH", "data")
DATA_FILE = os.getenv("DATA_FILE", "db.json")
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════
#  Загрузка данных с кэшированием
# ══════════════════════════════════════════════

_db_cache = None
_db_cache_time = 0
DB_CACHE_TTL = 60

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


# ══════════════════════════════════════════════
#  Стемминг, поиск, форматирование
# ══════════════════════════════════════════════

def stem(word: str) -> str:
    if len(word) <= 3:
        return word
    for suf in ["ами","ями","ов","ев","ей","ах","ях","ом","ем",
                "ой","ей","ам","ям","ий","ый","ая","яя","ое",
                "ее","ую","юю","ые","ие","ок","ек","ик",
                "а","я","о","е","у","ю","ы","и"]:
        if word.endswith(suf) and len(word) - len(suf) >= 3:
            return word[:-len(suf)]
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


_STOP_WORDS = {
    "на","у","нас","в","по","мне","нам","ещё","еще","мы","я","он","они",
    "есть","ли","же","бы","не","и","а","то","от","до","за","из","об",
    "что","как","где","кто","чей","это","вот","тот","так","там",
    "всё","все","весь","общий","общие","итого","только","уже",
    "скажи","покажи","подскажи","расскажи","напиши","дай","выведи",
    "пожалуйста","можно","нужно","хочу","знать","узнать",
    "какой","какая","какие","каков","сколько","много",
}

def clean_search(text: str, extra_stops: list = None) -> str:
    text = re.sub(r'[?!.,;:\-—–()\"\'\«\»]', ' ', text.lower())
    stops = set(_STOP_WORDS)
    if extra_stops:
        for w in extra_stops:
            stops.add(w)
            stops.add(stem(w))
    return " ".join(w for w in text.split() if w not in stops and stem(w) not in stops and len(w) > 1)


# ══════════════════════════════════════════════
#  Справочники → имена по ID
# ══════════════════════════════════════════════

def _lookup(items, item_id, field="name"):
    for item in items:
        if item.get("id") == item_id:
            return item.get(field, item_id)
    return item_id

def get_account_name(db, acc_id):
    raw = acc_id.split(":", 1)[-1] if ":" in acc_id else acc_id
    prefix = acc_id.split(":", 1)[0] if ":" in acc_id else ""
    if prefix == "cash":
        name = _lookup(db.get("cashs", []), raw)
        if name != raw:
            return name
    name = _lookup(db.get("accounts", []), raw)
    return name if name != raw else (raw or "Без счёта")

def get_warehouse_name(db, wh_id):
    return _lookup(db.get("trade", {}).get("warehouses", []), wh_id) if wh_id else "Без склада"

def get_contractor_name(db, c_id):
    return _lookup(db.get("trade", {}).get("contractors", []), c_id)

def get_position_name(db, pos_id):
    return _lookup(db.get("trade", {}).get("positions", []), pos_id)


# ══════════════════════════════════════════════
#  Вычисления
# ══════════════════════════════════════════════

_DOC_IN  = {"prihod","postupleniye","receipt","purchase","vozvrat_pokup"}
_DOC_OUT = {"rashod","realizaciya","sale","shipment","vozvrat_post","spisaniye"}

def calc_stock(db):
    stock = {}
    for doc in db.get("trade", {}).get("docs", []):
        dtype = doc.get("type", "")
        is_in, is_out = dtype in _DOC_IN, dtype in _DOC_OUT
        if not is_in and not is_out:
            continue
        wh = doc.get("warehouse", doc.get("warehouseFrom", ""))
        for row in doc.get("rows", []):
            nid = row.get("nomenclature", row.get("nom", ""))
            qty = float(row.get("qty", row.get("quantity", 0)))
            if is_out:
                qty = -qty
            if nid:
                stock.setdefault(nid, {})
                stock[nid][wh] = stock[nid].get(wh, 0) + qty
    return stock


def calc_balances(db):
    bal = {}
    for doc in db.get("bankDocuments", []):
        key = "bank:" + doc.get("account", "")
        s = float(doc.get("sum", 0))
        if doc["type"] == "payment_in":    bal[key] = bal.get(key, 0) + s
        elif doc["type"] == "payment_out": bal[key] = bal.get(key, 0) - s
    for doc in db.get("cashDocuments", []):
        key = "cash:" + doc.get("cash", doc.get("account", ""))
        s = float(doc.get("sum", 0))
        if doc["type"] in ("cash_in", "pko"):    bal[key] = bal.get(key, 0) + s
        elif doc["type"] in ("cash_out", "rko"): bal[key] = bal.get(key, 0) - s
    return bal


def calc_debts(db):
    debts = {}
    for doc in db.get("trade", {}).get("docs", []):
        c_id = doc.get("contractor", "")
        if not c_id:
            continue
        total = sum(float(r.get("total", r.get("sum", 0))) for r in doc.get("rows", []))
        if not total:
            total = float(doc.get("total", doc.get("sum", 0)))
        dtype = doc.get("type", "")
        if dtype in ("realizaciya", "sale", "shipment"):
            debts[c_id] = debts.get(c_id, 0) + total
        elif dtype in ("postupleniye", "prihod", "purchase", "receipt"):
            debts[c_id] = debts.get(c_id, 0) - total
    for doc in db.get("bankDocuments", []) + db.get("cashDocuments", []):
        c_id = doc.get("contractor", "")
        if not c_id:
            continue
        s = float(doc.get("sum", 0))
        dtype = doc.get("type", "")
        if dtype in ("payment_in", "cash_in", "pko"):
            debts[c_id] = debts.get(c_id, 0) - s
        elif dtype in ("payment_out", "cash_out", "rko"):
            debts[c_id] = debts.get(c_id, 0) + s
    return debts


# ══════════════════════════════════════════════
#  Поиск
# ══════════════════════════════════════════════

def find_nomenclature(db, q):
    return [n for n in db.get("trade",{}).get("nomenclature",[]) if fuzzy_match(q, n.get("name",""))]

def find_contractor(db, q):
    return [c for c in db.get("trade",{}).get("contractors",[]) if fuzzy_match(q, c.get("name","")+" "+c.get("full",""))]

def find_employee(db, q):
    return [e for e in db.get("trade",{}).get("employees",[]) if fuzzy_match(q, e.get("name",""))]


# ══════════════════════════════════════════════
#  Форматирование ответов (Telegram Markdown)
# ══════════════════════════════════════════════

def _esc(text: str) -> str:
    """Экранирует спецсимволы MarkdownV2."""
    return re.sub(r'([_*\[\]()~`>#+\-=|{}.!\\])', r'\\\1', str(text))

def _bold(text: str) -> str:
    return f"*{_esc(text)}*"

def _line(label: str, value: str) -> str:
    return f"  {_esc(label)}: {_bold(value)}"


# ══════════════════════════════════════════════
#  Обработчики вопросов
# ══════════════════════════════════════════════

def handle_summary(db, text):
    kw = ["сводк","отчёт","отчет","итог","обзор","дашборд","dashboard","статус"]
    if not any(k in text for k in kw):
        return None

    bal = calc_balances(db)
    total_money = sum(bal.values())
    stock = calc_stock(db)
    stock_items = sum(1 for nid in stock if sum(stock[nid].values()) > 0)
    debts = calc_debts(db)
    owe_us = sum(v for v in debts.values() if v > 0)
    we_owe = sum(-v for v in debts.values() if v < 0)
    contractors = db.get("trade",{}).get("contractors",[])
    employees = db.get("trade",{}).get("employees",[])
    noms = db.get("trade",{}).get("nomenclature",[])

    lines = [f"{'='*28}", _bold("СВОДКА"), f"{'='*28}\n"]
    lines.append(f"💰 {_bold('Финансы')}")
    for acc_id, amount in bal.items():
        lines.append(f"  {_esc(get_account_name(db, acc_id))}: {_bold(fmt(amount) + ' сом')}")
    lines.append(f"  {'─'*20}")
    lines.append(f"  Итого: {_bold(fmt(total_money) + ' сом')}\n")

    lines.append(f"📦 {_bold('Склад')}: {_esc(str(stock_items))} позиций с остатками")
    lines.append(f"📋 {_bold('Номенклатура')}: {_esc(str(len(noms)))} позиций")
    lines.append(f"🤝 {_bold('Контрагенты')}: {_esc(str(len(contractors)))}")
    lines.append(f"👥 {_bold('Сотрудники')}: {_esc(str(len(employees)))}\n")

    if owe_us or we_owe:
        lines.append(f"📊 {_bold('Взаиморасчёты')}")
        if owe_us:
            lines.append(f"  Нам должны: {_bold(fmt(owe_us) + ' сом')}")
        if we_owe:
            lines.append(f"  Мы должны: {_bold(fmt(we_owe) + ' сом')}")

    return "\n".join(lines)


def handle_money(db, text):
    kw = ["деньг","денег","денежн","баланс","счёт","счет","касс","финанс","бюджет"]
    if "взаиморасч" in text:
        return None
    if not any(k in text for k in kw):
        return None
    bal = calc_balances(db)
    if not bal:
        return _bold("Нет данных по движению денег\\.")

    lines = [f"💰 {_bold('Остатки по счетам')}\n"]
    bank_total, cash_total = 0, 0
    banks, cashes = [], []
    for acc_id, amount in bal.items():
        name = get_account_name(db, acc_id)
        entry = f"  {_esc(name)}: {_bold(fmt(amount) + ' сом')}"
        if acc_id.startswith("bank:"):
            banks.append(entry)
            bank_total += amount
        else:
            cashes.append(entry)
            cash_total += amount
    if banks:
        lines.append(f"🏦 {_esc('Банковские счета')}:")
        lines.extend(banks)
        lines.append("")
    if cashes:
        lines.append(f"💵 {_esc('Кассы')}:")
        lines.extend(cashes)
        lines.append("")
    lines.append(f"{'─'*24}")
    lines.append(f"Итого: {_bold(fmt(bank_total + cash_total) + ' сом')}")
    return "\n".join(lines)


def handle_debts(db, text):
    kw_all = ["должн","должен","долг","долж","дебитор","задолжен","задолженн","взаиморасчёт","взаиморасчет"]
    kw_we = ["мы должн","мы должен","наш долг","кредитор","наша задолженн","кому мы должн","кому должны мы"]
    if not any(k in text for k in kw_all):
        return None

    debts = calc_debts(db)
    if not debts:
        return f"📊 {_bold('Взаиморасчёты')}\n\nДанных нет\\. Долги появятся после оформления документов\\."

    search = clean_search(text, ["должн","долг","дебитор","кредитор","задолженн","контрагент","поставщик","покупател","клиент","нам","мы"])
    if search:
        found = find_contractor(db, search)
        if found:
            lines = [f"📊 {_bold('Взаиморасчёты')}\n"]
            for c in found:
                d = debts.get(c["id"], 0)
                if d > 0:
                    lines.append(f"🔴 {_esc(c['name'])}: должен нам {_bold(fmt(d) + ' сом')}")
                elif d < 0:
                    lines.append(f"🟡 {_esc(c['name'])}: мы должны {_bold(fmt(-d) + ' сом')}")
                else:
                    lines.append(f"🟢 {_esc(c['name'])}: расчёты закрыты")
            return "\n".join(lines)

    owe_us, we_owe = {}, {}
    for c_id, amount in debts.items():
        if amount > 0:   owe_us[c_id] = amount
        elif amount < 0:  we_owe[c_id] = -amount

    we_ask_ours = any(k in text for k in kw_we)
    lines = [f"📊 {_bold('Взаиморасчёты')}\n"]

    if not we_ask_ours and owe_us:
        lines.append(f"🔴 {_esc('Нам должны')}:")
        for c_id, amt in sorted(owe_us.items(), key=lambda x: -x[1]):
            lines.append(f"  {_esc(get_contractor_name(db, c_id))}: {_bold(fmt(amt) + ' сом')}")
        lines.append(f"  Итого: {_bold(fmt(sum(owe_us.values())) + ' сом')}\n")
    if we_owe:
        lines.append(f"🟡 {_esc('Мы должны')}:")
        for c_id, amt in sorted(we_owe.items(), key=lambda x: -x[1]):
            lines.append(f"  {_esc(get_contractor_name(db, c_id))}: {_bold(fmt(amt) + ' сом')}")
        lines.append(f"  Итого: {_bold(fmt(sum(we_owe.values())) + ' сом')}")

    if not owe_us and not we_owe:
        lines.append("🟢 Все расчёты закрыты, долгов нет\\.")
    return "\n".join(lines)


def handle_price(db, text):
    kw = ["стоит","стоимость","цен","прайс","расценк"]
    if not any(k in text for k in kw):
        return None
    search = clean_search(text, ["стоит","стоимость","цена","прайс","расценка","товар"])
    noms = db.get("trade",{}).get("nomenclature",[])
    if search:
        noms = find_nomenclature(db, search)
        if not noms:
            return f"Товар '{_esc(search)}' не найден\\."

    lines = [f"🏷 {_bold('Прайс-лист')}\n"]
    for n in noms:
        p = n.get("price", 0)
        c = n.get("cost", 0)
        if p > 0 or search:
            line = f"  {_esc(n['name'])}: {_bold(fmt(p) + ' сом')}/{_esc(n.get('unit','шт'))}"
            if c > 0 and search:
                line += f" \\(себест\\. {_esc(fmt(c))} сом\\)"
            lines.append(line)
    return "\n".join(lines)


def handle_warehouses(db, text):
    if not re.search(r'\bсклад[ыа]?\b', text):
        return None
    if any(k in text for k in ["сколько","остат","наличи","есть","что на"]):
        return None
    whs = db.get("trade",{}).get("warehouses",[])
    if not whs:
        return "Склады не найдены\\."
    lines = [f"🏭 {_bold('Склады')}\n"]
    for w in whs:
        lines.append(f"  📍 {_bold(w['name'])}")
        lines.append(f"     Тип: {_esc(w.get('kind',''))}")
        lines.append(f"     Ответственный: {_esc(w.get('responsible','-'))}")
    return "\n".join(lines)


def handle_stock(db, text):
    kw = ["склад","остат","наличи","сколько","есть ли","имеется"]
    if not any(k in text for k in kw):
        return None
    search = clean_search(text, [
        "склад","складе","складу","складах","остатки","остаток","остатков",
        "наличие","наличии","имеется","имеются","товар","штук","штуки","кг","шт",
    ])

    noms = db.get("trade",{}).get("nomenclature",[])
    if search:
        found = find_nomenclature(db, search)
        if not found:
            suggestions = [n["name"] for n in noms if any(w in n["name"].lower() for w in search.split() if len(w)>2)]
            if suggestions:
                return f"Товар '{_esc(search)}' не найден\\. Может имелось в виду:\n" + "\n".join(f"  • {_esc(s)}" for s in suggestions[:10])
            return f"Товар '{_esc(search)}' не найден\\."
        noms = found

    stock = calc_stock(db)
    has_any = any(stock.get(n["id"]) for n in noms)

    if not has_any:
        lines = [f"📦 {_bold('Остатки на складе')}\n"]
        show = [n for n in noms if n.get("kind") in ("Товар","Материал","Продукция","Тара")]
        if show:
            for n in show:
                lines.append(f"  {_esc(n['name'])}: {_bold('0')} {_esc(n.get('unit','шт'))}")
            lines.append(f"\n⚠️ {_esc('Движений пока нет. Оформите поступление товара на сайте.')}")
        return "\n".join(lines)

    lines = [f"📦 {_bold('Остатки на складе')}\n"]
    for n in noms:
        ns = stock.get(n["id"], {})
        total = sum(ns.values())
        if total > 0:
            lines.append(f"  ✅ {_esc(n['name'])}: {_bold(fmt(total))} {_esc(n.get('unit','шт'))}")
            if len(ns) > 1:
                for wh_id, qty in ns.items():
                    lines.append(f"       {_esc(get_warehouse_name(db, wh_id))}: {_esc(fmt(qty))}")
        elif total < 0:
            lines.append(f"  ⚠️ {_esc(n['name'])}: {_bold(fmt(total))} {_esc(n.get('unit','шт'))} \\(минус\\!\\)")
        elif search:
            lines.append(f"  ◻️ {_esc(n['name'])}: {_bold('0')} {_esc(n.get('unit','шт'))}")
    return "\n".join(lines)


def handle_goods(db, text):
    kw = ["товар","номенклатур","ассортимент","что продаём","что продаем","каталог"]
    if not any(k in text for k in kw):
        return None
    noms = db.get("trade",{}).get("nomenclature",[])
    if not noms:
        return "Номенклатура пуста\\."

    lines = [f"📋 {_bold('Номенклатура')} \\({_esc(str(len(noms)))} позиций\\)\n"]
    by_kind = {}
    for n in noms:
        by_kind.setdefault(n.get("kind","Прочее"), []).append(n)
    for kind, items in by_kind.items():
        lines.append(f"  {_bold(kind)}:")
        for n in items:
            p = fmt(n.get("price",0))
            lines.append(f"    {_esc(n['name'])} — {_bold(p + ' сом')}/{_esc(n.get('unit','шт'))}")
        lines.append("")
    return "\n".join(lines)


def handle_contractors(db, text):
    kw = ["контрагент","поставщик","покупател","клиент","партнёр","партнер"]
    if not any(k in text for k in kw):
        return None
    search = clean_search(text, kw + ["список","информац","данные","инн"])
    contractors = db.get("trade",{}).get("contractors",[])
    if search:
        contractors = find_contractor(db, search)
    if not contractors:
        return "Контрагенты не найдены\\."

    if len(contractors) <= 3 and search:
        lines = []
        for c in contractors:
            lines.append(f"🤝 {_bold(c.get('name',''))}\n")
            if c.get("full"):    lines.append(_line("Полное название", c["full"]))
            if c.get("inn"):     lines.append(_line("ИНН", c["inn"]))
            if c.get("kind"):    lines.append(_line("Тип", c["kind"]))
            if c.get("address"): lines.append(_line("Адрес", c["address"]))
            if c.get("phone"):   lines.append(_line("Телефон", c["phone"]))
            contracts = [d for d in db.get("trade",{}).get("contracts",[]) if d.get("contractor")==c["id"]]
            if contracts:
                lines.append(f"\n  📝 {_esc('Договоры')}:")
                for d in contracts:
                    lines.append(f"    №{_esc(d.get('number','?'))} от {_esc(d.get('date','?'))} — {_esc(d.get('name',''))}")
        return "\n".join(lines)

    lines = [f"🤝 {_bold('Контрагенты')} \\({_esc(str(len(contractors)))}\\)\n"]
    for c in contractors:
        lines.append(f"  • {_esc(c['name'])}  \\|  ИНН: {_esc(c.get('inn','-'))}")
    return "\n".join(lines)


def handle_employees(db, text):
    kw = ["сотрудник","работник","персонал","кадр","штат","табельн"]
    if not any(k in text for k in kw):
        employees = db.get("trade",{}).get("employees",[])
        search = clean_search(text, ["информация","инфо","данные","расскажи","про","кто","такой","такая"])
        if not (search and any(fuzzy_match(search, e.get("name","")) for e in employees)):
            return None

    search = clean_search(text, kw + ["список","информац","данные"])
    employees = db.get("trade",{}).get("employees",[])
    if search:
        found = find_employee(db, search)
        if found:
            employees = found

    if not employees:
        return "Список сотрудников пуст\\."

    if len(employees) <= 2 and search:
        lines = []
        for e in employees:
            lines.append(f"👤 {_bold(e.get('name',''))}\n")
            lines.append(_line("Таб. №", e.get("tabNo","-")))
            lines.append(_line("Должность", get_position_name(db, e.get("position",""))))
            if e.get("phone"):    lines.append(_line("Телефон", e["phone"]))
            if e.get("birth"):    lines.append(_line("Дата рождения", e["birth"]))
            if e.get("inn"):      lines.append(_line("ИНН", e["inn"]))
            if e.get("address"):  lines.append(_line("Адрес", e["address"]))
            salary = e.get("salary", 0)
            if salary:            lines.append(_line("Оклад", f"{fmt(salary)} сом"))
            if e.get("status"):   lines.append(_line("Статус", e["status"]))
        return "\n".join(lines)

    lines = [f"👥 {_bold('Сотрудники')} \\({_esc(str(len(employees)))}\\)\n"]
    for e in employees:
        pos = get_position_name(db, e.get("position",""))
        lines.append(f"  • {_esc(e['name'])} — {_esc(pos)}")
    return "\n".join(lines)


def handle_org(db, text):
    kw = ["организац","компани","фирм","реквизит","наша компан","юрлиц"]
    if not any(k in text for k in kw):
        return None
    org = db.get("trade",{}).get("org",{})
    if not org or not org.get("name"):
        return "Данные организации не заполнены\\."
    lines = [f"🏢 {_bold('Организация')}\n"]
    if org.get("name"):       lines.append(_line("Название", org["name"]))
    if org.get("inn"):        lines.append(_line("ИНН", org["inn"]))
    if org.get("address"):    lines.append(_line("Адрес", org["address"]))
    if org.get("phone"):      lines.append(_line("Телефон", org["phone"]))
    if org.get("director"):   lines.append(_line("Руководитель", org["director"]))
    if org.get("accountant"): lines.append(_line("Бухгалтер", org["accountant"]))
    return "\n".join(lines)


def handle_payroll(db, text):
    kw = ["зарплат","оклад","налог","отчислен","ставк","соцфонд","подоходн"]
    if not any(k in text for k in kw):
        return None
    pr = db.get("trade",{}).get("payroll",{})
    if not pr:
        return "Зарплатные настройки не заданы\\."
    lines = [f"💼 {_bold('Зарплатные настройки')}\n"]
    lines.append(f"  Подоходный налог: {_bold(str(pr.get('incomeTax',0)) + '%')}")
    lines.append(f"  Соцфонд \\(сотрудник\\): {_bold(str(pr.get('sfEmployee',0)) + '%')}")
    lines.append(f"  Соцфонд \\(работодатель\\): {_bold(str(pr.get('sfEmployer',0)) + '%')}")
    lines.append(f"  Стандартный вычет: {_bold(fmt(pr.get('stdDeduction',0)) + ' сом')}")
    lines.append(f"  Мин\\. зарплата: {_bold(fmt(pr.get('minSalary',0)) + ' сом')}")
    return "\n".join(lines)


def handle_contracts(db, text):
    kw = ["договор","контракт"]
    if not any(k in text for k in kw):
        return None
    contracts = db.get("trade",{}).get("contracts",[])
    if not contracts:
        return "Договоры не найдены\\."
    lines = [f"📝 {_bold('Договоры')} \\({_esc(str(len(contracts)))}\\)\n"]
    for d in contracts:
        c_name = get_contractor_name(db, d.get("contractor",""))
        lines.append(f"  №{_esc(d.get('number','?'))} от {_esc(d.get('date','?'))}")
        lines.append(f"    {_esc(d.get('name',''))} \\| {_esc(c_name)}")
        lines.append("")
    return "\n".join(lines)


# ══════════════════════════════════════════════
#  Маршрутизация
# ══════════════════════════════════════════════

HANDLERS = [
    handle_summary, handle_money, handle_debts, handle_price,
    handle_payroll, handle_org, handle_contracts, handle_warehouses,
    handle_stock, handle_goods, handle_contractors, handle_employees,
]


def process_question(text: str) -> str:
    text_lower = text.lower().strip()
    try:
        db = fetch_db()
    except Exception as e:
        log.error(f"Ошибка загрузки данных: {e}")
        return "⚠️ Не удалось загрузить данные\\. Попробуйте позже\\."

    for handler in HANDLERS:
        result = handler(db, text_lower)
        if result is not None:
            return result
    return None


# ══════════════════════════════════════════════
#  Telegram: кнопки быстрых действий
# ══════════════════════════════════════════════

MAIN_MENU = InlineKeyboardMarkup([
    [InlineKeyboardButton("📊 Сводка", callback_data="q:сводка"),
     InlineKeyboardButton("💰 Деньги", callback_data="q:сколько денег на счетах")],
    [InlineKeyboardButton("📦 Склад", callback_data="q:остатки на складе"),
     InlineKeyboardButton("📊 Долги", callback_data="q:кто нам должен")],
    [InlineKeyboardButton("📋 Товары", callback_data="q:список товаров"),
     InlineKeyboardButton("🤝 Контрагенты", callback_data="q:контрагенты")],
    [InlineKeyboardButton("👥 Сотрудники", callback_data="q:сотрудники"),
     InlineKeyboardButton("🏷 Прайс", callback_data="q:прайс-лист")],
])

HELP_MD = (
    f"Я могу ответить на:\n\n"
    f"  📦 _Сколько \\[товар\\] на складе?_\n"
    f"  💰 _Сколько денег на счетах?_\n"
    f"  🏷 _Сколько стоит \\[товар\\]?_\n"
    f"  📊 _Кто нам должен? / Наши долги_\n"
    f"  📈 _Сводка / Отчёт_\n"
    f"  📋 _Список товаров_\n"
    f"  🤝 _Контрагенты_\n"
    f"  👥 _Сотрудники_\n"
    f"  📝 _Договоры_\n"
    f"  🏢 _Реквизиты организации_\n"
    f"  💼 _Зарплатные ставки_\n\n"
    f"Или нажмите кнопку ниже 👇"
)


# ══════════════════════════════════════════════
#  Telegram: обработчики
# ══════════════════════════════════════════════

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    await update.message.reply_text(
        f"Здравствуйте, {_esc(user.first_name)}\\! 👋\n\n"
        f"Я — бот бухгалтерии\\. Задайте вопрос или выберите раздел:\n\n"
        + HELP_MD,
        parse_mode=ParseMode.MARKDOWN_V2,
        reply_markup=MAIN_MENU,
    )

async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP_MD, parse_mode=ParseMode.MARKDOWN_V2, reply_markup=MAIN_MENU)

async def cmd_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Выберите раздел:", reply_markup=MAIN_MENU)


async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data.startswith("q:"):
        question = data[2:]
        await query.message.chat.send_action(ChatAction.TYPING)
        answer = process_question(question)
        if not answer:
            answer = "Нет данных по этому запросу\\."
        if len(answer) > 4000:
            answer = answer[:4000] + "\n\n\\.\\.\\. \\(обрезано\\)"
        try:
            await query.message.reply_text(answer, parse_mode=ParseMode.MARKDOWN_V2, reply_markup=MAIN_MENU)
        except Exception:
            await query.message.reply_text(answer.replace("\\", ""), reply_markup=MAIN_MENU)


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if not text:
        return
    log.info(f"[{update.effective_user.first_name}] {text}")

    await update.message.chat.send_action(ChatAction.TYPING)

    answer = process_question(text)
    if not answer:
        answer = f"🤔 Не совсем понял вопрос\\.\n\n{HELP_MD}"

    if len(answer) > 4000:
        answer = answer[:4000] + "\n\n\\.\\.\\. \\(обрезано\\)"

    try:
        await update.message.reply_text(answer, parse_mode=ParseMode.MARKDOWN_V2, reply_markup=MAIN_MENU)
    except Exception as e:
        log.warning(f"Markdown error, fallback to plain: {e}")
        plain = re.sub(r'\\(.)', r'\1', answer)
        plain = re.sub(r'\*([^*]+)\*', r'\1', plain)
        await update.message.reply_text(plain, reply_markup=MAIN_MENU)

    log.info(f"[ответ] {answer[:80]}...")


# ══════════════════════════════════════════════
#  Запуск
# ══════════════════════════════════════════════

async def post_init(app: Application):
    await app.bot.set_my_commands([
        BotCommand("start", "Начать работу"),
        BotCommand("menu", "Главное меню"),
        BotCommand("help", "Справка"),
    ])
    log.info("Команды бота зарегистрированы")


def main():
    if not TELEGRAM_TOKEN:
        log.error("TELEGRAM_TOKEN не задан!")
        return
    if not GITHUB_REPO:
        log.error("GITHUB_REPO не задан!")
        return

    log.info("Запуск бота...")
    app = Application.builder().token(TELEGRAM_TOKEN).post_init(post_init).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("menu", cmd_menu))
    app.add_handler(CallbackQueryHandler(handle_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    log.info("Бот запущен!")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
