"""
Telegram-бот для my-buh-dev
Читает db.json из GitHub (ветка data) и отвечает на вопросы:
- остатки на складе
- деньги на счетах
- список товаров, контрагентов, сотрудников
"""

import json
import logging
import os
import re
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

# --- Загрузка данных с GitHub ---

def fetch_db() -> dict:
    """Загружает db.json из ветки data репозитория."""
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{DATA_FILE}?ref={DATA_BRANCH}"
    headers = {"Accept": "application/vnd.github.v3.raw"}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    resp = requests.get(url, headers=headers, timeout=15)
    resp.raise_for_status()
    return resp.json()


# --- Вычисление остатков на складе ---

def calc_stock(db: dict) -> dict:
    """
    Считает остатки по складским документам.
    Приход (prihod/postupleniye) -> +qty
    Расход (rashod/realizaciya) -> -qty
    Возвращает: { nom_id: { warehouse_id: qty } }
    """
    stock = {}
    docs = db.get("trade", {}).get("docs", [])
    for doc in docs:
        dtype = doc.get("type", "")
        is_in = dtype in ("prihod", "postupleniye", "receipt", "purchase", "vozvrat_pokup")
        is_out = dtype in ("rashod", "realizaciya", "sale", "shipment", "vozvrat_post", "spisaniye")
        if not is_in and not is_out:
            # попробуем угадать по наличию типовых полей
            if "rows" in doc:
                pass  # обработаем строки ниже
            else:
                continue
        wh = doc.get("warehouse", doc.get("warehouseFrom", ""))
        rows = doc.get("rows", [])
        for row in rows:
            nom_id = row.get("nomenclature", row.get("nom", row.get("id", "")))
            qty = float(row.get("qty", row.get("quantity", 0)))
            if is_out:
                qty = -qty
            if nom_id:
                stock.setdefault(nom_id, {})
                stock[nom_id][wh] = stock[nom_id].get(wh, 0) + qty
    return stock


# --- Вычисление остатков по счетам ---

def calc_balances(db: dict) -> dict:
    """
    Считает остатки по банковским/кассовым документам.
    payment_in -> + на счёт
    payment_out -> - со счёта
    Возвращает: { account_id: сумма }
    """
    balances = {}
    for doc in db.get("bankDocuments", []):
        acc = doc.get("account", "")
        s = float(doc.get("sum", 0))
        if doc.get("type") == "payment_in":
            balances[acc] = balances.get(acc, 0) + s
        elif doc.get("type") == "payment_out":
            balances[acc] = balances.get(acc, 0) - s
    for doc in db.get("cashDocuments", []):
        acc = doc.get("cash", doc.get("account", ""))
        s = float(doc.get("sum", 0))
        if doc.get("type") in ("cash_in", "pko"):
            balances[acc] = balances.get(acc, 0) + s
        elif doc.get("type") in ("cash_out", "rko"):
            balances[acc] = balances.get(acc, 0) - s
    return balances


# --- Поиск по тексту ---

def stem(word: str) -> str:
    """Грубая обрезка окончаний для русских слов."""
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
    """Нечёткое сравнение: стемминг + все слова запроса есть в названии."""
    q_words = [stem(w) for w in query.lower().split() if len(w) > 1]
    if not q_words:
        return False
    name_lower = name.lower()
    name_stems = [stem(w) for w in name_lower.split()]
    for qw in q_words:
        if not any(qw in ns or ns in qw or qw in name_lower for ns in name_stems):
            return False
    return True


def find_nomenclature(db: dict, query: str) -> list:
    """Ищет номенклатуру по названию."""
    noms = db.get("trade", {}).get("nomenclature", [])
    results = []
    for n in noms:
        if fuzzy_match(query, n.get("name", "")):
            results.append(n)
    return results


def find_contractor(db: dict, query: str) -> list:
    contractors = db.get("trade", {}).get("contractors", [])
    return [c for c in contractors if fuzzy_match(query, c.get("name", "") + " " + c.get("full", ""))]


def find_employee(db: dict, query: str) -> list:
    employees = db.get("trade", {}).get("employees", [])
    return [e for e in employees if fuzzy_match(query, e.get("name", ""))]


# --- Форматирование ---

def fmt_number(n: float) -> str:
    if n == int(n):
        return f"{int(n):,}".replace(",", " ")
    return f"{n:,.2f}".replace(",", " ")


def get_account_name(db: dict, acc_id: str) -> str:
    for a in db.get("accounts", []):
        if a["id"] == acc_id:
            return a.get("name", acc_id)
    return acc_id or "Без счёта"


def get_warehouse_name(db: dict, wh_id: str) -> str:
    for w in db.get("trade", {}).get("warehouses", []):
        if w["id"] == wh_id:
            return w.get("name", wh_id)
    return wh_id or "Без склада"


# --- Обработка вопросов ---

def process_question(text: str) -> str:
    """Главная логика: разбирает вопрос и формирует ответ."""
    text_lower = text.lower().strip()

    try:
        db = fetch_db()
    except Exception as e:
        log.error(f"Ошибка загрузки данных: {e}")
        return "Не удалось загрузить данные с сервера. Попробуйте позже."

    # --- Команда: деньги / баланс (проверяем ДО склада, чтобы "сколько денег" не уходило в склад) ---
    money_keywords = ["деньг", "денег", "денежн", "баланс", "счёт", "счет", "касс", "остаток по счет", "финанс"]
    if any(kw in text_lower for kw in money_keywords):
        balances = calc_balances(db)
        if not balances:
            return "Нет данных по движению денег."
        lines = ["Остатки по счетам:"]
        total = 0
        for acc_id, amount in balances.items():
            acc_name = get_account_name(db, acc_id)
            lines.append(f"  {acc_name}: {fmt_number(amount)} сом")
            total += amount
        lines.append(f"\nИтого: {fmt_number(total)} сом")
        return "\n".join(lines)

    # --- Команда: список складов (только если спрашивают именно про склады, не про остатки) ---
    if re.search(r'\bсклад[ыа]?\b', text_lower) and not any(kw in text_lower for kw in ["сколько", "остат", "наличи", "есть", "что"]):
        warehouses = db.get("trade", {}).get("warehouses", [])
        if not warehouses:
            return "Склады не найдены."
        lines = ["Склады:"]
        for w in warehouses:
            lines.append(f"  {w['name']} ({w.get('kind', '')}), ответственный: {w.get('responsible', '-')}")
        return "\n".join(lines)

    # --- Команда: остатки на складе (конкретный товар) ---
    stock_keywords = ["склад", "остат", "наличи", "сколько", "есть ли", "имеется"]
    if any(kw in text_lower for kw in stock_keywords):
        # Убираем ключевые слова чтобы получить название товара
        search = text_lower
        # Убираем пунктуацию и стоп-слова
        search = re.sub(r'[?!.,;:\-—–()\"\'«»]', ' ', search)
        stop_stems = set()
        for kw in stock_keywords + [
            "на", "у", "нас", "в", "по", "мне", "нам", "ещё", "еще",
            "есть", "ли", "же", "бы", "не", "и", "а", "то",
            "товар", "штук", "штуки", "кг", "шт",
            "скажи", "покажи", "какой", "какая", "какие", "каков",
            "склад", "складе", "складу", "остатки", "остаток", "остатков",
            "что", "всё", "все", "весь", "общий", "общие", "итого",
        ]:
            stop_stems.add(stem(kw))
            stop_stems.add(kw)
        words = [w for w in search.split() if w not in stop_stems and stem(w) not in stop_stems and len(w) > 1]
        search = " ".join(words)

        # Если указан конкретный товар — ищем его
        if search:
            noms = find_nomenclature(db, search)
            if not noms:
                all_noms = db.get("trade", {}).get("nomenclature", [])
                suggestions = [n["name"] for n in all_noms if any(w in n["name"].lower() for w in search.split() if len(w) > 2)]
                if suggestions:
                    return f"Товар '{search}' не найден. Может вы имели в виду:\n" + "\n".join(f"  - {s}" for s in suggestions[:10])
                return f"Товар '{search}' не найден в номенклатуре."
        else:
            # Без конкретного товара — показываем ВСЕ остатки
            noms = db.get("trade", {}).get("nomenclature", [])

        stock = calc_stock(db)
        has_any_stock = any(stock.get(n["id"]) for n in noms)

        if not has_any_stock:
            lines = ["На складе пока нет движений (приход/расход не оформлялся).\n"]
            lines.append(f"Номенклатура ({len(noms)} позиций):")
            for nom in noms:
                if nom.get("kind") in ("Товар", "Материал", "Продукция", "Тара"):
                    lines.append(f"  {nom['name']} [{nom.get('kind','')}] — 0 {nom.get('unit', 'шт')}")
            lines.append("\nЧтобы появились остатки, оформите поступление товара на сайте.")
            return "\n".join(lines)

        lines = ["Остатки на складе:"]
        for nom in noms:
            nom_stock = stock.get(nom["id"], {})
            total = sum(nom_stock.values())
            if total != 0:
                lines.append(f"  {nom['name']}: {fmt_number(total)} {nom.get('unit', 'шт')}")
                for wh_id, qty in nom_stock.items():
                    wh_name = get_warehouse_name(db, wh_id)
                    lines.append(f"      {wh_name}: {fmt_number(qty)} {nom.get('unit', 'шт')}")
            elif not search:
                # При общем запросе показываем только товары с остатками
                pass
            else:
                lines.append(f"  {nom['name']}: 0 {nom.get('unit', 'шт')} (нет движений)")
        return "\n".join(lines)

    # --- Команда: список товаров ---
    if any(kw in text_lower for kw in ["товар", "номенклатур", "ассортимент", "что продаём", "что продаем"]):
        noms = db.get("trade", {}).get("nomenclature", [])
        if not noms:
            return "Номенклатура пуста."
        lines = [f"Номенклатура ({len(noms)} позиций):"]
        for n in noms:
            kind = n.get("kind", "")
            price = fmt_number(n.get("price", 0))
            lines.append(f"  {n['name']} [{kind}] - {price} сом/{n.get('unit', 'шт')}")
        return "\n".join(lines)

    # --- Команда: контрагенты ---
    if any(kw in text_lower for kw in ["контрагент", "поставщик", "покупател", "клиент", "партнёр", "партнер"]):
        search = text_lower
        for kw in ["контрагент", "поставщик", "покупател", "клиент", "партнёр", "партнер", "список", "все", "кто"]:
            search = search.replace(kw, "")
        search = search.strip()
        contractors = db.get("trade", {}).get("contractors", [])
        if search:
            contractors = find_contractor(db, search)
        if not contractors:
            return "Контрагенты не найдены."
        lines = [f"Контрагенты ({len(contractors)}):"]
        for c in contractors:
            lines.append(f"  {c['name']} (ИНН: {c.get('inn', '-')})")
        return "\n".join(lines)

    # --- Команда: сотрудники ---
    if any(kw in text_lower for kw in ["сотрудник", "работник", "персонал", "кадр", "штат"]):
        employees = db.get("trade", {}).get("employees", [])
        positions = {p["id"]: p["name"] for p in db.get("trade", {}).get("positions", [])}
        if not employees:
            return "Список сотрудников пуст."
        lines = [f"Сотрудники ({len(employees)}):"]
        for e in employees:
            pos_id = e.get("position", "")
            pos_name = positions.get(pos_id, pos_id)
            lines.append(f"  {e['name']} - {pos_name}")
        return "\n".join(lines)

    # --- Команда: склады ---
    if any(kw in text_lower for kw in ["склад"]) and not any(kw in text_lower for kw in ["сколько", "остат"]):
        warehouses = db.get("trade", {}).get("warehouses", [])
        if not warehouses:
            return "Склады не найдены."
        lines = ["Склады:"]
        for w in warehouses:
            lines.append(f"  {w['name']} ({w.get('kind', '')}), ответственный: {w.get('responsible', '-')}")
        return "\n".join(lines)

    # --- Не распознано ---
    return (
        "Не совсем понял вопрос. Я могу ответить на:\n\n"
        "  - Сколько [товар] на складе?\n"
        "  - Сколько денег на счетах?\n"
        "  - Список товаров\n"
        "  - Контрагенты\n"
        "  - Сотрудники\n"
        "  - Склады\n\n"
        "Просто напишите вопрос!"
    )


# --- Telegram handlers ---

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Я бот бухгалтерии my-buh-dev.\n\n"
        "Спросите меня:\n"
        "  - Сколько ноутбуков на складе?\n"
        "  - Сколько денег на счетах?\n"
        "  - Список товаров\n"
        "  - Контрагенты\n"
        "  - Сотрудники\n\n"
        "Просто напишите вопрос!"
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if not text:
        return
    log.info(f"Вопрос от {update.effective_user.first_name}: {text}")
    answer = process_question(text)
    await update.message.reply_text(answer)
    log.info(f"Ответ: {answer[:100]}...")


def main():
    log.info("Запуск бота...")
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    log.info("Бот запущен! Ожидаю сообщения...")
    app.run_polling()


if __name__ == "__main__":
    main()
