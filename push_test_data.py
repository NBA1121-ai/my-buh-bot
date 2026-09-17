"""Заливает полную тестовую базу с 10 операциями в GitHub."""
import json, base64, os, requests
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), ".env"))
TOKEN = os.getenv("GITHUB_TOKEN")
REPO = "NBA1121-ai/my-buh-dev"
HEADERS = {"Authorization": f"token {TOKEN}"}

sha = requests.get(f"https://api.github.com/repos/{REPO}/contents/db.json?ref=data", headers=HEADERS).json()["sha"]

db = {
  "bankDocuments": [
    {"id":"t1","type":"payment_in","number":"Тест 1","date":"2026-09-17","status":"conducted",
     "account":"_1","contractor":"c1","article":"_1","currency":"","sum":15000,"rate":1,
     "purpose":"Оплата от Альфа Трейд за ноутбуки","note":"Тест 1 — приход на банк",
     "created":"2026-09-17T08:00:00.000Z"},

    {"id":"t2","type":"payment_out","number":"Тест 2","date":"2026-09-17","status":"conducted",
     "account":"_1","contractor":"c2","article":"_5","currency":"","sum":8000,"rate":1,
     "purpose":"Оплата поставщику Бета Снаб","note":"Тест 2 — расход с банка",
     "created":"2026-09-17T08:10:00.000Z"},

    {"id":"t7","type":"payment_in","number":"Тест 7","date":"2026-09-17","status":"conducted",
     "account":"_2","contractor":"c3","article":"_1","currency":"","sum":20000,"rate":1,
     "purpose":"Оплата от ИП Асанов","note":"Тест 7 — приход на банк",
     "created":"2026-09-17T09:00:00.000Z"},

    {"id":"t8","type":"payment_out","number":"Тест 8","date":"2026-09-17","status":"conducted",
     "account":"_1","contractor":"","article":"_9","currency":"","sum":7000,"rate":1,
     "purpose":"Аренда офиса за сентябрь","note":"Тест 8 — расход с банка",
     "created":"2026-09-17T09:10:00.000Z"},
  ],

  "cashDocuments": [
    {"id":"t3","type":"cash_in","number":"Тест 3","date":"2026-09-17","status":"conducted",
     "cash":"_1","contractor":"c1","article":"_1","currency":"","sum":12000,"rate":1,
     "purpose":"Оплата наличными от Альфа Трейд","note":"Тест 3 — приход в кассу",
     "created":"2026-09-17T08:20:00.000Z"},

    {"id":"t4","type":"cash_out","number":"Тест 4","date":"2026-09-17","status":"conducted",
     "cash":"_1","contractor":"","article":"_6","currency":"","sum":5000,"rate":1,
     "purpose":"Выдача под отчёт на канцтовары","note":"Тест 4 — расход из кассы",
     "created":"2026-09-17T08:30:00.000Z"},
  ],

  "accounts": [
    {"id":"_1","name":"Сомовый счёт","number":"1280010000000100","bank":"АБ Кыргызстан","type":"bank"},
    {"id":"_2","name":"Долларовый счёт","number":"1280020000000200","bank":"НИСПА","type":"bank"},
    {"id":"_3","name":"Касса офиса","number":"","bank":"","type":"cash"},
  ],
  "banks": [
    {"id":"_1","name":"АБ Кыргызстан","bik":"170201001","corrAccount":"1234567890"},
    {"id":"_2","name":"НИСПА","bik":"170202001","corrAccount":"0987654321"},
  ],
  "cashs": [
    {"id":"_1","name":"Касса магазина","responsible":"Иванов И.И."},
    {"id":"_2","name":"Касса офиса","responsible":"Петрова М.М."},
  ],
  "currencies": [
    {"id":"_1","code":"KGS","name":"Кыргызский сом","rate":1},
    {"id":"_2","code":"USD","name":"Доллар США","rate":89},
    {"id":"_3","code":"RUB","name":"Российский рубль","rate":0.95},
  ],
  "articles": [
    {"id":"_1","name":"Оплата от покупателя","type":"in"},
    {"id":"_2","name":"Возврат подотчета","type":"in"},
    {"id":"_3","name":"Взнос учредителя","type":"in"},
    {"id":"_5","name":"Оплата поставщику","type":"out"},
    {"id":"_6","name":"Выдача под отчет","type":"out"},
    {"id":"_7","name":"Зарплата","type":"out"},
    {"id":"_8","name":"Налоги","type":"out"},
    {"id":"_9","name":"Аренда","type":"out"},
  ],

  "trade": {
    "docs": [
      {"id":"t5","type":"postupleniye","number":"Тест 5","date":"2026-09-17","status":"conducted",
       "contractor":"c2","contract":"d2","warehouse":"w1",
       "note":"Поступление товара от Бета Снаб",
       "rows":[
         {"nomenclature":"n1","qty":2,"price":40000,"total":80000},
         {"nomenclature":"n3","qty":20,"price":600,"total":12000},
       ],"total":18000,"created":"2026-09-17T08:40:00.000Z"},

      {"id":"t6","type":"realizaciya","number":"Тест 6","date":"2026-09-17","status":"conducted",
       "contractor":"c1","contract":"d1","warehouse":"w1",
       "note":"Реализация товара Альфа Трейд",
       "rows":[
         {"nomenclature":"n3","qty":5,"price":900,"total":4500},
         {"nomenclature":"n4","qty":10,"price":450,"total":4500},
       ],"total":9000,"created":"2026-09-17T08:50:00.000Z"},

      {"id":"t9","type":"postupleniye","number":"Тест 9","date":"2026-09-17","status":"conducted",
       "contractor":"c2","contract":"d2","warehouse":"w1",
       "note":"Поступление мониторов от Бета Снаб",
       "rows":[
         {"nomenclature":"n2","qty":3,"price":12000,"total":36000},
         {"nomenclature":"n4","qty":20,"price":350,"total":7000},
       ],"total":14800,"created":"2026-09-17T09:20:00.000Z"},

      {"id":"t10","type":"realizaciya","number":"Тест 10","date":"2026-09-17","status":"conducted",
       "contractor":"c3","contract":"d3","warehouse":"w1",
       "note":"Реализация монитора ИП Асанов",
       "rows":[
         {"nomenclature":"n2","qty":1,"price":15000,"total":15000},
       ],"total":10000,"created":"2026-09-17T09:30:00.000Z"},
    ],
    "noDemo": True, "demoSeeded": True, "stockSeeded": True, "hrSeeded": True,
    "org": {
      "name": "ОсОО «Эльмырза»", "inn": "01234567890123",
      "address": "г. Бишкек, ул. Киевская, 1", "phone": "+996 312 000000",
      "director": "Директор", "accountant": "Главный бухгалтер",
    },
    "contractors": [
      {"id":"c1","name":"ОсОО «Альфа Трейд»","full":"ОсОО «Альфа Трейд»","inn":"02001201010101","kind":"Юридическое лицо","address":"г. Бишкек, пр. Чуй, 100","phone":"+996 555 111111"},
      {"id":"c2","name":"ОсОО «Бета Снаб»","full":"ОсОО «Бета Снаб»","inn":"02002202020202","kind":"Юридическое лицо","address":"г. Ош, ул. Ленина, 5","phone":"+996 555 222222"},
      {"id":"c3","name":"ИП Асанов А.А.","full":"ИП Асанов Асан Асанович","inn":"20101199000001","kind":"ИП","address":"г. Бишкек, ул. Ахунбаева, 45","phone":"+996 700 333333"},
    ],
    "contracts": [
      {"id":"d1","contractor":"c1","kind":"С покупателем","number":"1","date":"2026-01-10","name":"Договор с покупателем","currency":"KGS","priceType":"pt1"},
      {"id":"d2","contractor":"c2","kind":"С поставщиком","number":"7","date":"2026-01-15","name":"Договор поставки №7","currency":"KGS","priceType":"pt2"},
      {"id":"d3","contractor":"c3","kind":"С покупателем","number":"3","date":"2026-02-01","name":"Договор №3","currency":"KGS","priceType":"pt1"},
    ],
    "nomenclature": [
      {"id":"n1","name":"Ноутбук Lenovo IdeaPad 3","article":"NB-001","kind":"Товар","unit":"шт","vat":12,"price":48000,"cost":40000},
      {"id":"n2","name":"Монитор Samsung 24","article":"MN-024","kind":"Товар","unit":"шт","vat":12,"price":15000,"cost":12000},
      {"id":"n3","name":"Мышь беспроводная","article":"MS-010","kind":"Товар","unit":"шт","vat":12,"price":900,"cost":600},
      {"id":"n4","name":"Бумага А4, 500 л.","article":"PP-A4","kind":"Материал","unit":"пач","vat":12,"price":450,"cost":350},
      {"id":"n5","name":"Клавиатура USB","article":"KB-005","kind":"Товар","unit":"шт","vat":12,"price":1200,"cost":800},
    ],
    "warehouses": [
      {"id":"w1","name":"Основной склад","kind":"Оптовый склад","responsible":"Иванов И.И."},
      {"id":"w2","name":"Розничный магазин","kind":"Розничный магазин","responsible":"Петрова М.М."},
    ],
    "priceTypes": [
      {"id":"pt1","name":"Оптовая","vatIncluded":True,"markup":20},
      {"id":"pt2","name":"Закупочная","vatIncluded":True,"markup":0},
    ],
    "sfCounter": {"in":0,"out":0},
    "departments": [
      {"id":"dep1","name":"Склад","head":"Иванов И.И."},
      {"id":"dep2","name":"Администрация","head":"Директор"},
    ],
    "nomGroups": [{"id":"ng1","name":"Оргтехника"},{"id":"ng2","name":"Канцтовары"}],
    "costItems": [],
    "units": [
      {"id":"u1","code":"796","name":"шт","full":"Штука"},
      {"id":"u2","code":"778","name":"пач","full":"Пачка"},
    ],
    "payroll": {
      "incomeTax":10,"sfEmployee":10,"sfEmployeePF":8,"sfEmployeeGNPF":2,
      "sfEmployer":17.25,"sfEmployerPF":15,"sfEmployerFOMS":2,"sfEmployerFOT":0.25,
      "stdDeduction":650,"dependentDeduction":100,"minSalary":2460,
    },
    "positions": [
      {"id":"pos1","name":"Директор","baseSalary":90000},
      {"id":"pos2","name":"Бухгалтер","baseSalary":60000},
      {"id":"pos3","name":"Менеджер","baseSalary":40000},
    ],
    "schedules": [{"id":"sch1","name":"Пятидневка","hoursPerDay":8,"daysPerWeek":5}],
    "accruals": [], "deductions": [],
    "employees": [
      {"id":"e1","tabNo":"001","name":"Абдылдаев Бакыт Т.","inn":"21504198500123","position":"pos1","salary":90000,"phone":"+996 555 100001","status":"Работает","hireDate":"2024-01-10"},
      {"id":"e2","tabNo":"002","name":"Петрова Марина М.","inn":"11209197800456","position":"pos2","salary":60000,"phone":"+996 555 100002","status":"Работает","hireDate":"2024-01-10"},
      {"id":"e3","tabNo":"003","name":"Асанова Айгуль А.","inn":"12207199500654","position":"pos3","salary":40000,"phone":"+996 555 100005","status":"Работает","hireDate":"2025-06-02"},
    ],
  },
  "_activityLog": [],
}

b64 = base64.b64encode(json.dumps(db, ensure_ascii=False, indent=2).encode("utf-8")).decode("ascii")
resp = requests.put(
    f"https://api.github.com/repos/{REPO}/contents/db.json",
    headers=HEADERS,
    json={"message": "Полная тестовая база: 10 операций + справочники", "content": b64, "sha": sha, "branch": "data"}
)
print(f"{resp.status_code}: {resp.json().get('commit',{}).get('message', resp.text[:200])}")
