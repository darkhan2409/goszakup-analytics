import os
from dotenv import load_dotenv

# Загружаем переменные из .env
load_dotenv()

# Пути проекта
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REPORTS_DIR = os.path.join(BASE_DIR, "reports")

# Создаем папку для отчетов, если её нет
if not os.path.exists(REPORTS_DIR):
    os.makedirs(REPORTS_DIR)

# === КОНФИГУРАЦИЯ API ===

# Токен авторизации (получите на портале goszakup.gov.kz)
TOKEN = os.getenv("TOKEN")

if not TOKEN:
    print("ВНИМАНИЕ: TOKEN не найден в переменном окружении или .env файле!")
    print("Пожалуйста, скопируйте .env.example в .env и укажите ваш токен.")

# Базовый URL API
BASE_URL = "https://ows.goszakup.gov.kz"

# Лимит записей на страницу (макс 200)
PAGE_LIMIT = 200

# === ПАРАМЕТРЫ ПОИСКА ===

# БИН заказчика
BIN_COMPANY = os.getenv("BIN_COMPANY", "000000000000")

# Финансовый год
FIN_YEAR = int(os.getenv("FIN_YEAR", 2024))

# === ПАРАМЕТРЫ ОТЧЁТА ===

# Статусы для включения в отчёт
# 390 = Исполнен, 375 = Частично исполнен, 190 = Действует
CONTRACT_STATUSES = [390, 375, 190]

# Статусы расторгнутых договоров
# 340 = Расторгнут в одностороннем порядке, 350 = Расторгнут по соглашению сторон
TERMINATED_STATUSES = [340, 350]

# Тип договора: 1 - основной, 2 - допик (включаем оба)
CONTRACT_TYPES = [1, 2]

# === ПАРАМЕТРЫ ОТЧЁТА ПО ОБЪЯВЛЕНИЯМ ===

# Период для отчёта по объявлениям (формат: ГГГГ-ММ-ДД)
DATE_FROM = os.getenv("DATE_FROM", "2024-01-01")
DATE_TO = os.getenv("DATE_TO", "2024-12-31")
