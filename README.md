# WRS Mobile

Мобильная веб-версия **WRS — Автоматизация отчётности** для деплоя на [Railway](https://railway.app).

## Возможности (MVP)

- Импорт **адресной книги** (Excel: «Адрес магазина», «Менеджер»)
- Построение **отчёта WRS** (не авторизованные, только прилегающая, SL по менеджерам/регионам)
- **Корректировки SL** (Н/к уборка, реконструкция, тех. причины)
- **Анализ коротких смен** (< 3 ч)
- **Авторизация дворников** (график прилегающей + сверка со сменами, SL)
- Адаптивный UI (телефон / планшет), тёмная Steam-подобная тема как в десктопе

## Локальный запуск

```bash
cd wrs-mobile
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Откройте http://localhost:8000

## Деплой на Railway

1. Создайте проект на [railway.app](https://railway.app) и подключите этот репозиторий (или `railway up` из CLI).
2. Railway подхватит `Dockerfile` и `railway.toml`.
3. **Важно:** данные (адресная книга, корректировки) пишутся в `/data`.  
   Добавьте **Volume** с mount path `/data`, иначе файлы пропадут при рестарте.
4. Переменная `PORT` выставляется Railway автоматически.
5. Healthcheck: `GET /health`

### CLI (кратко)

```bash
npm i -g @railway/cli
railway login
railway init
railway up
railway volume add --mount /data   # если поддерживается в вашем плане
```

## Структура

```
wrs-mobile/
  app/
    main.py              # FastAPI
    config.py
    services/            # process_wrs, process_shifts
    utils/               # address, helpers, storage
    templates/           # mobile HTML
    static/app.css
  data/                  # JSON (локально)
  Dockerfile
  railway.toml
  requirements.txt
```

## Формула SL

```
SL = (Всего − (Н/к уборка + Не авт + Прилег) + (Реконстр + Тех)) / Всего
```

результат ограничен 0–100%.

## Отличия от десктопа

| Десктоп (Tkinter) | Mobile (web) |
|---|---|
| Локальные файлы | Upload Excel в браузере |
| Полный Excel-экспорт смен | Список коротких смен + копирование отчёта |
| Графики matplotlib | KPI + прогресс-бары SL |
| SQLite history / day pipeline | JSON history (расширяемо) |

Дальше можно добавить: реконструкцию UI, сверку часов, экспорт Excel, авторизацию.
