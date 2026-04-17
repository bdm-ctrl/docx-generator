# DOCX Generator Service

Сервис генерации Word документов для Telegram бота.

## API

### POST /generate

Генерирует документ на основе шаблона.

**Body:**
```json
{
  "doc_type": "doc_act",
  "fields": {
    "doc_number": "17042026-11",
    "doc_date": "17.04.2026",
    "client_name": "ООО Пример",
    "amount_number": "120000",
    "amount_words": "Сто двадцать тысяч рублей",
    "event_name": "Миссия 8 бит",
    "event_date": "19.04.2026",
    "event_time": "12:00",
    "participants": "35",
    "service_description": "Организация мероприятия",
    "payment_deadline": "27.04.2026"
  }
}
```

**Возвращает:** .docx файл

### GET /health

Проверка статуса сервиса.
