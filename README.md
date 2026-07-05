# AgroDrone.BY — интерфейсы

Telegram Mini App интерфейсы для экосистемы AgroDrone.BY — сервиса
обработки сельскохозяйственных полей агродронами в Республике Беларусь.

## Что это

Набор HTML-интерфейсов, открывающихся внутри Telegram как Mini App:
кабинет агронома, интерфейс оператора, менеджера и руководителя полётов.
Деплой через GitHub Pages.

## Архитектура

```
Telegram Bot → Leadtex → Telegram Mini App (этот репозиторий, GitHub Pages)
  → Cloudflare Workers (CORS-прокси) → Google Apps Script → Google Sheets
```

## Структура

**/agronom/** — кабинет агронома
- `AD_agronom_cabinet.html` — кабинет
- `AD_agronom_field_form.html` — форма поля
- `AD_agronom_register.html` — регистрация

**/operator/** — интерфейс оператора БАК
- `AD_op_menu.html` — главное меню
- `AD_op_trip.html` — командировка
- `AD_op_day.html` — рабочий день
- `AD_op_expenses.html` — расходы
- `AD_op_map.html` — карта полей
- `AD_op_issues.html` — проблемы с техникой
- `AD_op_stats.html` — статистика
- `AD_op_weather.html` — погода
- `AD_op_docs.html` — документы

**Корень**
- `AD_candidate.html` — заявка кандидата
- `AD_client_form.html` — заявка с сайта
- `AD_manager.html` / `AD_manager_v3.html` — интерфейс менеджера
- `AD_my_menu.html` — личное меню
- `AD_stats.html` — статистика
- `AD_rp.html` — интерфейс руководителя полётов
- `AD_admin_coefs.html` — коэффициенты
- `AD_faq.html` — FAQ
- `AD_journal.html` — журнал
- `config.js` — конфигурация (публичные параметры, без ключей)

## Технологии

Чистый HTML/CSS/JS, без фреймворков. Интеграция с Telegram через
Telegram Web App SDK.

## Только гражданское применение

Проект и все его компоненты предназначены исключительно для гражданских
задач сельского хозяйства.

---
© AgroDrone.BY
