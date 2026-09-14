# Адреси проєкту

Цей файл — не для git-логіки (немає секретів), а для швидкого орієнтування:
що відкривати для якої мети.

## Що відкривати викладачу

| Мета | Адреса |
|------|--------|
| Сканер відвідуваності (щодня на парі) | https://vyklyuk.github.io/QR-students/scanner.html |
| Генератор карток (роздрукувати нові картки) | https://vyklyuk.github.io/QR-students/generator.html |
| Стартова сторінка з посиланнями на обидві | https://vyklyuk.github.io/QR-students/ |
| Сама Google Таблиця (дані, звіти) | https://docs.google.com/spreadsheets/d/1zDfmwpkYDvbMF12KW2uNcQheIiKK5I_AIIPNTxV5RMM/edit?gid=628695362#gid=628695362 |
| QR-код «СТОП» (сигнал завершення сканування для «Команди» на iPhone) | https://raw.githubusercontent.com/vyklyuk/QR-students/main/stop-qr.png |

## Технічне (для налагодження)

| Мета | Де |
|------|----|
| Перевірити, що бекенд живий (має повернути `{"status":"ok",...}`) | Відкрити в браузері поточний URL з `web/config.local.js` / секрету `APPS_SCRIPT_URL` |
| Редактор Apps Script (код бекенда, Script Properties, розгортання) | У самій Google Таблиці: **Розширення → Apps Script** |
| Секрет `APPS_SCRIPT_URL` на GitHub (те, що бекенд віддає сторінкам) | https://github.com/vyklyuk/QR-students/settings/secrets/actions |
| Налаштування GitHub Pages | https://github.com/vyklyuk/QR-students/settings/pages |
| Історія деплоїв сторінок (Actions) | https://github.com/vyklyuk/QR-students/actions/workflows/pages.yml |
| Сам репозиторій | https://github.com/vyklyuk/QR-students |

## Коли міняється URL бекенда

URL веб-застосунку **не змінюється**, якщо оновлювати наявне розгортання:
**Розгорнути → Керувати розгортаннями → олівець → Версія: Нова версія → Розгорнути**
(див. `SETUP.md`, розділ «Найважливіша пастка»).

Якщо ж у редакторі Apps Script натиснути **Розгорнути → Нове розгортання**,
з'явиться **новий** URL. Тоді треба:

1. Оновити секрет `APPS_SCRIPT_URL` на GitHub (посилання вище) новим значенням.
2. Запустити деплой сторінок ще раз: вкладка **Actions** в репозиторії →
   **Deploy GitHub Pages** → **Run workflow** (або просто зробити будь-який
   push у `web/`).

Сам URL ніде в git-репозиторії не зберігається (`CLAUDE.md`, правило 5) —
тримається лише в секреті GitHub Actions і в приватному `web/config.local.js`.
