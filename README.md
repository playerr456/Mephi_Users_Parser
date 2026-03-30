# MEPhI users -> Excel

## Установка

```bash
python -m pip install -r requirements.txt
```

## Запуск

```bash
python parse_users_to_excel.py -u YOUR_LOGIN
```

Скрипт запросит пароль интерактивно и сохранит файл `mephi_users.xlsx`.

## Файл с логином и паролем

Создайте рядом со скриптом файл `mephi_credentials.env`:

```env
MEPHI_USERNAME=YOUR_LOGIN
MEPHI_PASSWORD=YOUR_PASSWORD
```

После этого можно запускать без параметров:

```bash
python parse_users_to_excel.py
```

Если хотите другой путь к файлу:

```bash
python parse_users_to_excel.py --creds-file C:\path\to\creds.env
```

### Полезные параметры

```bash
python parse_users_to_excel.py -u YOUR_LOGIN -o users.xlsx --verbose
python parse_users_to_excel.py --max-pages 1000 --timeout 60
python parse_users_to_excel.py --start-page 501 --max-pages 1017 -o users_part2.xlsx --verbose
python parse_users_to_excel.py --save-every-pages 10 -o users_live.xlsx
```

`--save-every-pages` (по умолчанию `20`) включает промежуточное сохранение:
файл Excel обновляется каждые N страниц, и можно видеть, что строки уже появляются в процессе.

Также можно передать учетные данные через переменные окружения:

```bash
set MEPHI_USERNAME=YOUR_LOGIN
set MEPHI_PASSWORD=YOUR_PASSWORD
python parse_users_to_excel.py
```

В Excel будут колонки: `ФИО`, `Номер группы`.
Если у пользователя на странице нет номера группы, в колонку записывается `NULL`.
