# VBA Extractor

> **Extract, view, and analyze VBA macro code from password-protected Microsoft Office files.**

![Python](https://img.shields.io/badge/Python-3.6%2B-blue)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)
![License](https://img.shields.io/badge/License-MIT-green)

---

## English

### What is this?

VBA Extractor is a desktop GUI application for extracting, reading, and analyzing VBA macro code from Microsoft Office files — including **password-protected VBA projects** where the built-in editor simply asks for a password and shows you nothing.

The main purpose is simple: if an Office file contains VBA macros and the project is locked, this tool lets you read the code anyway.

---

### Features

- **VBA extraction** — pulls all macro modules from `.xls`, `.xlsm`, `.xlsb`, `.xltm`, `.docm`, `.pptm`, `.xlam`
- **Password-protected projects** — reads VBA code regardless of project-level password
- **Syntax highlighting** — keywords, built-in functions, types, strings, comments, numbers
- **Accurate line numbers** — separator lines between procedures are not counted, so line numbers in the viewer always match line numbers in analysis reports
- **Suspicious code scanner** — scans all modules and flags potentially dangerous constructs across three severity levels:
  - 🔴 **High** — shell execution, network downloads, registry writes, ADODB streams, `.exe` launches
  - 🟠 **Medium** — auto-run macros (`Auto_Open`, `Workbook_Open`, `Document_Open`), `Shell()`, `FileSystemObject`, `SendKeys`
  - 🟡 **Low / Obfuscation** — `Chr()` chains, `StrReverse`, `Evaluate`, VBProject self-modification
- **Password hash extraction** — extracts the VBA project password hash in formats ready for hashcat (`-m 9700` for legacy `.xls`, `-m 29500` for modern formats)
- **VBA protection removal** — creates an unlocked copy of the file:
  - `.xls` — patches the DPB field (`DPB→DPx` trick)
  - `.xlsm` / `.docm` / `.pptm` — removes the protection block (`CMG`, `DPB`, `GC`, `WEP`, `ID`) directly from `vbaProject.bin` inside the ZIP container
- **Dark / Light theme** — toggle at any time
- **Module list** — all modules listed in the left panel; click to switch between them; modules with scan hits are marked with ⚠
- **Copy & Save** — copy current module, copy all modules, save as `.bas`, save all as ZIP
- **Open in .txt** — opens current view in the system text editor
- **Context menu** — right-click for Copy / Paste / Select All / Copy Hash / Clear
- **Hotkeys** — `Ctrl+C`, `Ctrl+V`, `Ctrl+A` in the output area

---

### Requirements

```
Python 3.6+
pip install oletools olefile
```

`tkinter` is included with most Python distributions. On Linux you may need:

```bash
sudo apt install python3-tk
```

---

### Usage

```bash
python VBA_Extractor_RU.py   # Russian UI
python VBA_Extractor_EN.py   # English UI
```

1. Click **Browse** and select an Office file
2. Click **Extract VBA** — modules appear in the left panel, code on the right
3. After extraction, two additional buttons appear:
   - **Search for suspicious code** — scans all modules, highlights dangerous lines, shows a full report
   - **🔓 Remove VBA protection** — creates an unlocked copy of the file
4. **Extract password hash** — use the hash with hashcat or John the Ripper

---

### How VBA protection removal works

| Format | Method |
|--------|--------|
| `.xls` (OLE binary) | Patches `DPB=` → `DPx=` directly in the file bytes. Excel ignores the unknown tag and opens the project without a password prompt. |
| `.xlsm`, `.docm`, `.pptm`, etc. (ZIP container) | Extracts `vbaProject.bin`, removes the `CMG=`, `DPB=`, `GC=`, `WEP=`, `ID=` lines from the PROJECT stream, repacks the ZIP. The `DPB→DPx` trick does **not** work for these formats — Excel validates the entire protection block. |

After opening the unlocked file in Excel/Word/PowerPoint, click OK on any recovery prompt, then open the VBA editor with `Alt+F11`.

---

### Screenshots

> *Dark theme · Module list · Syntax highlighting · Line numbers*

---

### Disclaimer

This tool is intended for analyzing files you own or have explicit permission to examine. Use responsibly.

---

---
### Afterword

In 2002, Carlos Rondão, a Portuguese programmer and professor at the Lisbon School of Business and Economics, wrote Tetris directly in Excel—in pure VBA, without any external libraries. And, as often happens with people who truly love their work, he password-protected the project—not out of spite, perhaps, just for the hell of it.

Carlos is no longer alive.

I stumbled upon this file and wanted to see how he did it. The standard editor, of course, asked for a password. That was the impetus for writing the app.
---

## Русский

### Что это?

VBA Extractor — десктопное приложение для извлечения, просмотра и анализа VBA-кода из файлов Microsoft Office, в том числе из **запароленных VBA-проектов**, где встроенный редактор просто показывает поле ввода пароля и ничего больше.

Главная цель — читать код макросов, когда проект заблокирован.

---

### Возможности

- **Извлечение VBA** — достаёт все модули из `.xls`, `.xlsm`, `.xlsb`, `.xltm`, `.docm`, `.pptm`, `.xlam`
- **Запароленные проекты** — читает код независимо от пароля на VBA-проект
- **Подсветка синтаксиса** — ключевые слова, встроенные функции, типы, строки, комментарии, числа
- **Точная нумерация строк** — разделители между процедурами (─────) не нумеруются, поэтому номера строк в просмотрщике всегда совпадают с номерами в отчётах анализа
- **Поиск подозрительного кода** — сканирует все модули и помечает потенциально опасные конструкции по трём уровням:
  - 🔴 **Высокий** — запуск процессов, загрузка файлов из сети, запись в реестр, ADODB, запуск `.exe`
  - 🟠 **Средний** — макросы автозапуска (`Auto_Open`, `Workbook_Open`, `Document_Open`), `Shell()`, `FileSystemObject`, `SendKeys`
  - 🟡 **Низкий / Обфускация** — цепочки `Chr()`, `StrReverse`, `Evaluate`, самомодификация через VBProject
- **Извлечение хэша пароля** — достаёт хэш пароля VBA-проекта в формате для hashcat (`-m 9700` для старых `.xls`, `-m 29500` для новых форматов)
- **Снятие защиты VBA** — создаёт разблокированную копию файла:
  - `.xls` — патчит поле DPB (`DPB→DPx`)
  - `.xlsm` / `.docm` / `.pptm` — удаляет блок защиты (`CMG`, `DPB`, `GC`, `WEP`, `ID`) прямо из `vbaProject.bin` внутри ZIP-контейнера
- **Тёмная / светлая тема** — переключается в любой момент
- **Список модулей** — все модули в левой панели; клик для переключения; модули с находками помечаются ⚠
- **Копирование и сохранение** — скопировать текущий модуль, скопировать все, сохранить как `.bas`, сохранить все в ZIP
- **Открыть в .txt** — открывает текущий вид в системном текстовом редакторе
- **Контекстное меню** — ПКМ: Копировать / Вставить / Выделить всё / Копировать хэш / Очистить
- **Горячие клавиши** — `Ctrl+C`, `Ctrl+V`, `Ctrl+A` в области вывода

---

### Требования

```
Python 3.6+
pip install oletools olefile
```

`tkinter` входит в стандартную поставку Python. На Linux может потребоваться:

```bash
sudo apt install python3-tk
```

---

### Запуск

```bash
python VBA_Extractor_RU.py   # Русский интерфейс
python VBA_Extractor_EN.py   # Английский интерфейс
```

1. Нажмите **Обзор** и выберите файл Office
2. Нажмите **Извлечь VBA** — модули появятся в левой панели, код — справа
3. После извлечения появятся две дополнительные кнопки:
   - **Поиск подозрительного кода** — сканирует все модули, подсвечивает опасные строки, формирует отчёт
   - **🔓 Снять защиту VBA** — создаёт разблокированную копию файла
4. **Извлечь хэш пароля** — используйте хэш в hashcat или John the Ripper

---

### Как работает снятие защиты

| Формат | Метод |
|--------|-------|
| `.xls` (бинарный OLE) | Патчит `DPB=` → `DPx=` прямо в байтах файла. Excel не знает тег `DPx`, игнорирует его и открывает проект без запроса пароля. |
| `.xlsm`, `.docm`, `.pptm` и др. (ZIP-контейнер) | Извлекает `vbaProject.bin`, удаляет строки `CMG=`, `DPB=`, `GC=`, `WEP=`, `ID=` из потока PROJECT, переупаковывает ZIP. Трюк `DPB→DPx` для этих форматов **не работает** — Excel проверяет весь блок защиты целиком. |

После открытия разблокированного файла в Excel/Word/PowerPoint нажмите ОК на предупреждении о восстановлении, затем откройте редактор VBA через `Alt+F11`.

---

### Дисклеймер

Инструмент предназначен для анализа файлов, которые принадлежат вам или которые вы имеете право исследовать.

---

### Послесловие

В 2002 году португальский программист и преподаватель Лиссабонской школы бизнеса и экономики **Карлос Рондао** написал Тетрис прямо в Excel — на чистом VBA, без внешних библиотек. И, как это часто бывает у людей, которые по-настоящему любят своё дело, запаролил проект — не из вредности, наверное, просто так.

Карлоса уже нет в живых.

Я наткнулся на этот файл и захотел посмотреть, как он это сделал. Стандартный редактор, понятное дело, просил пароль. Это и стало поводом написать приложение.
