import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import re
import zipfile
import tempfile
import subprocess
import platform
import shutil

try:
    from oletools.olevba import VBA_Parser
except ImportError:
    messagebox.showerror("Ошибка", "Библиотека 'oletools' не установлена.\nВыполните: pip install oletools")
    exit(1)

try:
    import olefile
except ImportError:
    messagebox.showerror("Ошибка", "Библиотека 'olefile' не установлена.\nВыполните: pip install olefile")
    exit(1)


def clean_vba_code(raw_code):
    """Удаляет Attribute VB_* и блоки p-code, возвращает чистый код или None."""
    if not raw_code:
        return None
    lines = raw_code.splitlines()
    if any(line.strip().startswith("' Line #") for line in lines):
        return None
    cleaned = [line for line in lines if not line.strip().lower().startswith("attribute vb_")]
    result = "\n".join(cleaned).strip()
    return result if result else None


class LineNumberedText(tk.Frame):
    """
    Виджет с нумерацией строк на Canvas.
    Использует dlineinfo для пиксель-в-пиксель позиционирования.
    Разделители (─────) пропускаются.
    Поддерживает горизонтальную прокрутку при wrap='none'.
    """
    def __init__(self, master, theme_colors=None, **kwargs):
        super().__init__(master)
        self.theme_colors = theme_colors or {}
        self.line_mapping = []          # [code_line_num | None] по виджет-строкам (0-based)
        self._fg_color = "#555555"
        self._bg_color = "#1e1e1e"
        self.show_line_numbers = True   # флаг показа номеров строк
        self.wrap_mode = kwargs.pop('wrap', 'word')

        self.text_frame = tk.Frame(self)
        self.text_frame.pack(fill=tk.BOTH, expand=True)

        # Canvas для номеров
        self.canvas = tk.Canvas(
            self.text_frame, width=40, bd=0,
            highlightthickness=0, takefocus=0
        )
        self.canvas.pack(side=tk.LEFT, fill=tk.Y)

        # Вертикальный скроллбар
        self.v_scrollbar = tk.Scrollbar(self.text_frame, orient=tk.VERTICAL)
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Горизонтальный скроллбар (появляется только при wrap='none')
        self.h_scrollbar = tk.Scrollbar(self, orient=tk.HORIZONTAL)
        # Пока скрыт, будет упакован при необходимости

        self.text = tk.Text(
            self.text_frame, wrap=self.wrap_mode,
            yscrollcommand=self._on_text_scroll,
            xscrollcommand=self._on_text_xscroll,
            font=("Consolas", 10), border=0, highlightthickness=0,
            padx=8, pady=6, **kwargs
        )
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.v_scrollbar.config(command=self._on_scrollbar)
        self.h_scrollbar.config(command=self._on_hscrollbar)

        self.text.bind("<MouseWheel>", self._on_mousewheel)
        self.text.bind("<Button-4>",   self._on_mousewheel)
        self.text.bind("<Button-5>",   self._on_mousewheel)
        self.text.bind("<Configure>",  self._update_line_numbers)

        if self.wrap_mode == 'none':
            self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

    # ---- Скролл ----
    def _on_text_scroll(self, *args):
        self.v_scrollbar.set(*args)
        self._redraw_line_numbers()

    def _on_text_xscroll(self, *args):
        self.h_scrollbar.set(*args)

    def _on_scrollbar(self, *args):
        self.text.yview(*args)
        self._redraw_line_numbers()

    def _on_hscrollbar(self, *args):
        self.text.xview(*args)

    def _on_mousewheel(self, event):
        if event.num == 4:
            self.text.yview_scroll(-1, "units")
        elif event.num == 5:
            self.text.yview_scroll(1, "units")
        else:
            self.text.yview_scroll(int(-1*(event.delta/120)), "units")
        self._redraw_line_numbers()
        return "break"

    def configure_wrap(self, mode):
        """Переключает режим переноса и управляет горизонтальным скроллбаром."""
        if mode == self.wrap_mode:
            return
        self.wrap_mode = mode
        self.text.configure(wrap=mode)
        if mode == 'none':
            if not self.h_scrollbar.winfo_ismapped():
                self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        else:
            self.h_scrollbar.pack_forget()
        self._update_line_numbers()

    # ---- Управление номерами строк ----
    def enable_line_numbers(self, enable):
        self.show_line_numbers = enable
        self._update_line_numbers()

    # ---- Карта строк ----
    def _rebuild_line_mapping(self):
        if not self.show_line_numbers:
            self.line_mapping = []
            return

        content = self.text.get("1.0", "end-1c")
        self.line_mapping = []
        code_line = 0
        for line in content.splitlines():
            stripped = line.strip()
            if stripped and all(c == "─" for c in stripped):
                self.line_mapping.append(None)
            else:
                code_line += 1
                self.line_mapping.append(code_line)

    # ---- Отрисовка номеров ----
    def _redraw_line_numbers(self, event=None):
        self.canvas.delete("all")
        if not self.show_line_numbers:
            return

        # Страховка: если карта расходится с количеством строк, перестроить
        try:
            real_lines = int(self.text.index('end-1c').split('.')[0])
        except:
            real_lines = 0
        if not self.line_mapping or len(self.line_mapping) != real_lines:
            self._rebuild_line_mapping()

        if not self.line_mapping:
            return

        max_num = max((n for n in self.line_mapping if n is not None), default=0)
        digits = max(2, len(str(max_num)))
        width = digits * 7 + 14
        self.canvas.config(width=width)

        idx = self.text.index("@0,0")
        while True:
            dline = self.text.dlineinfo(idx)
            if dline is None:
                break
            y = dline[1]
            linenum = int(str(idx).split(".")[0])

            if 0 < linenum <= len(self.line_mapping):
                code_num = self.line_mapping[linenum - 1]
                if code_num is not None:
                    self.canvas.create_text(
                        width - 6, y,
                        anchor="ne",
                        text=str(code_num),
                        font=("Consolas", 10),
                        fill=self._fg_color
                    )

            nxt = self.text.index(f"{idx}+1line")
            if nxt == idx:
                break
            idx = nxt

    def _update_line_numbers(self, event=None):
        self._rebuild_line_mapping()
        self._redraw_line_numbers()

    def get_code_to_widget_map(self):
        result = {}
        for w_idx, c_num in enumerate(self.line_mapping, start=1):
            if c_num is not None and c_num not in result:
                result[c_num] = w_idx
        return result

    # ---- Прокси-методы к self.text ----
    def insert(self, index, text, *args):
        self.text.insert(index, text, *args)
        self._update_line_numbers()

    def delete(self, start, end=None):
        if end:
            self.text.delete(start, end)
        else:
            self.text.delete(start)
        self._update_line_numbers()

    def get(self, start, end=None):
        return self.text.get(start, end) if end else self.text.get(start)

    def see(self, index):
        self.text.see(index)
        self._redraw_line_numbers()

    def see_code_line(self, code_line_num):
        for w_idx, orig in enumerate(self.line_mapping, start=1):
            if orig == code_line_num:
                self.text.see(f"{w_idx}.0")
                self._redraw_line_numbers()
                return

    def tag_configure(self, tag, **kwargs): self.text.tag_configure(tag, **kwargs)
    def tag_add(self, tag, start, end=None):
        if end:
            self.text.tag_add(tag, start, end)
        else:
            self.text.tag_add(tag, start)
    def tag_remove(self, tag, start, end=None):
        if end:
            self.text.tag_remove(tag, start, end)
        else:
            self.text.tag_remove(tag, start)
    def tag_raise(self, tag):    self.text.tag_raise(tag)
    def mark_set(self, n, i):    self.text.mark_set(n, i)
    def edit_modified(self, v=None):
        return self.text.edit_modified(v) if v is None else self.text.edit_modified(v)
    def yview(self, *a):         return self.text.yview(*a)
    def configure(self, **kw):   self.text.configure(**kw)
    def config(self, **kw):      self.text.config(**kw)

    def bind(self, seq=None, func=None, add=None):
        if seq is None:  return self.text.bind()
        if func is None: return self.text.bind(seq)
        return self.text.bind(seq, func, add)

    def apply_theme(self, colors):
        self.theme_colors = colors
        self._fg_color = colors["separator"]
        self._bg_color = colors["bg_main"]
        self.canvas.configure(bg=colors["bg_main"])
        self.configure(
            bg=colors["bg_main"],
            fg=colors["fg_text"],
            insertbackground=colors["fg_text"],
            selectbackground=colors["selection"],
            selectforeground="white",
        )
        self._redraw_line_numbers()


class SimpleVBAExtractor:
    THEMES = {
        "dark": {
            "bg_main": "#1e1e1e", "bg_panel": "#252526", "bg_input": "#3c3c3c",
            "fg_text": "#d4d4d4", "fg_comment": "#6a9955", "keyword": "#569cd6",
            "string": "#ce9178", "number": "#b5cea8", "builtin": "#dcdcaa",
            "type": "#4ec9b0", "operator": "#d4d4d4", "separator": "#555555",
            "accent_green": "#4e9a06", "accent_blue": "#3584e4",
            "accent_orange": "#cd9309", "accent_red": "#c01c28",
            "border": "#3e3e42", "btn_bg": "#3e3e42", "btn_fg": "#ffffff",
            "btn_active": "#505055", "selection": "#3584e4",
        },
        "light": {
            "bg_main": "#ffffff", "bg_panel": "#f5f5f5", "bg_input": "#ffffff",
            "fg_text": "#1a1a1a", "fg_comment": "#008000", "keyword": "#0000ff",
            "string": "#a31515", "number": "#098658", "builtin": "#795e26",
            "type": "#267f99", "operator": "#1a1a1a", "separator": "#cccccc",
            "accent_green": "#2e7d32", "accent_blue": "#1976d2",
            "accent_orange": "#ed6c02", "accent_red": "#d32f2f",
            "border": "#cccccc", "btn_bg": "#e8e8e8", "btn_fg": "#1a1a1a",
            "btn_active": "#d0d0d0", "selection": "#1976d2",
        }
    }

    VBA_KEYWORDS = [
        "As","Binary","ByRef","ByVal","Date","Else","Empty","Error","False","For",
        "Friend","Get","Input","Is","Len","Let","Lock","Me","Mid","New","Next",
        "Nothing","Null","On","Option","Optional","ParamArray","Print","Private",
        "Property","Public","Resume","Seek","Set","Static","Step","String","Then",
        "Time","To","True","WithEvents","And","Eqv","Imp","Not","Or","Xor",
        "Call","Case","Close","Const","Declare","Dim","Do","Each","ElseIf","End",
        "Enum","Erase","Event","Exit","Function","GoSub","GoTo","If","Implements",
        "In","Loop","LSet","Open","Preserve","RaiseEvent","ReDim","Rem","Return",
        "RSet","Select","Stop","Sub","Type","Unlock","Wend","While","With",
        "Write","Attribute","Global"
    ]
    VBA_BUILTIN_FUNCS = [
        "Abs","Array","Asc","Atn","CBool","CByte","CCur","CDate","CDbl","CInt",
        "CLng","CSng","CStr","CVar","Choose","Chr","Command","Cos","CreateObject",
        "CurDir","Date","DateAdd","DateDiff","DatePart","DateSerial","DateValue",
        "Day","DDB","Dir","DoEvents","Environ","EOF","Error","Exp","FileAttr",
        "FileDateTime","FileLen","Filter","Format","FormatCurrency","FormatDateTime",
        "FormatNumber","FormatPercent","FreeFile","FV","GetAllSettings","GetAttr",
        "GetObject","GetSetting","Hex","Hour","IIf","InputBox","InStr","InStrRev",
        "Int","IPmt","IRR","IsArray","IsDate","IsEmpty","IsError","IsMissing",
        "IsNull","IsNumeric","IsObject","Join","LBound","LCase","Left","Len",
        "Loc","LOF","Log","LTrim","Mid","Minute","MIRR","Month","MonthName",
        "MsgBox","Now","NPer","NPV","Oct","Partition","Pmt","PPmt","PV","QBColor",
        "Rate","Replace","RGB","Right","Rnd","Round","RTrim","Second","Seek",
        "Sgn","Shell","Sin","SLN","Space","Spc","Split","Sqr","Str","StrComp",
        "StrConv","String","StrReverse","Switch","SYD","Tab","Tan","Time",
        "Timer","TimeSerial","TimeValue","Trim","TypeName","UBound","UCase","Val",
        "VarType","Weekday","WeekdayName","Year","Execute","Eval"
    ]
    VBA_TYPES = [
        "Boolean","Byte","Currency","Date","Decimal","Double","Integer","Long",
        "LongLong","Object","Single","String","Variant"
    ]

    SUSPICIOUS_PATTERNS = {
        "high": {
            "description": "Высокая опасность",
            "patterns": [
                {"regex": r"CreateObject\s*\(\s*[\"']Shell\.Application[\"']\s*\)", "explanation": "Запуск произвольных программ через Shell.Application"},
                {"regex": r"CreateObject\s*\(\s*[\"']WScript\.Shell[\"']\s*\)", "explanation": "Создание WScript.Shell для выполнения команд ОС"},
                {"regex": r"CreateObject\s*\(\s*[\"']MSXML2\.XMLHTTP[\"']\s*\)", "explanation": "Загрузка данных из сети через MSXML2.XMLHTTP"},
                {"regex": r"\.Download\s+[\"']?http", "explanation": "Скачивание файла по HTTP"},
                {"regex": r"URLDownloadToFile", "explanation": "API-вызов URLDownloadToFile для скачивания файла"},
                {"regex": r"ShellExecute", "explanation": "Выполнение файла через ShellExecute"},
                {"regex": r"RegWrite\s+[\"']HKEY_", "explanation": "Запись в реестр (возможно, для автозагрузки)"},
                {"regex": r"Binary\.Write", "explanation": "Запись бинарных данных"},
                {"regex": r"ADODB\.Stream", "explanation": "Работа с ADODB.Stream для загрузки/сохранения файлов"},
                {"regex": r"SaveAs\s+.*\.exe", "explanation": "Сохранение файла с расширением .exe"},
                {"regex": r"Run\s*\(\s*[\"'].*\.exe", "explanation": "Запуск .exe файла"},
                {"regex": r"Exec\s*\(\s*[\"']", "explanation": "Выполнение команды через WScript.Shell.Exec"},
            ]
        },
        "medium": {
            "description": "Средняя опасность",
            "patterns": [
                {"regex": r"\bAutoOpen\b", "explanation": "Автозапуск макроса при открытии документа"},
                {"regex": r"\bDocument_Open\b", "explanation": "Макрос, срабатывающий при открытии документа Word"},
                {"regex": r"\bWorkbook_Open\b", "explanation": "Макрос, срабатывающий при открытии книги Excel"},
                {"regex": r"\bAutoExec\b", "explanation": "Автозапуск макроса (старое имя)"},
                {"regex": r"\bAuto_Open\b", "explanation": "Автоматический запуск при открытии файла (Excel 4.0)"},
                {"regex": r"Shell\s*\(", "explanation": "Вызов внешней программы через Shell"},
                {"regex": r"Kill\s+[\"'].*\.\*", "explanation": "Удаление файлов с расширением .exe"},
                {"regex": r"FileSystemObject", "explanation": "Работа с файловой системой через FSO"},
                {"regex": r"CreateTextFile", "explanation": "Создание текстового файла"},
                {"regex": r"Open\s+[\"'].*\.exe", "explanation": "Открытие/создание .exe файла"},
                {"regex": r"Environ\s*\(", "explanation": "Чтение переменных окружения"},
                {"regex": r"WScript\.Shell", "explanation": "Ссылка на WScript.Shell без CreateObject"},
                {"regex": r"SendKeys", "explanation": "Эмуляция нажатий клавиш"},
                {"regex": r"Call\s+.*\.(exe|bat|cmd)", "explanation": "Вызов процедуры с расширением исполняемого файла"},
            ]
        },
        "low": {
            "description": "Низкая опасность (обфускация)",
            "patterns": [
                {"regex": r"Chr\s*\(", "explanation": "Использование Chr() может быть частью обфускации строк"},
                {"regex": r"StrReverse\s*\(", "explanation": "StrReverse может использоваться для сокрытия строк"},
                {"regex": r"Replace\s*\(", "explanation": "Replace иногда применяется в обфускации"},
                {"regex": r"Mid\s*\(", "explanation": "Выделение подстроки – возможный приём обфускации"},
                {"regex": r"Evaluate\s*\(", "explanation": "Evaluate выполняет строку как выражение (eval)"},
                {"regex": r"Application\.Run\s*", "explanation": "Динамический вызов макроса"},
                {"regex": r"ThisWorkbook\.VBProject", "explanation": "Доступ к проекту VBA (самомодификация)"},
                {"regex": r"ActiveVBProject", "explanation": "Доступ к активному VBA-проекту"},
            ]
        }
    }

    def __init__(self, root):
        self.root = root
        self.root.title("VBA Extractor")
        self.root.geometry("1070x720")   # расширено для размещения кнопок
        self.root.minsize(950, 600)

        self.current_theme = "dark"
        self.colors = self.THEMES[self.current_theme]

        self.file_path = tk.StringVar()
        self.modules = []
        self.current_module = -1
        self.last_hash_value = ""
        self.hash_displayed = False
        self.scan_results = {}
        self._line_map = {}

        self.w = {}
        self._create_ui()
        self._apply_theme()
        self.show_welcome_message()

    def _create_ui(self):
        c = self.colors
        self.paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=4, bd=1, relief=tk.SOLID)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # Левая панель
        self.left_frame = tk.Frame(self.paned, width=220, padx=4, pady=4)
        self.paned.add(self.left_frame, minsize=200)
        self.w["left_frame"] = self.left_frame
        tk.Label(self.left_frame, text="Модули", font=("Segoe UI", 9, "bold"), anchor="w").pack(pady=(0, 4))
        self.w["lbl_modules"] = self.left_frame.winfo_children()[0]
        self.modules_listbox = tk.Listbox(
            self.left_frame, font=("Consolas", 9), selectmode=tk.SINGLE,
            exportselection=False, relief=tk.SOLID, bd=1
        )
        self.modules_listbox.pack(fill=tk.BOTH, expand=True)
        self.w["modules_listbox"] = self.modules_listbox
        self.modules_listbox.bind("<<ListboxSelect>>", self.on_module_select)

        # Правая панель
        self.right_frame = tk.Frame(self.paned, padx=6, pady=4)
        self.paned.add(self.right_frame)
        self.w["right_frame"] = self.right_frame

        # Верх: поле ввода + Обзор + Тема
        self.top_frame = tk.Frame(self.right_frame)
        self.top_frame.pack(pady=(0, 8), fill=tk.X)
        self.w["top_frame"] = self.top_frame

        self.file_entry = tk.Entry(self.top_frame, textvariable=self.file_path, width=52,
                                   font=("Consolas", 9), relief=tk.SOLID, bd=1)
        self.file_entry.pack(side=tk.LEFT, padx=6)
        self.w["file_entry"] = self.file_entry

        self.btn_browse = tk.Button(self.top_frame, text="Обзор...", command=self.select_file,
                                    font=("Segoe UI", 9), relief=tk.RAISED, bd=2)
        self.btn_browse.pack(side=tk.LEFT, padx=(0, 4))
        self.w["btn_browse"] = self.btn_browse

        self.btn_theme = tk.Button(self.top_frame, text="🌓 Тема", command=self.toggle_theme,
                                   font=("Segoe UI", 9), relief=tk.RAISED, bd=2)
        self.btn_theme.pack(side=tk.RIGHT)
        self.w["btn_theme"] = self.btn_theme

        # Кнопки действий
        self.btn_frame = tk.Frame(self.right_frame)
        self.btn_frame.pack(pady=(0, 8), fill=tk.X)
        self.w["btn_frame"] = self.btn_frame

        btn_opts = {"font": ("Segoe UI", 9), "relief": tk.RAISED, "bd": 2, "padx": 14, "pady": 4}

        self.btn_extract_vba = tk.Button(self.btn_frame, text="Извлечь VBA",
                                         command=self.extract_vba_threaded, **btn_opts)
        self.btn_extract_vba.pack(side=tk.LEFT, padx=3)
        self.w["btn_extract_vba"] = self.btn_extract_vba

        self.btn_scan = tk.Button(self.btn_frame, text="Поиск подозрительного кода",
                                  command=self.scan_all_modules_threaded, **btn_opts)
        self.btn_scan.pack_forget()
        self.w["btn_scan"] = self.btn_scan

        self.btn_extract_hash = tk.Button(self.btn_frame, text="Извлечь хэш пароля",
                                          command=self.extract_hash_threaded, **btn_opts)
        self.btn_extract_hash.pack(side=tk.LEFT, padx=3)
        self.w["btn_extract_hash"] = self.btn_extract_hash

        # Кнопка "Копировать хэш" теперь левее "Снять защиту VBA"
        self.btn_copy_hash = tk.Button(self.btn_frame, text="Копировать хэш",
                                       command=self.copy_hash_only, state=tk.DISABLED, **btn_opts)
        self.btn_copy_hash.pack_forget()
        self.w["btn_copy_hash"] = self.btn_copy_hash

        self.btn_unlock_vba = tk.Button(self.btn_frame, text="🔓 Снять защиту VBA",
                                        command=self.unlock_vba_threaded, **btn_opts)
        self.btn_unlock_vba.pack_forget()
        self.w["btn_unlock_vba"] = self.btn_unlock_vba

        self.progress = ttk.Progressbar(self.right_frame, mode='indeterminate')
        self.w["progress"] = self.progress

        # Область вывода с нумерацией строк
        self.plain_text = LineNumberedText(
            self.right_frame, theme_colors=self.colors, wrap='word'
        )
        self.active_text_widget = self.plain_text
        self.plain_text.pack(pady=(0, 6), fill=tk.BOTH, expand=True)

        self.w["output_text"] = self.plain_text
        self._configure_syntax_tags()

        # Нижние кнопки
        self.action_frame = tk.Frame(self.right_frame)
        self.action_frame.pack(pady=(0, 4), fill=tk.X)
        self.w["action_frame"] = self.action_frame
        action_opts = {"font": ("Segoe UI", 9), "relief": tk.RAISED, "bd": 2, "padx": 12, "pady": 3}

        self.w["btn_copy_mod"] = tk.Button(self.action_frame, text="Копировать модуль", command=self.copy_current_module, **action_opts)
        self.w["btn_copy_mod"].pack(side=tk.LEFT, padx=2)

        self.w["btn_copy_all"] = tk.Button(self.action_frame, text="Копировать всё", command=self.copy_all, **action_opts)
        self.w["btn_copy_all"].pack(side=tk.LEFT, padx=2)

        self.w["btn_save_mod"] = tk.Button(self.action_frame, text="Сохранить модуль", command=self.save_current_module, **action_opts)
        self.w["btn_save_mod"].pack(side=tk.LEFT, padx=2)

        self.w["btn_save_all"] = tk.Button(self.action_frame, text="Сохранить всё (ZIP)", command=self.save_all_to_zip, **action_opts)
        self.w["btn_save_all"].pack(side=tk.LEFT, padx=2)

        self.w["btn_show_all"] = tk.Button(self.action_frame, text="Показать все", command=self.show_all_modules, **action_opts)
        self.w["btn_show_all"].pack(side=tk.LEFT, padx=2)

        self.w["btn_open_txt"] = tk.Button(self.action_frame, text="Открыть в .txt", command=self.open_current_in_txt, **action_opts)
        self.w["btn_open_txt"].pack(side=tk.LEFT, padx=2)

        self.w["btn_clear"] = tk.Button(self.action_frame, text="Очистить", command=self.clear_output, **action_opts)
        self.w["btn_clear"].pack(side=tk.LEFT, padx=2)

        self.create_context_menu()
        self.setup_hotkeys()

    def _configure_syntax_tags(self):
        c = self.colors
        tags = {
            "keyword": {"foreground": c["keyword"], "font": ("Consolas", 10, "bold")},
            "string": {"foreground": c["string"]},
            "number": {"foreground": c["number"]},
            "builtin": {"foreground": c["builtin"]},
            "type": {"foreground": c["type"]},
            "comment": {"foreground": c["fg_comment"]},
            "separator": {"foreground": c["separator"], "font": ("Consolas", 8)},
            "operator": {"foreground": c["operator"]},
        }
        for tag, conf in tags.items():
            self.plain_text.tag_configure(tag, **conf)

        self.plain_text.tag_configure("bg_high", background="#8b0000", foreground="white")
        self.plain_text.tag_configure("bg_medium", background="#b85c00", foreground="white")
        self.plain_text.tag_configure("bg_low", background="#b8860b", foreground="white")

    def apply_syntax_highlighting(self):
        text_widget = self.plain_text
        content = text_widget.get("1.0", tk.END)
        for tag in ("keyword", "string", "number", "builtin", "type", "comment", "separator", "operator"):
            text_widget.tag_remove(tag, "1.0", tk.END)

        for match in re.finditer(r'\b\d+\.?\d*([eE][+-]?\d+)?\b', content):
            start_idx = f"1.0 + {match.start()} chars"
            end_idx = f"1.0 + {match.end()} chars"
            text_widget.tag_add("number", start_idx, end_idx)

        for kw in self.VBA_KEYWORDS:
            for match in re.finditer(r'\b' + re.escape(kw) + r'\b', content):
                start_idx = f"1.0 + {match.start()} chars"
                end_idx = f"1.0 + {match.end()} chars"
                text_widget.tag_add("keyword", start_idx, end_idx)

        for func in self.VBA_BUILTIN_FUNCS:
            for match in re.finditer(r'\b' + re.escape(func) + r'\b', content):
                start_idx = f"1.0 + {match.start()} chars"
                end_idx = f"1.0 + {match.end()} chars"
                text_widget.tag_add("builtin", start_idx, end_idx)

        for typ in self.VBA_TYPES:
            for match in re.finditer(r'\b' + re.escape(typ) + r'\b', content):
                start_idx = f"1.0 + {match.start()} chars"
                end_idx = f"1.0 + {match.end()} chars"
                text_widget.tag_add("type", start_idx, end_idx)

        for match in re.finditer(r'^─+$', content, re.MULTILINE):
            line_num = content[:match.start()].count('\n') + 1
            start_idx = f"{line_num}.0"
            end_idx = f"{line_num}.end"
            text_widget.tag_add("separator", start_idx, end_idx)

        for match in re.finditer(r'"(?:[^"]|"")*"', content):
            start_idx = f"1.0 + {match.start()} chars"
            end_idx = f"1.0 + {match.end()} chars"
            text_widget.tag_add("string", start_idx, end_idx)

        for match in re.finditer(r"(?:^|\s)('[^\n]*|Rem\s[^\n]*)", content, re.IGNORECASE | re.MULTILINE):
            start_idx = f"1.0 + {match.start()} chars"
            end_idx = f"1.0 + {match.end()} chars"
            text_widget.tag_add("comment", start_idx, end_idx)

        text_widget.tag_raise("string")
        text_widget.tag_raise("comment")

    def _insert_separators(self, code):
        text, _ = self._insert_separators_with_map(code)
        return text

    def _insert_separators_with_map(self, code):
        lines = code.splitlines()
        new_lines = []
        line_map = {}
        sep_line = "─" * 80
        widget_line = 1
        for code_idx, line in enumerate(lines):
            code_line_num = code_idx + 1
            new_lines.append(line)
            line_map[code_line_num] = widget_line
            widget_line += 1
            if re.match(r'^\s*End\s+(Sub|Function|Property)\b', line, re.IGNORECASE):
                if code_idx < len(lines) - 1:
                    new_lines.append(sep_line)
                    widget_line += 1
        return '\n'.join(new_lines), line_map

    def toggle_theme(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.colors = self.THEMES[self.current_theme]
        self._apply_theme()
        if self.modules or self.hash_displayed:
            self.apply_syntax_highlighting()
            if self.current_module >= 0 and self.scan_results:
                module_name = self.modules[self.current_module][0]
                if module_name in self.scan_results:
                    self._highlight_lines_in_current_view(self.scan_results[module_name])
                    self._scroll_to_first_suspicious(self.scan_results[module_name])

    def _apply_theme(self):
        c = self.colors
        self.root.configure(bg=c["bg_main"])
        self.w["left_frame"].configure(bg=c["bg_panel"])
        self.w["right_frame"].configure(bg=c["bg_main"])
        self.paned.configure(bg=c["border"])
        self.w["lbl_modules"].configure(bg=c["bg_panel"], fg=c["fg_text"])
        self.w["modules_listbox"].configure(
            bg=c["bg_input"], fg=c["fg_text"],
            selectbackground=c["selection"], selectforeground="white",
            highlightbackground=c["border"], highlightthickness=1
        )
        self.w["top_frame"].configure(bg=c["bg_main"])
        self.w["btn_frame"].configure(bg=c["bg_main"])
        self.w["action_frame"].configure(bg=c["bg_main"])
        self.w["file_entry"].configure(
            bg=c["bg_input"], fg=c["fg_text"], insertbackground=c["fg_text"],
            highlightbackground=c["border"], highlightthickness=1
        )

        btn_map = {
            "btn_browse": (c["btn_bg"], c["btn_fg"]),
            "btn_theme": (c["btn_bg"], c["btn_fg"]),
            "btn_extract_vba": (c["accent_green"], "white"),
            "btn_extract_hash": (c["accent_blue"], "white"),
            "btn_copy_hash": (c["accent_orange"], "white"),
            "btn_clear": (c["accent_red"], "white"),
        }
        for key, (bg, fg) in btn_map.items():
            if key in self.w:
                self.w[key].configure(bg=bg, fg=fg, activebackground=self._darken(bg, 30), activeforeground="white")

        if "btn_scan" in self.w:
            self.w["btn_scan"].configure(bg="#f1c40f", fg="black", activebackground="#f39c12", activeforeground="black")

        for key in ["btn_copy_mod", "btn_copy_all", "btn_save_mod", "btn_save_all", "btn_show_all", "btn_open_txt"]:
            if key in self.w:
                self.w[key].configure(bg=c["btn_bg"], fg=c["btn_fg"], activebackground=c["btn_active"])

        # Розовая кнопка снятия защиты
        if "btn_unlock_vba" in self.w and self.w["btn_unlock_vba"] is not None:
            self.w["btn_unlock_vba"].configure(
                bg="#e91e63", fg="white",
                activebackground="#c2185b", activeforeground="white"
            )

        self.plain_text.apply_theme(c)
        self._configure_syntax_tags()
        self.create_context_menu()

    def _darken(self, color, amount=30):
        if color.startswith("#") and len(color) == 7:
            r, g, b = int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)
            return f"#{max(0,r-amount):02x}{max(0,g-amount):02x}{max(0,b-amount):02x}"
        return color

    def setup_hotkeys(self):
        self.plain_text.bind("<Control-c>", self.copy_selection)
        self.plain_text.bind("<Control-v>", self.paste_text)
        self.plain_text.bind("<Control-a>", self.select_all)

    def copy_selection(self, event=None):
        widget = self.plain_text
        try:
            selected = widget.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected)
            self.root.update()
        except tk.TclError:
            pass
        return "break"

    def paste_text(self, event=None):
        widget = self.plain_text
        try:
            text = self.root.clipboard_get()
            widget.insert(tk.INSERT, text)
        except tk.TclError:
            pass
        return "break"

    def select_all(self, event=None):
        widget = self.plain_text
        widget.tag_add(tk.SEL, "1.0", tk.END)
        widget.mark_set(tk.INSERT, "1.0")
        widget.see(tk.INSERT)
        return "break"

    def show_welcome_message(self):
        text = """VBA Extractor — извлечение макросов и хэшей паролей

Поддерживаемые форматы: .xls, .xlsm, .xlsb, .xltm, .docm, .pptm, .xlam

Требования: Python 3.6+, pip install oletools olefile

Как использовать:
  1. Нажмите "Обзор" и выберите файл Excel/Word/PowerPoint с макросами
  2. Нажмите "Извлечь VBA" — код появится справа, модули — слева
  3. Нажмите "Извлечь хэш пароля" — получите хэш для подбора
  4. Кнопка "Копировать хэш" появится автоматически после извлечения
  5. После извлечения VBA станет доступна кнопка "Поиск подозрительного кода"
  6. Также станет доступна кнопка "Снять защиту VBA" (розовая) — убирает пароль с VBA-проекта

"""
        self._set_output_text(text, is_code=False)

    def _set_output_text(self, text, is_code=False):
        self.active_text_widget = self.plain_text
        self.plain_text.pack(pady=(0, 6), fill=tk.BOTH, expand=True)
        self.plain_text.delete(1.0, tk.END)
        self.plain_text.insert(1.0, text)
        self.plain_text.see("1.0")

        # Управление номерами строк и переносом
        if is_code:
            self.plain_text.enable_line_numbers(True)
            self.plain_text.configure_wrap('none')      # без переноса
        else:
            self.plain_text.enable_line_numbers(False)
            self.plain_text.configure_wrap('word')      # с переносом

    def _clear_output(self):
        self.plain_text.delete(1.0, tk.END)

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл с макросами",
            filetypes=[("Office files with macros", "*.xls *.xlsm *.xlsb *.xltm *.docm *.pptm *.xlam"),
                       ("All files", "*.*")])
        if filename:
            self.file_path.set(filename)
            self._reset_buttons()

    def _reset_buttons(self, hide_scan=True):
        self.w["btn_extract_vba"].config(state=tk.NORMAL, text="Извлечь VBA")
        self.w["btn_extract_hash"].config(state=tk.NORMAL, text="Извлечь хэш пароля")
        if hide_scan:
            self.w["btn_scan"].pack_forget()
            self.w["btn_unlock_vba"].pack_forget()
            self.w["btn_copy_hash"].pack_forget()
        else:
            if self.modules:
                self.w["btn_scan"].pack(side=tk.LEFT, padx=3, before=self.w["btn_extract_hash"])
                self.w["btn_unlock_vba"].pack(side=tk.LEFT, padx=3)
                # Кнопка "Копировать хэш" показывается только при наличии хэша
            else:
                self.w["btn_scan"].pack_forget()
                self.w["btn_unlock_vba"].pack_forget()
        self.w["btn_copy_hash"].config(state=tk.DISABLED)
        self.last_hash_value = ""
        self.hash_displayed = False
        self.scan_results = {}

    def start_progress(self):
        self.progress.pack(pady=(0, 6), fill=tk.X)
        self.progress.start(10)

    def stop_progress(self):
        self.progress.stop()
        self.progress.pack_forget()

    def extract_vba_threaded(self):
        if not self.file_path.get():
            messagebox.showwarning("Предупреждение", "Сначала выберите файл", parent=self.root)
            return
        self._reset_buttons()
        self.w["btn_extract_vba"].config(state=tk.DISABLED, text="Извлечение...")
        self.w["btn_extract_hash"].config(state=tk.DISABLED)
        self._clear_output()
        self.modules_listbox.delete(0, tk.END)
        self.start_progress()
        thread = threading.Thread(target=self.extract_vba)
        thread.daemon = True
        thread.start()

    def extract_vba(self):
        path = self.file_path.get()
        modules = []
        info = []
        try:
            info.append(f"Файл: {os.path.basename(path)}\n")
            info.append("Анализ через olevba...\n")
            vba_parser = VBA_Parser(path, relaxed=True)
            for _, _, vba_filename, vba_code in vba_parser.extract_macros():
                clean = clean_vba_code(vba_code)
                if clean:
                    modules.append((vba_filename, clean))
            if not modules:
                info.append("Макросы не найдены.\nПроверьте, что файл содержит VBA-проект.")
            vba_parser.close()
        except Exception as e:
            info.append(f"Ошибка: {str(e)}")

        self.modules = sorted(modules, key=lambda x: x[0].lower())
        self.extraction_info = "\n".join(info)
        self.root.after(0, self.on_extraction_complete)

    def on_extraction_complete(self):
        self.stop_progress()
        self.hash_displayed = False
        self.modules_listbox.delete(0, tk.END)
        self.current_module = -1
        self._clear_output()
        if not self.modules:
            self._set_output_text(self.extraction_info, is_code=False)
            messagebox.showinfo("Результат", "Макросы не найдены.", parent=self.root)
            self.w["btn_scan"].pack_forget()
            self.w["btn_unlock_vba"].pack_forget()
            self.w["btn_copy_hash"].pack_forget()
        else:
            for name, _ in self.modules:
                self.modules_listbox.insert(tk.END, name)
            self.modules_listbox.selection_set(0)
            self.current_module = 0
            self.display_module(0)
            # Показываем кнопки поиска и снятия защиты
            self.w["btn_scan"].pack(side=tk.LEFT, padx=3, before=self.w["btn_extract_hash"])
            self.w["btn_scan"].config(state=tk.NORMAL)
            # Кнопка "Копировать хэш" ещё не нужна, но место для неё зарезервируем
            self.w["btn_unlock_vba"].pack(side=tk.LEFT, padx=3)
            self.w["btn_unlock_vba"].config(state=tk.NORMAL)
            self.w["btn_copy_hash"].pack_forget()  # скрыта
        self.w["btn_extract_vba"].config(state=tk.NORMAL, text="Извлечь VBA")
        self.w["btn_extract_hash"].config(state=tk.NORMAL)

    def display_module(self, index):
        if 0 <= index < len(self.modules):
            name, code = self.modules[index]
            display_code, line_map = self._insert_separators_with_map(code)
            self._line_map = line_map
            self._set_output_text(display_code, is_code=True)
            self.apply_syntax_highlighting()
            if name in self.scan_results:
                self._highlight_lines_in_current_view(self.scan_results[name])
                self._scroll_to_first_suspicious(self.scan_results[name])
            self.hash_displayed = False

    def on_module_select(self, event):
        sel = self.modules_listbox.curselection()
        if sel:
            self.current_module = sel[0]
            self.display_module(sel[0])

    def show_all_modules(self):
        if not self.modules:
            self._set_output_text(self.extraction_info, is_code=False)
            self.apply_syntax_highlighting()
            return
        content = self.extraction_info + "\n" + "-"*50 + "\n"
        for name, code in self.modules:
            content += f"\n' Module: {name}\n"
            content += self._insert_separators(code) + "\n" + "-"*50 + "\n"
        self._set_output_text(content, is_code=False)
        self.apply_syntax_highlighting()

    def scan_all_modules_threaded(self):
        if not self.modules:
            messagebox.showwarning("Нет данных", "Сначала извлеките VBA-код.", parent=self.root)
            return
        self.start_progress()
        thread = threading.Thread(target=self.scan_all_modules)
        thread.daemon = True
        thread.start()

    def scan_all_modules(self):
        all_results = {}
        total_found = 0

        for name, code in self.modules:
            results = self._scan_code_for_patterns(code)
            if results:
                all_results[name] = results
                total_found += sum(len(v) for v in results.values())

        self.scan_results = all_results

        report_lines = []
        report_lines.append("="*80)
        report_lines.append("РЕЗУЛЬТАТЫ ПОИСКА ПОДОЗРИТЕЛЬНОГО КОДА (по всем модулям)")
        report_lines.append("="*80)
        if total_found == 0:
            report_lines.append("\n✅ Подозрительных конструкций не найдено.")
        else:
            report_lines.append(f"\n⚠ Найдено потенциально опасных мест: {total_found}\n")
            for mod_name, res in all_results.items():
                report_lines.append(f"\n📄 Модуль: {mod_name}")
                for level in ["high", "medium", "low"]:
                    items = res.get(level, [])
                    if items:
                        desc = self.SUSPICIOUS_PATTERNS[level]["description"]
                        report_lines.append(f"  🔴 {level.upper()} – {desc}:")
                        for item in items:
                            macro_name = item.get("macro", "неизвестный макрос")
                            report_lines.append(f"      строка {item['line']} (макрос {macro_name}): {item['explanation']} (шаблон: {item['pattern']})")
                report_lines.append("-"*40)

        self.root.after(0, self._display_scan_report, "\n".join(report_lines))
        self.root.after(0, self.stop_progress)

    def _display_scan_report(self, report_text):
        self._set_output_text(report_text, is_code=False)
        self.hash_displayed = True

    def _scan_code_for_patterns(self, code):
        results = {"high": [], "medium": [], "low": []}
        lines = code.splitlines()
        current_macro = None
        for i, line in enumerate(lines, start=1):
            match_macro = re.search(r'^\s*(?:Sub|Function)\s+(\w+)', line, re.IGNORECASE)
            if match_macro:
                current_macro = match_macro.group(1)
            for level in ["high", "medium", "low"]:
                for item in self.SUSPICIOUS_PATTERNS[level]["patterns"]:
                    if re.search(item["regex"], line, re.IGNORECASE):
                        results[level].append({
                            "line": i,
                            "macro": current_macro if current_macro else "глобальный",
                            "pattern": item["regex"],
                            "explanation": item["explanation"],
                            "line_text": line.strip()
                        })
                        break
        return {k: v for k, v in results.items() if v}

    def _highlight_lines_in_current_view(self, results):
        tw = self.plain_text
        for tag in ("bg_high", "bg_medium", "bg_low"):
            tw.tag_remove(tag, "1.0", tk.END)
        level_tag = {"high": "bg_high", "medium": "bg_medium", "low": "bg_low"}
        for level, items in results.items():
            tag = level_tag.get(level)
            if not tag:
                continue
            for item in items:
                code_line = item["line"]
                widget_line = self._line_map.get(code_line, code_line)
                tw.tag_add(tag, f"{widget_line}.0", f"{widget_line}.end")
        for tag in ("bg_high", "bg_medium", "bg_low"):
            tw.tag_raise(tag)

    def _scroll_to_first_suspicious(self, results):
        first_code_line = None
        for level in ["high", "medium", "low"]:
            for item in results.get(level, []):
                ln = item["line"]
                if first_code_line is None or ln < first_code_line:
                    first_code_line = ln
        if first_code_line is not None:
            self.plain_text.see_code_line(first_code_line)

    def extract_hash_threaded(self):
        if not self.file_path.get():
            messagebox.showwarning("Предупреждение", "Сначала выберите файл", parent=self.root)
            return
        self._reset_buttons(hide_scan=False)
        self.w["btn_extract_hash"].config(state=tk.DISABLED, text="Извлечение...")
        self.w["btn_extract_vba"].config(state=tk.DISABLED)
        self._clear_output()
        self.last_hash_value = ""
        self.start_progress()
        thread = threading.Thread(target=self.extract_hash)
        thread.daemon = True
        thread.start()

    def extract_hash(self):
        path = self.file_path.get()
        result = []
        temp_file = None
        hashes = {"old": None, "new": None}

        try:
            result.append(f"Хэш пароля для: {os.path.basename(path)}\n")
            result.append("="*60 + "\n\n")

            ole_data = path
            if path.lower().endswith(('.xlsm', '.xlsb', '.xltm', '.xlam', '.docm', '.pptm')):
                with zipfile.ZipFile(path, 'r') as z:
                    vba_bin = None
                    for cand in ['xl/vbaProject.bin', 'word/vbaProject.bin', 'ppt/vbaProject.bin', 'vbaProject.bin']:
                        if cand in z.namelist():
                            vba_bin = z.read(cand)
                            break
                    if vba_bin:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.bin') as tmp:
                            tmp.write(vba_bin)
                            ole_data = tmp.name
                            temp_file = ole_data
                    else:
                        self.root.after(0, self.display_result, "❌ Не найден vbaProject.bin в архиве.")
                        self.root.after(0, self.stop_progress)
                        return

            if not olefile.isOleFile(ole_data):
                self.root.after(0, self.display_result, "❌ Ошибка: не удалось прочитать OLE-структуру.")
                self.root.after(0, self.stop_progress)
                return

            ole = olefile.OleFileIO(ole_data)
            stream = None
            for candidate in ['VBA/PROJECT', '_VBA_PROJECT_CUR/PROJECT', 'PROJECT']:
                if ole.exists(candidate):
                    stream = candidate
                    break
            if not stream:
                self.root.after(0, self.display_result, "⚠ Поток PROJECT не найден. Возможно, проект не защищён.")
                ole.close()
                self.root.after(0, self.stop_progress)
                return

            data = ole.openstream(stream).read()
            ole.close()
            text = data.decode('latin-1', errors='ignore')

            dpb = re.search(r'(?:DPB|DPx)="([A-Fa-f0-9]+)"', text)
            cmg = re.search(r'CMG="([A-Fa-f0-9]+)"', text)
            gc  = re.search(r'GC="([A-Fa-f0-9]+)"', text)
            wid = re.search(r'ID="([A-Fa-f0-9]+)"', text)
            wep = re.search(r'WEP="([A-Fa-f0-9]+)"', text)

            if path.lower().endswith('.xls') and dpb:
                h = f"$oldoffice$0*{dpb.group(1)}"
                hashes["old"] = h
                result.append("🔐 СТАРЫЙ ФОРМАТ (Office 97–2010)\n")
                result.append("-"*40 + "\n")
                result.append(f"Хэш:\n  {h}\n\n")
                result.append("Использование с hashcat:\n")
                result.append("  Режим: -m 9700\n")
                result.append("  hashcat -m 9700 -a 0 hash.txt rockyou.txt\n\n")
                result.append("  Маска (6 символов):\n")
                result.append("  hashcat -m 9700 -a 3 hash.txt ?a?a?a?a?a?a\n\n")

            elif cmg and dpb and gc:
                h = f"$vba$*{cmg.group(1)}*{dpb.group(1)}*{gc.group(1)}"
                hashes["new"] = h
                result.append("🔐 НОВЫЙ ФОРМАТ (Office 2013–2024)\n")
                result.append("-"*40 + "\n")
                result.append(f"Хэш:\n  {h}\n\n")
                if cmg: result.append(f"  CMG : {cmg.group(1)}\n")
                if dpb: result.append(f"  DPB : {dpb.group(1)}\n")
                if gc:  result.append(f"  GC  : {gc.group(1)}\n")
                if wid: result.append(f"  ID  : {wid.group(1)}\n")
                if wep: result.append(f"  WEP : {wep.group(1)}\n")
                result.append("\nИспользование с hashcat:\n")
                result.append("  Режим: -m 29500 (VBA)\n")
                result.append("  hashcat -m 29500 hash.txt wordlist.txt\n\n")

            elif dpb:
                h = f"$partial$*{dpb.group(1)}"
                hashes["new"] = h
                result.append("⚠ ЧАСТИЧНЫЙ НОВЫЙ ФОРМАТ\n")
                result.append("-"*40 + "\n")
                result.append(f"{h}\n\n")
                result.append("Найдены не все поля защиты VBA-проекта.\n")

            if not hashes["old"] and not hashes["new"] and not hashes.get("partial"):
                result.append("Не найдено полей защиты (DPB, CMG, GC, ID, WEP).\n")
                result.append("Возможно, VBA-проект не защищён паролем.\n")
            else:
                self.last_hash_value = "\n".join(h for h in [hashes["old"], hashes["new"]] if h)

        except Exception as e:
            result.append(f"❌ Ошибка: {str(e)}")
        finally:
            if temp_file and os.path.exists(temp_file):
                try: os.unlink(temp_file)
                except: pass
            self.root.after(0, self.display_result, "".join(result))
            self.root.after(0, self._enable_hash_buttons)
            self.root.after(0, self.stop_progress)

    def _enable_hash_buttons(self):
        self.w["btn_extract_hash"].config(state=tk.NORMAL, text="Извлечь хэш пароля")
        self.w["btn_extract_vba"].config(state=tk.NORMAL)
        if self.last_hash_value:
            # Показываем кнопку "Копировать хэш" левее кнопки "Снять защиту VBA"
            # Если кнопка "Снять защиту" уже видна, вставим перед ней
            if self.w["btn_unlock_vba"].winfo_ismapped():
                self.w["btn_copy_hash"].pack(side=tk.LEFT, padx=3, before=self.w["btn_unlock_vba"])
            else:
                # Если кнопка "Снять защиту" скрыта, размещаем после кнопки "Извлечь хэш"
                self.w["btn_copy_hash"].pack(side=tk.LEFT, padx=3, before=self.w["btn_extract_hash"])
            self.w["btn_copy_hash"].config(state=tk.NORMAL)
            self.hash_displayed = True
        else:
            self.hash_displayed = False
        if self.modules and not self.w["btn_scan"].winfo_ismapped():
            self.w["btn_scan"].pack(side=tk.LEFT, padx=3, before=self.w["btn_extract_hash"])

    def display_result(self, text):
        self._set_output_text(text, is_code=False)
        self.apply_syntax_highlighting()

    def unlock_vba_threaded(self):
        if not self.file_path.get():
            messagebox.showwarning("Предупреждение", "Сначала выберите файл", parent=self.root)
            return
        self.w["btn_unlock_vba"].config(state=tk.DISABLED, text="Снятие защиты...")
        thread = threading.Thread(target=self.unlock_vba)
        thread.daemon = True
        thread.start()

    def unlock_vba(self):
        path = self.file_path.get()
        result = ["🔓 Попытка снять защиту VBA-проекта...\n", "="*60 + "\n\n"]
        output_path = self._get_unblocked_path(path)

        try:
            shutil.copy2(path, output_path)
            result.append(f"✅ Создан файл: {os.path.basename(output_path)}\n\n")

            ext = path.lower()
            if ext.endswith('.xls'):
                msg = self._unlock_old_xls(path, output_path)
            elif ext.endswith(('.xlsm', '.xlsb', '.xltm', '.xlam', '.docm', '.pptm')):
                msg = self._unlock_new_office(path, output_path)
            else:
                msg = "❌ Формат файла не поддерживается для автоматического снятия защиты."

            result.append(msg)

        except Exception as e:
            result.append(f"❌ Ошибка при снятии защиты: {str(e)}")
        finally:
            self.root.after(0, self._unlock_complete, "".join(result), output_path)

    def _get_unblocked_path(self, original_path):
        base, ext = os.path.splitext(original_path)
        return f"{base}_unblocked{ext}"

    def _unlock_old_xls(self, input_path, output_path):
        with open(input_path, "rb") as f:
            data = f.read()
        text = data.decode('latin-1', errors='ignore')
        match = re.search(r'(DPB=)([A-Fa-f0-9]+)', text, re.IGNORECASE)
        if not match:
            return "⚠ Поле DPB не найдено. Файл возможно не защищён паролем."
        new_text = text.replace(match.group(0), "DPx=" + match.group(2), 1)
        with open(output_path, "wb") as f:
            f.write(new_text.encode('latin-1'))
        return (
            "✅ Защита снята (DPB → DPx).\n\n"
            f"Создан файл: {os.path.basename(output_path)}\n\n"
            "Что делать дальше:\n"
            "  1. Откройте созданный файл в Excel\n"
            "  2. Если появится предупреждение о восстановлении проекта — нажмите ОК\n"
            "  3. Откройте редактор VBA: Alt+F11\n"
            "  4. Пароль запрашиваться не будет\n"
        )

    def _unlock_new_office(self, input_path, output_path):
        try:
            with zipfile.ZipFile(input_path, 'r') as z:
                namelist = z.namelist()
        except Exception as e:
            return f"❌ Не удалось открыть файл как ZIP: {e}"

        vba_path = None
        for candidate in ['xl/vbaProject.bin', 'word/vbaProject.bin',
                          'ppt/vbaProject.bin', 'vbaProject.bin']:
            if candidate in namelist:
                vba_path = candidate
                break
        if not vba_path:
            return "❌ Файл vbaProject.bin не найден внутри архива."

        with zipfile.ZipFile(input_path, 'r') as zf:
            vba_data = bytearray(zf.read(vba_path))

        if bytes(vba_data[:8]) != b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
            return "❌ vbaProject.bin не является корректным OLE-файлом."

        patch_result = self._patch_vba_project_bin_bytes(vba_data)
        if not patch_result["success"]:
            return f"❌ {patch_result['error']}"

        patched_vba_data = patch_result["data"]

        with zipfile.ZipFile(input_path, 'r') as z_src:
            with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as z_dst:
                for item in namelist:
                    if item == vba_path:
                        z_dst.writestr(item, bytes(patched_vba_data))
                    else:
                        z_dst.writestr(item, z_src.read(item))

        return (
            "✅ Защита снята.\n\n"
            f"Создан файл: {os.path.basename(output_path)}\n\n"
            f"Метод: удалены строки защиты из потока PROJECT ({patch_result['removed']} полей)\n\n"
            "Что делать дальше:\n"
            "  1. Откройте созданный файл в Excel/Word/PowerPoint\n"
            "  2. Если появится предупреждение о восстановлении проекта — нажмите ОК\n"
            "  3. Откройте редактор VBA: Alt+F11\n"
            "  4. Пароль запрашиваться не будет\n"
        )

    def _patch_vba_project_bin_bytes(self, vba_bytes):
        data = bytearray(vba_bytes)
        text = data.decode('latin-1', errors='replace')

        protection_patterns = [
            r'CMG="[A-Fa-f0-9]*"\r?\n',
            r'DPB="[A-Fa-f0-9]*"\r?\n',
            r'GC="[A-Fa-f0-9]*"\r?\n',
            r'WEP="[A-Fa-f0-9]*"\r?\n',
            r'ID="[A-Fa-f0-9]*"\r?\n',
            r'CMG=[A-Fa-f0-9]+\r?\n',
            r'DPB=[A-Fa-f0-9]+\r?\n',
            r'GC=[A-Fa-f0-9]+\r?\n',
            r'WEP=[A-Fa-f0-9]+\r?\n',
            r'ID=[A-Fa-f0-9]+\r?\n',
        ]

        removed = 0
        new_text = text
        seen_keys = set()
        for pat in protection_patterns:
            key = pat.split('=')[0]
            if key in seen_keys:
                continue
            new_candidate = re.sub(pat, '', new_text, flags=re.IGNORECASE)
            if new_candidate != new_text:
                removed += 1
                seen_keys.add(key)
                new_text = new_candidate

        if removed == 0:
            return {"success": False,
                    "error": "Поля защиты не найдены. Возможно, файл не защищён "
                             "или использует нестандартный формат."}

        new_bytes = bytearray(new_text.encode('latin-1', errors='replace'))
        orig_len = len(data)
        if len(new_bytes) < orig_len:
            new_bytes += b'\x00' * (orig_len - len(new_bytes))
        elif len(new_bytes) > orig_len:
            new_bytes = new_bytes[:orig_len]

        return {"success": True, "removed": removed, "data": new_bytes}

    def _unlock_complete(self, text, output_path):
        self.w["btn_unlock_vba"].config(state=tk.NORMAL, text="🔓 Снять защиту VBA")
        self._set_output_text(text, is_code=False)
        if os.path.exists(output_path):
            messagebox.showinfo("Готово",
                f"Файл создан:\n{os.path.basename(output_path)}\n\n"
                "Откройте его в Excel/Word/PowerPoint.\n"
                "Если появится предупреждение о восстановлении проекта — нажмите ОК.\n"
                "После этого зайдите в редактор VBA (Alt+F11) — пароль запрашиваться не будет.",
                parent=self.root)
        else:
            messagebox.showwarning("Предупреждение",
                "Файл не был создан. Подробности — в области вывода.", parent=self.root)

    def copy_hash_only(self):
        if not self.last_hash_value:
            messagebox.showwarning("Нет хэша", "Сначала извлеките хэш пароля", parent=self.root)
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.last_hash_value)
        self.root.update()
        messagebox.showinfo("Скопировано", "Хэш скопирован в буфер обмена.", parent=self.root)

    def copy_current_module(self):
        if 0 <= self.current_module < len(self.modules):
            code = self.modules[self.current_module][1]
            self.root.clipboard_clear()
            self.root.clipboard_append(code)
            self.root.update()
            messagebox.showinfo("Скопировано", f"Модуль '{self.modules[self.current_module][0]}' скопирован.", parent=self.root)
        else:
            messagebox.showwarning("Нет модуля", "Выберите модуль в списке.", parent=self.root)

    def copy_all(self):
        if self.hash_displayed:
            content = self.plain_text.get("1.0", tk.END).strip()
            if not content:
                messagebox.showwarning("Пусто", "Нет данных для копирования.", parent=self.root)
                return
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.root.update()
            messagebox.showinfo("Скопировано", "Информация о хэшах скопирована.", parent=self.root)
            return
        if not self.modules:
            messagebox.showwarning("Пусто", "Нет модулей для копирования.", parent=self.root)
            return
        all_code = [f"' Module: {name}\n{code}" for name, code in self.modules]
        text = "\n".join(all_code)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        messagebox.showinfo("Скопировано", "Все модули скопированы.", parent=self.root)

    def save_current_module(self):
        if 0 <= self.current_module < len(self.modules):
            name, code = self.modules[self.current_module]
            if '.' in name:
                base, ext = name.rsplit('.', 1)
            else:
                base, ext = name, 'bas'
            base = base.replace('.', '_')
            default = f"{base}.{ext}"
            path = filedialog.asksaveasfilename(
                defaultextension=f".{ext}",
                filetypes=[("VBA files", f"*.{ext}"), ("All files", "*.*")],
                initialfile=default,
                title="Сохранить модуль"
            )
            if path:
                try:
                    with open(path, "w", encoding="utf-8") as f: f.write(code)
                    messagebox.showinfo("Сохранено", f"Модуль сохранён:\n{path}", parent=self.root)
                except Exception as e: messagebox.showerror("Ошибка", f"Не удалось сохранить: {e}", parent=self.root)
        else: messagebox.showwarning("Нет модуля", "Выберите модуль.", parent=self.root)

    def save_all_to_zip(self):
        if not self.modules:
            messagebox.showwarning("Пусто", "Нет модулей для сохранения.", parent=self.root)
            return
        zip_path = filedialog.asksaveasfilename(
            defaultextension=".zip",
            filetypes=[("ZIP archive", "*.zip"), ("All files", "*.*")],
            initialfile="VBA_Modules.zip",
            title="Сохранить все модули в ZIP"
        )
        if not zip_path:
            return
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for name, code in self.modules:
                    if '.' in name:
                        base, ext = name.rsplit('.', 1)
                    else:
                        base, ext = name, 'bas'
                    base = base.replace('.', '_')
                    safe_name = f"{base}.{ext}"
                    zf.writestr(safe_name, code)
            messagebox.showinfo("Сохранено", f"Все модули сохранены в ZIP:\n{zip_path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить ZIP: {e}", parent=self.root)

    def open_current_in_txt(self):
        text = self.plain_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Пусто", "Нет данных для открытия.", parent=self.root)
            return
        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as tmp:
                tmp.write(text)
                tmp_path = tmp.name
            if platform.system() == 'Windows':
                os.startfile(tmp_path)
            elif platform.system() == 'Darwin':
                subprocess.Popen(['open', tmp_path])
            else:
                subprocess.Popen(['xdg-open', tmp_path])
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл: {e}", parent=self.root)

    def clear_output(self):
        self._clear_output()
        self.modules_listbox.delete(0, tk.END)
        self.modules = []
        self.current_module = -1
        self._reset_buttons()
        self.show_welcome_message()

    def create_context_menu(self):
        c = self.colors
        if hasattr(self, 'context_menu'): self.context_menu.destroy()
        self.context_menu = tk.Menu(self.root, tearoff=0, bg=c["bg_panel"], fg=c["fg_text"],
                                    activebackground=c["selection"], activeforeground="white")
        self.context_menu.add_command(label="Копировать", command=self.copy_selection)
        self.context_menu.add_command(label="Вставить", command=self.paste_text)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Копировать хэш", command=self.copy_hash_only)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Очистить", command=self.clear_output)

        def show_menu(event):
            self.context_menu.post(event.x_root, event.y_root)
        self.plain_text.bind("<Button-3>", show_menu)


if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleVBAExtractor(root)
    root.mainloop()
