import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import os
import re
import zipfile
import tempfile
import subprocess
import platform

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


class SimpleVBAExtractor:
    THEMES = {
        "dark": {
            "bg_main": "#1e1e1e",
            "bg_panel": "#252526",
            "bg_input": "#3c3c3c",
            "fg_text": "#d4d4d4",
            "fg_comment": "#6a9955",
            "keyword": "#569cd6",
            "string": "#ce9178",
            "number": "#b5cea8",
            "builtin": "#dcdcaa",
            "type": "#4ec9b0",
            "operator": "#d4d4d4",
            "separator": "#555555",
            "accent_green": "#4e9a06",
            "accent_blue": "#3584e4",
            "accent_orange": "#cd9309",
            "accent_red": "#c01c28",
            "border": "#3e3e42",
            "btn_bg": "#3e3e42",
            "btn_fg": "#ffffff",
            "btn_active": "#505055",
            "selection": "#3584e4",
        },
        "light": {
            "bg_main": "#ffffff",
            "bg_panel": "#f5f5f5",
            "bg_input": "#ffffff",
            "fg_text": "#1a1a1a",
            "fg_comment": "#008000",
            "keyword": "#0000ff",
            "string": "#a31515",
            "number": "#098658",
            "builtin": "#795e26",
            "type": "#267f99",
            "operator": "#1a1a1a",
            "separator": "#cccccc",
            "accent_green": "#2e7d32",
            "accent_blue": "#1976d2",
            "accent_orange": "#ed6c02",
            "accent_red": "#d32f2f",
            "border": "#cccccc",
            "btn_bg": "#e8e8e8",
            "btn_fg": "#1a1a1a",
            "btn_active": "#d0d0d0",
            "selection": "#1976d2",
        }
    }

    VBA_KEYWORDS = [
        "As", "Binary", "ByRef", "ByVal", "Date", "Else", "Empty", "Error", "False", "For",
        "Friend", "Get", "Input", "Is", "Len", "Let", "Lock", "Me", "Mid", "New", "Next",
        "Nothing", "Null", "On", "Option", "Optional", "ParamArray", "Print", "Private",
        "Property", "Public", "Resume", "Seek", "Set", "Static", "Step", "String", "Then",
        "Time", "To", "True", "WithEvents", "And", "Eqv", "Imp", "Not", "Or", "Xor",
        "Call", "Case", "Close", "Const", "Declare", "Dim", "Do", "Each", "ElseIf", "End",
        "Enum", "Erase", "Event", "Exit", "Function", "GoSub", "GoTo", "If", "Implements",
        "In", "Loop", "LSet", "Open", "Preserve", "RaiseEvent", "ReDim", "Rem", "Return",
        "RSet", "Select", "Stop", "Sub", "Type", "Unlock", "Wend", "While", "With",
        "Write", "Attribute", "Global"
    ]
    VBA_BUILTIN_FUNCS = [
        "Abs", "Array", "Asc", "Atn", "CBool", "CByte", "CCur", "CDate", "CDbl", "CInt",
        "CLng", "CSng", "CStr", "CVar", "Choose", "Chr", "Command", "Cos", "CreateObject",
        "CurDir", "Date", "DateAdd", "DateDiff", "DatePart", "DateSerial", "DateValue",
        "Day", "DDB", "Dir", "DoEvents", "Environ", "EOF", "Error", "Exp", "FileAttr",
        "FileDateTime", "FileLen", "Filter", "Format", "FormatCurrency", "FormatDateTime",
        "FormatNumber", "FormatPercent", "FreeFile", "FV", "GetAllSettings", "GetAttr",
        "GetObject", "GetSetting", "Hex", "Hour", "IIf", "InputBox", "InStr", "InStrRev",
        "Int", "IPmt", "IRR", "IsArray", "IsDate", "IsEmpty", "IsError", "IsMissing",
        "IsNull", "IsNumeric", "IsObject", "Join", "LBound", "LCase", "Left", "Len",
        "Loc", "LOF", "Log", "LTrim", "Mid", "Minute", "MIRR", "Month", "MonthName",
        "MsgBox", "Now", "NPer", "NPV", "Oct", "Partition", "Pmt", "PPmt", "PV", "QBColor",
        "Rate", "Replace", "RGB", "Right", "Rnd", "Round", "RTrim", "Second", "Seek",
        "Sgn", "Shell", "Sin", "SLN", "Space", "Spc", "Split", "Sqr", "Str", "StrComp",
        "StrConv", "String", "StrReverse", "Switch", "SYD", "Tab", "Tan", "Time",
        "Timer", "TimeSerial", "TimeValue", "Trim", "TypeName", "UBound", "UCase", "Val",
        "VarType", "Weekday", "WeekdayName", "Year", "Execute", "Eval"
    ]
    VBA_TYPES = [
        "Boolean", "Byte", "Currency", "Date", "Decimal", "Double", "Integer", "Long",
        "LongLong", "Object", "Single", "String", "Variant"
    ]

    def __init__(self, root):
        self.root = root
        self.root.title("VBA Extractor")
        self.root.geometry("1050x720")
        self.root.minsize(880, 600)

        self.current_theme = "dark"
        self.colors = self.THEMES[self.current_theme]

        self.file_path = tk.StringVar()
        self.modules = []
        self.current_module = -1
        self.last_hash_value = ""
        self.hash_displayed = False

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

        # Верх: поле ввода файла + кнопка Обзор + тема
        self.top_frame = tk.Frame(self.right_frame)
        self.top_frame.pack(pady=(0, 8), fill=tk.X)
        self.w["top_frame"] = self.top_frame

        # Только поле ввода и кнопка "Обзор"
        self.file_entry = tk.Entry(self.top_frame, textvariable=self.file_path, width=52, font=("Consolas", 9), relief=tk.SOLID, bd=1)
        self.file_entry.pack(side=tk.LEFT, padx=6)
        self.w["file_entry"] = self.file_entry

        self.btn_browse = tk.Button(self.top_frame, text="Обзор...", command=self.select_file, font=("Segoe UI", 9), relief=tk.RAISED, bd=2)
        self.btn_browse.pack(side=tk.LEFT, padx=(0, 4))
        self.w["btn_browse"] = self.btn_browse

        self.btn_theme = tk.Button(self.top_frame, text="🌓 Тема", command=self.toggle_theme, font=("Segoe UI", 9), relief=tk.RAISED, bd=2)
        self.btn_theme.pack(side=tk.RIGHT)
        self.w["btn_theme"] = self.btn_theme

        # Кнопки действий
        self.btn_frame = tk.Frame(self.right_frame)
        self.btn_frame.pack(pady=(0, 8), fill=tk.X)
        self.w["btn_frame"] = self.btn_frame
        btn_opts = {"font": ("Segoe UI", 9), "relief": tk.RAISED, "bd": 2, "padx": 14, "pady": 4}
        self.btn_extract_vba = tk.Button(self.btn_frame, text="Извлечь VBA", command=self.extract_vba_threaded, **btn_opts)
        self.btn_extract_vba.pack(side=tk.LEFT, padx=3)
        self.w["btn_extract_vba"] = self.btn_extract_vba
        self.btn_extract_hash = tk.Button(self.btn_frame, text="Извлечь хэш пароля", command=self.extract_hash_threaded, **btn_opts)
        self.btn_extract_hash.pack(side=tk.LEFT, padx=3)
        self.w["btn_extract_hash"] = self.btn_extract_hash
        self.btn_copy_hash = tk.Button(self.btn_frame, text="Копировать хэш", command=self.copy_hash_only, state=tk.DISABLED, **btn_opts)
        self.btn_copy_hash.pack_forget()
        self.w["btn_copy_hash"] = self.btn_copy_hash

        self.progress = ttk.Progressbar(self.right_frame, mode='indeterminate')
        self.w["progress"] = self.progress

        # Область вывода
        self.output_text = scrolledtext.ScrolledText(
            self.right_frame, wrap=tk.WORD, font=("Consolas", 10),
            relief=tk.SOLID, bd=1, padx=8, pady=6
        )
        self.output_text.pack(pady=(0, 6), fill=tk.BOTH, expand=True)
        self.w["output_text"] = self.output_text
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

    # ---------- Подсветка синтаксиса VBA ----------
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
            self.output_text.tag_configure(tag, **conf)

    def apply_syntax_highlighting(self):
        text_widget = self.output_text
        content = text_widget.get("1.0", tk.END)
        for tag in ("keyword", "string", "number", "builtin", "type", "comment", "separator", "operator"):
            text_widget.tag_remove(tag, "1.0", tk.END)

        # Числа
        for match in re.finditer(r'\b\d+\.?\d*([eE][+-]?\d+)?\b', content):
            start_idx = f"1.0 + {match.start()} chars"
            end_idx = f"1.0 + {match.end()} chars"
            text_widget.tag_add("number", start_idx, end_idx)

        # Ключевые слова
        for kw in self.VBA_KEYWORDS:
            for match in re.finditer(r'\b' + re.escape(kw) + r'\b', content):
                start_idx = f"1.0 + {match.start()} chars"
                end_idx = f"1.0 + {match.end()} chars"
                text_widget.tag_add("keyword", start_idx, end_idx)

        # Встроенные функции
        for func in self.VBA_BUILTIN_FUNCS:
            for match in re.finditer(r'\b' + re.escape(func) + r'\b', content):
                start_idx = f"1.0 + {match.start()} chars"
                end_idx = f"1.0 + {match.end()} chars"
                text_widget.tag_add("builtin", start_idx, end_idx)

        # Типы данных
        for typ in self.VBA_TYPES:
            for match in re.finditer(r'\b' + re.escape(typ) + r'\b', content):
                start_idx = f"1.0 + {match.start()} chars"
                end_idx = f"1.0 + {match.end()} chars"
                text_widget.tag_add("type", start_idx, end_idx)

        # Разделители (линия из тире)
        for match in re.finditer(r'^─+$', content, re.MULTILINE):
            line_num = content[:match.start()].count('\n') + 1
            start_idx = f"{line_num}.0"
            end_idx = f"{line_num}.end"
            text_widget.tag_add("separator", start_idx, end_idx)

        # Строки (с поддержкой экранирования "")
        for match in re.finditer(r'"(?:[^"]|"")*"', content):
            start_idx = f"1.0 + {match.start()} chars"
            end_idx = f"1.0 + {match.end()} chars"
            text_widget.tag_add("string", start_idx, end_idx)

        # Комментарии
        for match in re.finditer(r"(?:^|\s)('[^\n]*|Rem\s[^\n]*)", content, re.IGNORECASE | re.MULTILINE):
            start_idx = f"1.0 + {match.start()} chars"
            end_idx = f"1.0 + {match.end()} chars"
            text_widget.tag_add("comment", start_idx, end_idx)

        # Поднимаем строки и комментарии выше всех остальных тегов
        text_widget.tag_raise("string")
        text_widget.tag_raise("comment")

    def _insert_separators(self, code):
        lines = code.splitlines()
        new_lines = []
        sep_line = "─" * 80
        for i, line in enumerate(lines):
            new_lines.append(line)
            if re.match(r'^\s*End\s+(Sub|Function|Property)\b', line, re.IGNORECASE):
                if i < len(lines) - 1:
                    new_lines.append(sep_line)
        return '\n'.join(new_lines)

    # ================== Тема ==================
    def toggle_theme(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.colors = self.THEMES[self.current_theme]
        self._apply_theme()
        if self.modules or self.hash_displayed:
            self.apply_syntax_highlighting()

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

        for key in ["btn_copy_mod", "btn_copy_all", "btn_save_mod", "btn_save_all", "btn_show_all", "btn_open_txt"]:
            if key in self.w:
                self.w[key].configure(bg=c["btn_bg"], fg=c["btn_fg"], activebackground=c["btn_active"])

        self.w["output_text"].configure(
            bg=c["bg_main"], fg=c["fg_text"], insertbackground=c["fg_text"],
            highlightbackground=c["border"], highlightthickness=1,
            selectbackground=c["selection"], selectforeground="white"
        )
        self._configure_syntax_tags()
        self.create_context_menu()

    def _darken(self, color, amount=30):
        if color.startswith("#") and len(color) == 7:
            r, g, b = int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)
            return f"#{max(0,r-amount):02x}{max(0,g-amount):02x}{max(0,b-amount):02x}"
        return color

    # ================== Горячие клавиши ==================
    def setup_hotkeys(self):
        self.w["output_text"].bind("<Control-c>", self.copy_selection)
        self.w["output_text"].bind("<Control-v>", self.paste_text)
        self.w["output_text"].bind("<Control-a>", self.select_all)

    def copy_selection(self, event=None):
        try:
            selected = self.w["output_text"].get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected)
            self.root.update()
        except tk.TclError:
            pass
        return "break"

    def paste_text(self, event=None):
        try:
            text = self.root.clipboard_get()
            self.w["output_text"].insert(tk.INSERT, text)
        except tk.TclError:
            pass
        return "break"

    def select_all(self, event=None):
        self.w["output_text"].tag_add(tk.SEL, "1.0", tk.END)
        self.w["output_text"].mark_set(tk.INSERT, "1.0")
        self.w["output_text"].see(tk.INSERT)
        return "break"

    # ================== Приветствие ==================
    def show_welcome_message(self):
        text = """VBA Extractor — извлечение макросов и хэшей паролей

Поддерживаемые форматы: .xls, .xlsm, .xlsb, .xltm, .docm, .pptm, .xlam

Требования: Python 3.6+, pip install oletools olefile

Как использовать:
  1. Нажмите "Обзор" и выберите файл Excel/Word/PowerPoint с макросами
  2. Нажмите "Извлечь VBA" — код появится справа, модули — слева
  3. Нажмите "Извлечь хэш пароля" — получите хэш для подбора
  4. Кнопка "Копировать хэш" появится автоматически после извлечения

"""
        self.w["output_text"].delete(1.0, tk.END)
        self.w["output_text"].insert(tk.END, text)
        self.w["output_text"].see("1.0")
        self.apply_syntax_highlighting()

    # ================== Выбор файла ==================
    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл с макросами",
            filetypes=[("Office files with macros", "*.xls *.xlsm *.xlsb *.xltm *.docm *.pptm *.xlam"),
                       ("All files", "*.*")])
        if filename:
            self.file_path.set(filename)
            self._reset_buttons()

    def _reset_buttons(self):
        self.w["btn_extract_vba"].config(state=tk.NORMAL, text="Извлечь VBA")
        self.w["btn_extract_hash"].config(state=tk.NORMAL, text="Извлечь хэш пароля")
        self.w["btn_copy_hash"].pack_forget()
        self.w["btn_copy_hash"].config(state=tk.DISABLED)
        self.last_hash_value = ""
        self.hash_displayed = False

    # ================== Прогресс-бар ==================
    def start_progress(self):
        self.progress.pack(pady=(0, 6), fill=tk.X)
        self.progress.start(10)

    def stop_progress(self):
        self.progress.stop()
        self.progress.pack_forget()

    # ================== Извлечение VBA ==================
    def extract_vba_threaded(self):
        if not self.file_path.get():
            messagebox.showwarning("Предупреждение", "Сначала выберите файл", parent=self.root)
            return
        self._reset_buttons()
        self.w["btn_extract_vba"].config(state=tk.DISABLED, text="Извлечение...")
        self.w["btn_extract_hash"].config(state=tk.DISABLED)
        self.w["output_text"].delete(1.0, tk.END)
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
        self.w["output_text"].delete(1.0, tk.END)
        if not self.modules:
            self.w["output_text"].insert(tk.END, self.extraction_info)
            messagebox.showinfo("Результат", "Макросы не найдены.", parent=self.root)
        else:
            for name, _ in self.modules:
                self.modules_listbox.insert(tk.END, name)
            self.modules_listbox.selection_set(0)
            self.current_module = 0
            self.display_module(0)
        self.w["btn_extract_vba"].config(state=tk.NORMAL, text="Извлечь VBA")
        self.w["btn_extract_hash"].config(state=tk.NORMAL)

    def display_module(self, index):
        if 0 <= index < len(self.modules):
            _, code = self.modules[index]
            display_code = self._insert_separators(code)
            self.w["output_text"].delete(1.0, tk.END)
            self.w["output_text"].insert(tk.END, display_code)
            self.apply_syntax_highlighting()
            self.w["output_text"].see("1.0")
            self.hash_displayed = False

    def on_module_select(self, event):
        sel = self.modules_listbox.curselection()
        if sel:
            self.current_module = sel[0]
            self.display_module(sel[0])

    def show_all_modules(self):
        self.w["output_text"].delete(1.0, tk.END)
        self.hash_displayed = False
        if not self.modules:
            self.w["output_text"].insert(tk.END, self.extraction_info)
            self.apply_syntax_highlighting()
            return
        self.w["output_text"].insert(tk.END, self.extraction_info + "\n" + "-"*50 + "\n")
        for name, code in self.modules:
            self.w["output_text"].insert(tk.END, f"\n' Module: {name}\n")
            self.w["output_text"].insert(tk.END, self._insert_separators(code) + "\n" + "-"*50 + "\n")
        self.apply_syntax_highlighting()
        self.w["output_text"].see("1.0")

    # ================== Извлечение хэша ==================
    def extract_hash_threaded(self):
        if not self.file_path.get():
            messagebox.showwarning("Предупреждение", "Сначала выберите файл", parent=self.root)
            return
        self._reset_buttons()
        self.w["btn_extract_hash"].config(state=tk.DISABLED, text="Извлечение...")
        self.w["btn_extract_vba"].config(state=tk.DISABLED)
        self.w["output_text"].delete(1.0, tk.END)
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
                result.append("Обнаруженные поля:\n")
                if cmg: result.append(f"  CMG : {cmg.group(1)}\n")
                if dpb: result.append(f"  DPB : {dpb.group(1)}\n")
                if gc:  result.append(f"  GC  : {gc.group(1)}\n")
                if wid: result.append(f"  ID  : {wid.group(1)}\n")
                if wep: result.append(f"  WEP : {wep.group(1)}\n")
                result.append("\nЧто означает каждый параметр:\n")
                result.append("  CMG — encrypted state\n")
                result.append("  DPB — password verifier\n")
                result.append("  GC  — project constants\n")
                result.append("  ID  — project identifier\n")
                result.append("  WEP — encryption flags\n")
                result.append("\nИспользование с hashcat:\n")
                result.append("  Режим: -m 29500 (VBA)\n")
                result.append("  John the Ripper (office2john) также поддерживает этот формат.\n\n")
                result.append("Пример:\n")
                result.append("  office2john.py file.xlsm > hash.txt\n")
                result.append("  hashcat -m 29500 hash.txt wordlist.txt\n\n")

            elif dpb:
                h = f"$partial$*{dpb.group(1)}"
                hashes["new"] = h
                result.append("⚠ ЧАСТИЧНЫЙ НОВЫЙ ФОРМАТ\n")
                result.append("-"*40 + "\n")
                result.append(f"{h}\n\n")
                result.append("Найдены не все поля защиты VBA-проекта.\n")
                result.append("Для полного modern hash обычно нужны: CMG + DPB + GC\n\n")

            if not hashes["old"] and not hashes["new"] and not hashes.get("partial"):
                result.append("Не найдено полей защиты (DPB, CMG, GC, ID, WEP).\n")
                result.append("Возможно, VBA-проект не защищён паролем или использует нестандартную защиту.\n")
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
            self.w["btn_copy_hash"].pack(side=tk.LEFT, padx=3)
            self.w["btn_copy_hash"].config(bg=self.colors["accent_orange"], fg="white")
            self.w["btn_copy_hash"].config(state=tk.NORMAL)
            self.hash_displayed = True
        else:
            self.hash_displayed = False

    def display_result(self, text):
        self.w["output_text"].insert(tk.END, text)
        self.w["output_text"].see("1.0")
        self.apply_syntax_highlighting()

    # ================== Копирование / сохранение ==================
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
            content = self.w["output_text"].get("1.0", tk.END).strip()
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
            default = name.replace(".", "_") + ".bas"
            path = filedialog.asksaveasfilename(
                defaultextension=".bas",
                filetypes=[("VBA modules", "*.bas"), ("All files", "*.*")],
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
                    safe_name = name.replace(".", "_") + ".bas"
                    zf.writestr(safe_name, code)
            messagebox.showinfo("Сохранено", f"Все модули сохранены в ZIP:\n{zip_path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить ZIP: {e}", parent=self.root)

    def open_current_in_txt(self):
        text = self.w["output_text"].get("1.0", tk.END).strip()
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

    # ================== Очистка ==================
    def clear_output(self):
        self.w["output_text"].delete(1.0, tk.END)
        self.modules_listbox.delete(0, tk.END)
        self.modules = []
        self.current_module = -1
        self._reset_buttons()
        self.show_welcome_message()

    # ================== Контекстное меню ==================
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
        self.w["output_text"].bind("<Button-3>", show_menu)


if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleVBAExtractor(root)
    root.mainloop()
