import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import os
import re
import zipfile
import tempfile

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
        self.hash_displayed = False          # флаг: сейчас показана информация о хэше

        # Ссылки на все виджеты для надежной смены темы
        self.w = {}

        self._create_ui()
        self._apply_theme()
        self.show_welcome_message()

    def _create_ui(self):
        c = self.colors
        # === Главный разделитель ===
        self.paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=4, bd=1, relief=tk.SOLID)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # === Левая панель ===
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

        # === Правая панель ===
        self.right_frame = tk.Frame(self.paned, padx=6, pady=4)
        self.paned.add(self.right_frame)
        self.w["right_frame"] = self.right_frame

        # --- Верх: файл + тема ---
        self.top_frame = tk.Frame(self.right_frame)
        self.top_frame.pack(pady=(0, 8), fill=tk.X)
        self.w["top_frame"] = self.top_frame

        tk.Label(self.top_frame, text="Файл:", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.file_entry = tk.Entry(self.top_frame, textvariable=self.file_path, width=52, font=("Consolas", 9), relief=tk.SOLID, bd=1)
        self.file_entry.pack(side=tk.LEFT, padx=6)
        self.w["file_entry"] = self.file_entry

        self.btn_browse = tk.Button(self.top_frame, text="Обзор...", command=self.select_file, font=("Segoe UI", 9), relief=tk.RAISED, bd=2)
        self.btn_browse.pack(side=tk.LEFT, padx=(0, 4))
        self.w["btn_browse"] = self.btn_browse

        self.btn_theme = tk.Button(self.top_frame, text="🌓 Тема", command=self.toggle_theme, font=("Segoe UI", 9), relief=tk.RAISED, bd=2)
        self.btn_theme.pack(side=tk.RIGHT)
        self.w["btn_theme"] = self.btn_theme

        # --- Кнопки действий ---
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
        # Скрыта по умолчанию
        self.btn_copy_hash.pack_forget()
        self.w["btn_copy_hash"] = self.btn_copy_hash

        # --- Область вывода ---
        self.output_text = scrolledtext.ScrolledText(
            self.right_frame, wrap=tk.WORD, font=("Consolas", 10),
            relief=tk.SOLID, bd=1, padx=8, pady=6
        )
        self.output_text.pack(pady=(0, 6), fill=tk.BOTH, expand=True)
        self.w["output_text"] = self.output_text

        self.output_text.tag_configure("comment", foreground=c["fg_comment"])
        self.output_text.tag_configure("hash_line", foreground=c["accent_blue"], font=("Consolas", 10, "bold"))
        self.output_text.tag_configure("cmd", foreground="#569cd6")
        self.output_text.tag_configure("section", font=("Segoe UI", 10, "bold"))

        # --- Нижние кнопки ---
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
        self.w["btn_save_all"] = tk.Button(self.action_frame, text="Сохранить всё", command=self.save_all_modules, **action_opts)
        self.w["btn_save_all"].pack(side=tk.LEFT, padx=2)
        self.w["btn_show_all"] = tk.Button(self.action_frame, text="Показать все", command=self.show_all_modules, **action_opts)
        self.w["btn_show_all"].pack(side=tk.LEFT, padx=2)
        self.w["btn_clear"] = tk.Button(self.action_frame, text="Очистить", command=self.clear_output, **action_opts)
        self.w["btn_clear"].pack(side=tk.LEFT, padx=2)

        self.create_context_menu()
        self.setup_hotkeys()

    # ================== Тема ==================
    def toggle_theme(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.colors = self.THEMES[self.current_theme]
        self._apply_theme()
        self.highlight_comments()

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

        # Кнопки
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

        for key in ["btn_copy_mod", "btn_copy_all", "btn_save_mod", "btn_save_all", "btn_show_all"]:
            if key in self.w:
                self.w[key].configure(bg=c["btn_bg"], fg=c["btn_fg"], activebackground=c["btn_active"])

        self.w["output_text"].configure(
            bg=c["bg_main"], fg=c["fg_text"], insertbackground=c["fg_text"],
            highlightbackground=c["border"], highlightthickness=1,
            selectbackground=c["selection"], selectforeground="white"
        )
        self.w["output_text"].tag_configure("comment", foreground=c["fg_comment"])
        self.w["output_text"].tag_configure("hash_line", foreground=c["accent_blue"])
        self.w["output_text"].tag_configure("cmd", foreground="#569cd6")

        # Контекстное меню (пересоздаём с актуальными цветами)
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

    # ================== Подсветка комментариев ==================
    def highlight_comments(self):
        self.w["output_text"].tag_remove("comment", "1.0", tk.END)
        text = self.w["output_text"].get("1.0", tk.END)
        lines = text.splitlines()
        line_num = 1
        for line in lines:
            match = re.search(r"(?:'|\b[Rr][Ee][Mm]\b)", line)
            if match:
                start = match.start()
                self.w["output_text"].tag_add("comment", f"{line_num}.{start}", f"{line_num}.end")
            line_num += 1

    # ================== Приветствие ==================
    def show_welcome_message(self):
        text = """VBA Extractor — извлечение макросов и хэшей паролей

Поддерживаемые форматы: .xls, .xlsm, .xlsb, .xltm

Требования: Python 3.6+, pip install oletools olefile

Как использовать:
  1. Нажмите "Обзор" и выберите файл Excel
  2. Нажмите "Извлечь VBA" — код появится справа, модули — слева
  3. Нажмите "Извлечь хэш пароля" — получите хэш для подбора
  4. Кнопка "Копировать хэш" появится автоматически после извлечения

"""
        self.w["output_text"].delete(1.0, tk.END)
        self.w["output_text"].insert(tk.END, text)
        self.w["output_text"].see("1.0")
        self.highlight_comments()

    # ================== Выбор файла ==================
    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл Excel",
            filetypes=[("Excel files", "*.xls *.xlsm *.xlsx *.xlsb *.xltm"), ("All files", "*.*")])
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
            self.w["output_text"].delete(1.0, tk.END)
            self.w["output_text"].insert(tk.END, code)
            self.highlight_comments()
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
            return
        self.w["output_text"].insert(tk.END, self.extraction_info + "\n" + "-"*50 + "\n")
        for name, code in self.modules:
            self.w["output_text"].insert(tk.END, f"\n' Module: {name}\n")
            self.w["output_text"].insert(tk.END, code + "\n" + "-"*50 + "\n")
        self.highlight_comments()
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
                        return

            if not olefile.isOleFile(ole_data):
                self.root.after(0, self.display_result, "❌ Ошибка: не удалось прочитать OLE-структуру.")
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
                return

            data = ole.openstream(stream).read()
            ole.close()
            text = data.decode('latin-1', errors='ignore')

            # Поиск всех полей защиты
            dpb = re.search(r'(?:DPB|DPx)="([A-Fa-f0-9]+)"', text)
            cmg = re.search(r'CMG="([A-Fa-f0-9]+)"', text)
            gc  = re.search(r'GC="([A-Fa-f0-9]+)"', text)
            wid = re.search(r'ID="([A-Fa-f0-9]+)"', text)
            wep = re.search(r'WEP="([A-Fa-f0-9]+)"', text)

            # 🔹 СТАРЫЙ ФОРМАТ
            if path.lower().endswith('.xls') and dpb:
                h = f"$oldoffice$0*{dpb.group(1)}"
                hashes["old"] = h

                result.append("🔐 СТАРЫЙ ФОРМАТ (Office 97–2010)\n")
                result.append("-"*40 + "\n")
                result.append("Хэш:\n")
                result.append(f"  {h}\n\n")

                result.append("Использование с hashcat:\n")
                result.append("  Режим: -m 9700\n")
                result.append("  Словарь:\n")
                result.append("  hashcat -m 9700 -a 0 hash.txt rockyou.txt\n\n")

                result.append("  Маска (6 символов):\n")
                result.append("  hashcat -m 9700 -a 3 hash.txt ?a?a?a?a?a?a\n\n")

            # 🔹 НОВЫЙ ФОРМАТ
            elif cmg and dpb and gc:
                h = f"$vba$*{cmg.group(1)}*{dpb.group(1)}*{gc.group(1)}"
                hashes["new"] = h

                result.append("🔐 НОВЫЙ ФОРМАТ (Office 2013–2024)\n")
                result.append("-"*40 + "\n")
                result.append("Хэш:\n")
                result.append(f"  {h}\n\n")

                result.append("Обнаруженные поля:\n")
                if cmg: result.append(f"  CMG : {cmg.group(1)}\n")
                if dpb: result.append(f"  DPB : {dpb.group(1)}\n")
                if gc:  result.append(f"  GC  : {gc.group(1)}\n")
                if wid: result.append(f"  ID  : {wid.group(1)}\n")
                if wep: result.append(f"  WEP : {wep.group(1)}\n")

                result.append("\n")

                result.append("Использование с hashcat:\n")
                result.append("  Обычно используется John the Ripper / office2john\n")
                result.append("  либо кастомные VBA-модули.\n\n")

                result.append("Пример:\n")
                result.append("  office2john.py file.xlsm > hash.txt\n")
                result.append("  hashcat hash.txt wordlist.txt\n\n")

            # ⚠ Частичный новый формат
            elif dpb:
                h = f"$partial$*{dpb.group(1)}"
                hashes["new"] = h

                result.append("⚠ ЧАСТИЧНЫЙ НОВЫЙ ФОРМАТ\n")
                result.append("-"*40 + "\n")
                result.append(f"{h}\n\n")

                result.append("Найдены не все поля защиты VBA-проекта.\n")
                result.append("Для полного modern hash обычно нужны:\n")
                result.append("  CMG + DPB + GC\n\n")

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

    def _enable_hash_buttons(self):
        self.w["btn_extract_hash"].config(state=tk.NORMAL, text="Извлечь хэш пароля")
        self.w["btn_extract_vba"].config(state=tk.NORMAL)
        if self.last_hash_value:
            self.w["btn_copy_hash"].pack(side=tk.LEFT, padx=3)
            self.w["btn_copy_hash"].config(bg=self.colors["accent_orange"], fg="white")
            self.w["btn_copy_hash"].config(state=tk.NORMAL)
            self.hash_displayed = True   # показываем, что сейчас вывод хэша
        else:
            self.hash_displayed = False

    def display_result(self, text):
        self.w["output_text"].insert(tk.END, text)
        self.w["output_text"].see("1.0")
        self.highlight_comments()

    # ================== Копирование хэша ==================
    def copy_hash_only(self):
        if not self.last_hash_value:
            messagebox.showwarning("Нет хэша", "Сначала извлеките хэш пароля", parent=self.root)
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.last_hash_value)
        self.root.update()
        messagebox.showinfo("Скопировано", "Хэш скопирован в буфер обмена.", parent=self.root)

    # ================== Копирование модулей (и хэша) ==================
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
        # Если сейчас показана информация о хэше – копируем весь вывод
        if self.hash_displayed:
            content = self.w["output_text"].get("1.0", tk.END).strip()
            if not content:
                messagebox.showwarning("Пусто", "Нет данных для копирования.", parent=self.root)
                return
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            self.root.update()
            messagebox.showinfo("Скопировано", "Информация о хэшах скопирована в буфер обмена.", parent=self.root)
            return

        # Иначе копируем все модули (прежнее поведение)
        if not self.modules:
            messagebox.showwarning("Пусто", "Нет модулей для копирования.", parent=self.root)
            return
        all_code = [f"' Module: {name}\n{code}" for name, code in self.modules]
        text = "\n".join(all_code)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        messagebox.showinfo("Скопировано", "Все модули скопированы.", parent=self.root)

    # ================== Сохранение ==================
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

    def save_all_modules(self):
        if not self.modules:
            messagebox.showwarning("Пусто", "Нет модулей для сохранения.", parent=self.root)
            return
        dir_path = filedialog.askdirectory(title="Папка для сохранения")
        if dir_path:
            try:
                for name, code in self.modules:
                    safe = name.replace(".", "_") + ".bas"
                    with open(os.path.join(dir_path, safe), "w", encoding="utf-8") as f: f.write(code)
                messagebox.showinfo("Сохранено", f"Все модули сохранены в:\n{dir_path}", parent=self.root)
            except Exception as e: messagebox.showerror("Ошибка", f"Ошибка сохранения: {e}", parent=self.root)

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
