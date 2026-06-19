import logging
import os
import queue
import threading
import tkinter as tk
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

from generate_xml import process_files


class QueueHandler(logging.Handler):
    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))


class MemoryHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        self.records = []
        self.has_errors = False

    def emit(self, record):
        self.records.append(record)
        if record.levelno >= logging.ERROR:
            self.has_errors = True


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("xlsx2xml — Конвертация УПД")
        self.root.geometry("700x750")
        self.root.resizable(True, True)

        self.xlsx_files = []
        self.output_dir = ""
        self.log_queue = queue.Queue()
        self.running = False

        self._build_ui()
        self._poll_queue()

    def _build_ui(self):
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Files section ---
        file_frame = tk.LabelFrame(main_frame, text="Файлы XLSX", padx=5, pady=5)
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        self.files_listbox = tk.Listbox(file_frame, selectmode=tk.EXTENDED)
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(
            file_frame, orient=tk.VERTICAL, command=self.files_listbox.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.files_listbox.config(yscrollcommand=scrollbar.set)

        btn_file_frame = tk.Frame(file_frame)
        btn_file_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))

        tk.Button(
            btn_file_frame, text="Добавить файлы", command=self._add_files, width=18
        ).pack(pady=2)
        tk.Button(
            btn_file_frame,
            text="Удалить выбранные",
            command=self._remove_files,
            width=18,
        ).pack(pady=2)
        tk.Button(
            btn_file_frame, text="Очистить список", command=self._clear_files, width=18
        ).pack(pady=2)

        # --- Output directory section ---
        dir_frame = tk.LabelFrame(main_frame, text="Директория для XML", padx=5, pady=5)
        dir_frame.pack(fill=tk.X, pady=(0, 5))

        self.dir_entry = tk.Entry(dir_frame)
        self.dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        tk.Button(dir_frame, text="Обзор...", command=self._browse_dir, width=12).pack(
            side=tk.RIGHT
        )

        # --- Convert button ---
        self.convert_btn = tk.Button(
            main_frame,
            text="Запустить конвертацию",
            command=self._start_conversion,
            bg="#4CAF50",
            fg="white",
            font=("", 11, "bold"),
            height=2,
        )
        self.convert_btn.pack(fill=tk.X, pady=(0, 5))

        # --- Log section ---
        log_frame = tk.LabelFrame(main_frame, text="Лог", padx=5, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # --- About section ---
        about_frame = tk.LabelFrame(main_frame, text="О проекте", padx=5, pady=5)
        about_frame.pack(fill=tk.X)

        repo_link = tk.Label(
            about_frame,
            text="https://github.com/Idvon/xlsx2xml.git",
            fg="#1a73e8",
            cursor="hand2",
            anchor=tk.W,
        )
        repo_link.pack(side=tk.LEFT)
        repo_link.bind(
            "<Button-1>",
            lambda e: webbrowser.open("https://github.com/Idvon/xlsx2xml.git"),
        )

        tk.Label(about_frame, text="  |  Автор: Idvon  |  MIT", anchor=tk.W).pack(
            side=tk.LEFT
        )

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Выберите файлы XLSX",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        for f in files:
            if f not in self.xlsx_files:
                self.xlsx_files.append(f)
                self.files_listbox.insert(tk.END, os.path.basename(f))

    def _remove_files(self):
        selected = self.files_listbox.curselection()
        for i in reversed(selected):
            del self.xlsx_files[i]
            self.files_listbox.delete(i)

    def _clear_files(self):
        self.xlsx_files.clear()
        self.files_listbox.delete(0, tk.END)

    def _browse_dir(self):
        directory = filedialog.askdirectory(
            title="Выберите директорию для сохранения XML"
        )
        if directory:
            self.output_dir = directory
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)

    def _log(self, message: str):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _poll_queue(self):
        while not self.log_queue.empty():
            try:
                msg = self.log_queue.get_nowait()
                self._log(msg)
            except queue.Empty:
                break
        self.root.after(100, self._poll_queue)

    def _start_conversion(self):
        if self.running:
            return

        if not self.xlsx_files:
            messagebox.showwarning("Нет файлов", "Добавьте хотя бы один XLSX файл.")
            return

        output_dir = self.dir_entry.get().strip()
        if not output_dir:
            messagebox.showwarning(
                "Нет директории", "Укажите директорию для сохранения XML."
            )
            return

        out_path = Path(output_dir)
        try:
            out_path.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать директорию:\n{e}")
            return

        memory_handler = MemoryHandler()
        memory_handler.setFormatter(
            logging.Formatter(
                "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
            )
        )

        logger = logging.getLogger("xlsx2xml_gui")
        logger.setLevel(logging.INFO)
        logger.addHandler(memory_handler)
        queue_handler = QueueHandler(self.log_queue)
        queue_handler.setFormatter(
            logging.Formatter(
                "%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"
            )
        )
        logger.addHandler(queue_handler)

        self._log(f"Запуск конвертации {len(self.xlsx_files)} файлов...")
        self._log(f"Директория вывода: {out_path}")
        self.running = True
        self.convert_btn.config(state=tk.DISABLED, text="Конвертация...")

        def worker():
            try:
                results = process_files(self.xlsx_files, out_path, logger=logger)
                success = sum(1 for _, _, err in results if err is None)
                failed = len(results) - success
                self.log_queue.put(f"\nГотово. Успешно: {success}, Ошибок: {failed}")
            except Exception as e:
                self.log_queue.put(f"Критическая ошибка: {e}")
                memory_handler.has_errors = True
            finally:
                logger.removeHandler(memory_handler)
                logger.removeHandler(queue_handler)
                self.root.after(
                    0, lambda: self._conversion_done(memory_handler, out_path)
                )

        threading.Thread(target=worker, daemon=True).start()

    def _conversion_done(self, memory_handler: MemoryHandler, out_path: Path):
        if memory_handler.has_errors:
            log_file = (
                out_path / f"conversion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            )
            with open(log_file, "w", encoding="utf-8") as f:
                for record in memory_handler.records:
                    f.write(memory_handler.format(record) + "\n")
            self._log(f"Файл ошибок: {log_file}")
        self.running = False
        self.convert_btn.config(state=tk.NORMAL, text="Запустить конвертацию")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
