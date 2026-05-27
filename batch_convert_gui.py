import os
import csv
import traceback
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from dbfread import DBF

try:
    import pandas as pd
except Exception:
    pd = None


def dbf_to_csv(src_path: Path, dst_path: Path, delimiter=";", encoding="utf-8-sig"):
    """Convert one DBF file to CSV."""
    table = DBF(str(src_path), load=True, char_decode_errors="replace")
    fieldnames = list(table.field_names)

    dst_path.parent.mkdir(parents=True, exist_ok=True)
    with open(dst_path, "w", newline="", encoding=encoding) as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=delimiter)
        writer.writeheader()
        for record in table:
            writer.writerow(record)


def dbf_to_xlsx(src_path: Path, dst_path: Path):
    """Convert one DBF file to XLSX using pandas."""
    if pd is None:
        raise RuntimeError("pandas не установлен. Установите зависимости: pip install -r requirements.txt")

    table = DBF(str(src_path), load=True, char_decode_errors="replace")
    df = pd.DataFrame(list(table))
    dst_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(dst_path, index=False)


def csv_to_xlsx(src_path: Path, dst_path: Path, delimiter=";", encoding="utf-8-sig"):
    """Convert one CSV file to XLSX using pandas."""
    if pd is None:
        raise RuntimeError("pandas не установлен. Установите зависимости: pip install -r requirements.txt")

    df = pd.read_csv(src_path, sep=delimiter, encoding=encoding, dtype=str, keep_default_na=False)
    dst_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(dst_path, index=False)


def convert_file(src: Path, out_dir: Path, out_format: str, delimiter: str, encoding: str):
    out_dir.mkdir(parents=True, exist_ok=True)

    suffix = src.suffix.lower()

    # DBF -> CSV/XLSX
    if suffix == ".dbf" and out_format == "csv":
        dst = out_dir / (src.stem + ".csv")
        dbf_to_csv(src, dst, delimiter=delimiter, encoding=encoding)
        return dst

    if suffix == ".dbf" and out_format == "xlsx":
        dst = out_dir / (src.stem + ".xlsx")
        dbf_to_xlsx(src, dst)
        return dst

    # CSV -> XLSX
    if suffix == ".csv" and out_format == "xlsx":
        dst = out_dir / (src.stem + ".xlsx")
        csv_to_xlsx(src, dst, delimiter=delimiter, encoding=encoding)
        return dst

    raise ValueError(f"Нет конвертера для: {src.suffix} -> {out_format}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Batch Converter (DBF/CSV → CSV/XLSX)")
        self.geometry("860x560")

        self.in_dir = tk.StringVar()
        self.out_dir = tk.StringVar()
        self.out_format = tk.StringVar(value="csv")
        self.delimiter = tk.StringVar(value=";")
        self.encoding = tk.StringVar(value="utf-8-sig")
        self.recursive = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        # Input folder
        row = ttk.Frame(frm)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text="Папка с файлами:").pack(side="left")
        ttk.Entry(row, textvariable=self.in_dir).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Выбрать…", command=self.pick_in).pack(side="left")

        # Output folder
        row = ttk.Frame(frm)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text="Папка для результата:").pack(side="left")
        ttk.Entry(row, textvariable=self.out_dir).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Выбрать…", command=self.pick_out).pack(side="left")

        # Options
        opt = ttk.LabelFrame(frm, text="Параметры", padding=10)
        opt.pack(fill="x", pady=10)

        opt_row = ttk.Frame(opt)
        opt_row.pack(fill="x", pady=4)
        ttk.Label(opt_row, text="Выходной формат:").pack(side="left")
        ttk.Combobox(
            opt_row,
            textvariable=self.out_format,
            values=["csv", "xlsx"],
            width=10,
            state="readonly",
        ).pack(side="left", padx=8)

        opt_row = ttk.Frame(opt)
        opt_row.pack(fill="x", pady=4)
        ttk.Label(opt_row, text="Разделитель CSV:").pack(side="left")
        ttk.Combobox(opt_row, textvariable=self.delimiter, values=[",", ";", "\t"], width=10, state="readonly").pack(
            side="left", padx=8
        )

        opt_row = ttk.Frame(opt)
        opt_row.pack(fill="x", pady=4)
        ttk.Label(opt_row, text="Кодировка CSV:").pack(side="left")
        ttk.Combobox(opt_row, textvariable=self.encoding, values=["utf-8-sig", "utf-8", "cp1251"], width=12, state="readonly").pack(
            side="left", padx=8
        )

        ttk.Checkbutton(opt, text="Искать файлы рекурсивно (в подпапках)", variable=self.recursive).pack(anchor="w", pady=4)

        # Start button + progress
        run_row = ttk.Frame(frm)
        run_row.pack(fill="x", pady=8)
        ttk.Button(run_row, text="Конвертировать", command=self.run).pack(side="left")
        self.progress = ttk.Progressbar(run_row, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)
        self.status = ttk.Label(run_row, text="")
        self.status.pack(side="left")

        # Log
        log_box = ttk.LabelFrame(frm, text="Лог", padding=8)
        log_box.pack(fill="both", expand=True)
        self.log = tk.Text(log_box, height=16, wrap="word")
        self.log.pack(fill="both", expand=True)

        self._log("Готово. Выберите папки и нажмите «Конвертировать».\n")

    def _log(self, s: str):
        self.log.insert("end", s)
        self.log.see("end")
        self.update_idletasks()

    def pick_in(self):
        p = filedialog.askdirectory()
        if p:
            self.in_dir.set(p)

    def pick_out(self):
        p = filedialog.askdirectory()
        if p:
            self.out_dir.set(p)

    def run(self):
        in_dir = Path(self.in_dir.get().strip())
        out_dir = Path(self.out_dir.get().strip())

        if not in_dir.exists():
            messagebox.showerror("Ошибка", "Выберите существующую папку с файлами.")
            return
        if not out_dir:
            messagebox.showerror("Ошибка", "Выберите папку для результата.")
            return

        # Determine patterns based on output format
        if self.out_format.get() == "xlsx":
            # For XLSX output we support DBF->XLSX and CSV->XLSX
            patterns = ["**/*.dbf", "**/*.csv"] if self.recursive.get() else ["*.dbf", "*.csv"]
            files = []
            for pat in patterns:
                files.extend(list(in_dir.glob(pat)))
            files = sorted(set(files))
        else:
            pattern = "**/*.dbf" if self.recursive.get() else "*.dbf"
            files = sorted(in_dir.glob(pattern))

        if not files:
            messagebox.showinfo("Нет файлов", "В выбранной папке не найдено подходящих файлов.")
            return

        self.progress["value"] = 0
        self.progress["maximum"] = len(files)

        self._log(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Найдено файлов: {len(files)}\n")
        ok = 0
        fail = 0

        for i, src in enumerate(files, start=1):
            try:
                rel_parent = src.parent.relative_to(in_dir)
                dst_folder = out_dir / rel_parent
                dst = convert_file(
                    src=src,
                    out_dir=dst_folder,
                    out_format=self.out_format.get(),
                    delimiter=self.delimiter.get(),
                    encoding=self.encoding.get(),
                )
                ok += 1
                self._log(f"OK  {src}  ->  {dst}\n")
            except Exception as e:
                fail += 1
                self._log(f"ERR {src}: {e}\n")
                self._log(traceback.format_exc() + "\n")

            self.progress["value"] = i
            self.status.config(text=f"{i}/{len(files)}")

        self._log(f"\nГотово: OK={ok}, ERR={fail}\n")
        messagebox.showinfo("Готово", f"Конвертация завершена.\nOK={ok}\nERR={fail}")


if __name__ == "__main__":
    App().mainloop()
