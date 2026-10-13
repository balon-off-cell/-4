# -*- coding: utf-8 -*-
"""
Генератор коммерческих предложений (КП).

Использование (CLI):
  python main.py --estimate смета.xlsx --blank бланк.pdf --output КП.pdf --markup 1.3

Или запустить без аргументов — откроется графический интерфейс.
"""

from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

# Чтобы импорты работали и при запуске из папки, и как модуля
sys.path.insert(0, str(Path(__file__).resolve().parent))

from estimate_parser import parse_estimate
from blank_extractor import extract_company_assets
from kp_builder import build_kp_pdf


def run_cli(args: argparse.Namespace) -> int:
    estimate_path = Path(args.estimate)
    blank_path = Path(args.blank)
    output_path = Path(args.output)

    if not estimate_path.exists():
        print(f"Ошибка: файл сметы/спецификации не найден: {estimate_path}")
        return 1
    if not blank_path.exists():
        print(f"Ошибка: бланк фирмы не найден: {blank_path}")
        return 1

    print(f"Читаю смету: {estimate_path.name}")
    estimate = parse_estimate(estimate_path)
    print(f"  Найдено позиций материалов/оборудования: {len(estimate.items)}")
    if not estimate.items:
        print("Ошибка: в файле не найдено ни одной позиции материалов.")
        return 1

    print(f"Извлекаю логотип и печать из бланка: {blank_path.name}")
    with tempfile.TemporaryDirectory(prefix="kp_") as tmp:
        assets = extract_company_assets(blank_path, tmp)
        print(f"  Логотип: {'да' if assets.logo_path else 'нет'}")
        print(f"  Печать:  {'да' if assets.stamp_path else 'нет'}")
        if assets.company_name:
            print(f"  Компания: {assets.company_name}")

        print(f"Формирую КП (наценка ×{args.markup}) → {output_path}")
        build_kp_pdf(
            estimate,
            assets,
            output_path,
            markup=args.markup,
            ref_number=args.ref or "",
            subject=args.subject or "",
            delivery_days=args.delivery_days,
            advance_percent=args.advance,
            validity_days=args.validity,
            director_name=args.director or "",
        )

    print("Готово.")
    return 0


def run_gui() -> int:
    """Простой GUI на tkinter."""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except ImportError:
        print("tkinter недоступен. Используйте CLI:")
        print("  python main.py --estimate смета.xlsx --blank бланк.pdf --output КП.pdf")
        return 1

    root = tk.Tk()
    root.title("Генератор коммерческих предложений (КП)")
    root.geometry("640x480")
    root.minsize(560, 420)

    # --- переменные ---
    estimate_var = tk.StringVar()
    blank_var = tk.StringVar()
    output_var = tk.StringVar()
    markup_var = tk.StringVar(value="1.3")
    ref_var = tk.StringVar()
    subject_var = tk.StringVar()
    director_var = tk.StringVar()
    status_var = tk.StringVar(value="Выберите файлы сметы и бланка фирмы")

    def browse_estimate():
        p = filedialog.askopenfilename(
            title="Смета или спецификация",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Все файлы", "*.*")],
        )
        if p:
            estimate_var.set(p)

    def browse_blank():
        p = filedialog.askopenfilename(
            title="Бланк фирмы (PDF с печатью и подписью)",
            filetypes=[("PDF", "*.pdf"), ("Все файлы", "*.*")],
        )
        if p:
            blank_var.set(p)

    def browse_output():
        p = filedialog.asksaveasfilename(
            title="Сохранить КП как",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="КП.pdf",
        )
        if p:
            output_var.set(p)

    def generate():
        est = estimate_var.get().strip()
        blank = blank_var.get().strip()
        out = output_var.get().strip()
        if not est or not blank:
            messagebox.showerror("Ошибка", "Укажите файл сметы и бланк фирмы.")
            return
        if not out:
            out = str(Path(est).with_name("КП_готовое.pdf"))
            output_var.set(out)
        try:
            coef = float(markup_var.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Ошибка", "Коэффициент наценки должен быть числом (например 1.3)")
            return

        status_var.set("Обработка...")
        root.update_idletasks()

        try:
            estimate = parse_estimate(est)
            if not estimate.items:
                messagebox.showerror("Ошибка", "В смете/спецификации не найдено материалов.")
                status_var.set("Ошибка")
                return

            with tempfile.TemporaryDirectory(prefix="kp_") as tmp:
                assets = extract_company_assets(blank, tmp)
                build_kp_pdf(
                    estimate,
                    assets,
                    out,
                    markup=coef,
                    ref_number=ref_var.get().strip(),
                    subject=subject_var.get().strip(),
                    director_name=director_var.get().strip(),
                )

            status_var.set(f"Готово: {out}  ({len(estimate.items)} поз.)")
            messagebox.showinfo(
                "Успех",
                f"Коммерческое предложение сохранено:\n{out}\n\n"
                f"Позиций: {len(estimate.items)}\n"
                f"Логотип: {'да' if assets.logo_path else 'нет'}\n"
                f"Печать: {'да' if assets.stamp_path else 'нет'}",
            )
        except Exception as e:
            status_var.set("Ошибка")
            messagebox.showerror("Ошибка", str(e))

    # --- layout ---
    pad = {"padx": 8, "pady": 4}
    frm = ttk.Frame(root, padding=12)
    frm.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frm, text="Смета / спецификация (Excel):").grid(row=0, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=estimate_var, width=50).grid(row=0, column=1, sticky="ew", **pad)
    ttk.Button(frm, text="Обзор…", command=browse_estimate).grid(row=0, column=2, **pad)

    ttk.Label(frm, text="Бланк фирмы (PDF):").grid(row=1, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=blank_var, width=50).grid(row=1, column=1, sticky="ew", **pad)
    ttk.Button(frm, text="Обзор…", command=browse_blank).grid(row=1, column=2, **pad)

    ttk.Label(frm, text="Сохранить КП как:").grid(row=2, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=output_var, width=50).grid(row=2, column=1, sticky="ew", **pad)
    ttk.Button(frm, text="Обзор…", command=browse_output).grid(row=2, column=2, **pad)

    ttk.Separator(frm, orient="horizontal").grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)

    ttk.Label(frm, text="Коэффициент наценки:").grid(row=4, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=markup_var, width=10).grid(row=4, column=1, sticky="w", **pad)

    ttk.Label(frm, text="Исх. номер (необязательно):").grid(row=5, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=ref_var, width=30).grid(row=5, column=1, sticky="w", **pad)

    ttk.Label(frm, text="Тема / объект (необязательно):").grid(row=6, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=subject_var, width=50).grid(row=6, column=1, sticky="ew", **pad)

    ttk.Label(frm, text="ФИО директора (необязательно):").grid(row=7, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=director_var, width=30).grid(row=7, column=1, sticky="w", **pad)

    ttk.Separator(frm, orient="horizontal").grid(row=8, column=0, columnspan=3, sticky="ew", pady=10)

    btn = ttk.Button(frm, text="Сформировать КП", command=generate)
    btn.grid(row=9, column=0, columnspan=3, pady=8)

    ttk.Label(frm, textvariable=status_var, foreground="#333").grid(
        row=10, column=0, columnspan=3, sticky="w", **pad
    )

    frm.columnconfigure(1, weight=1)

    # Подсказка
    hint = (
        "Как пользоваться:\n"
        "1. Укажите Excel-смету (ЛСР Гранд-Смета или простую спецификацию с колонками "
        "Наименование / Кол-во / Цена).\n"
        "2. Укажите PDF-бланк компании, в котором уже есть логотип, печать и подпись — "
        "программа вырежет их автоматически.\n"
        "3. Задайте наценку (например 1.3) и нажмите «Сформировать КП»."
    )
    ttk.Label(frm, text=hint, wraplength=580, foreground="#555", justify="left").grid(
        row=11, column=0, columnspan=3, sticky="w", pady=12
    )

    root.mainloop()
    return 0


def main():
    parser = argparse.ArgumentParser(
        description="Генератор коммерческих предложений из сметы и бланка фирмы"
    )
    parser.add_argument("--estimate", "-e", help="Путь к Excel-смете или спецификации")
    parser.add_argument("--blank", "-b", help="Путь к PDF-бланку фирмы (с печатью)")
    parser.add_argument("--output", "-o", default="КП.pdf", help="Куда сохранить готовый КП")
    parser.add_argument("--markup", "-m", type=float, default=1.3, help="Коэффициент наценки (по умолчанию 1.3)")
    parser.add_argument("--ref", help="Исходящий номер КП")
    parser.add_argument("--subject", help="Тема / наименование объекта")
    parser.add_argument("--director", help="ФИО генерального директора")
    parser.add_argument("--delivery-days", type=int, default=30)
    parser.add_argument("--advance", type=int, default=80)
    parser.add_argument("--validity", type=int, default=30)
    parser.add_argument("--gui", action="store_true", help="Запустить графический интерфейс")

    args = parser.parse_args()

    if args.gui or (not args.estimate and not args.blank):
        return run_gui()
    return run_cli(args)


if __name__ == "__main__":
    sys.exit(main())
