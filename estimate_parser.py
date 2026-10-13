# -*- coding: utf-8 -*-
"""
Парсер локальных смет (ЛСР) и спецификаций из Excel.
Поддерживает формат Гранд-Смета / Методика 2020 (РИМ) и упрощённые спецификации.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple

import openpyxl


@dataclass
class MaterialItem:
    """Одна позиция материала / оборудования."""
    name: str
    unit: str
    qty: float
    price: float  # цена за единицу в текущем уровне (из сметы)
    code: str = ""
    source_row: int = 0

    @property
    def total(self) -> float:
        return round(self.qty * self.price, 2)

    def with_markup(self, coef: float) -> Tuple[float, float]:
        """Возвращает (цена_с_наценкой, сумма_с_наценкой)."""
        unit_price = round(self.price * coef, 2)
        amount = round(self.qty * unit_price, 2)
        return unit_price, amount


@dataclass
class EstimateData:
    """Данные, извлечённые из сметы/спецификации."""
    items: List[MaterialItem] = field(default_factory=list)
    object_name: str = ""
    work_name: str = ""
    source_file: str = ""

    @property
    def materials_total(self) -> float:
        return sum(i.total for i in self.items)


# Ключевые слова, по которым отсекаем не-материалы
_SKIP_NAME = (
    "от(", "отм(", "нр ", "сп ", "итого", "всего", "затраты труда",
    "средний разряд", "вспомогательные ненормируемые", "прямые затраты",
    "оплата труда", "монтажные работы", "справочно",
)
_SKIP_UNITS = {"маш.-ч", "чел.-ч", "%", "маш-ч", "чел-ч"}


def _is_material_row(code, name, unit, qty, price) -> bool:
    if name is None or unit is None or qty is None or price is None:
        return False
    name_str = str(name).strip()
    code_str = str(code).strip() if code else ""
    unit_str = str(unit).strip().lower()

    if unit_str in _SKIP_UNITS:
        return False
    if not code_str or len(code_str) < 3:
        return False
    if code_str in {"1", "2", "3", "4"}:
        return False
    name_lower = name_str.lower()
    if any(k in name_lower for k in _SKIP_NAME):
        return False
    try:
        q = float(qty)
        p = float(price)
    except (TypeError, ValueError):
        return False
    if q <= 0 or p < 0:
        return False
    return True


def parse_grand_smeta_lsr(path: str | Path, sheet_name: Optional[str] = None) -> EstimateData:
    """
    Разбор локальной сметы в формате Гранд-Смета (Методика 2020).
    Ожидаемые колонки (типичные):
      B (2) — код, C (3) — наименование, H (8) — ед.изм.,
      K (11) — количество, N (14) — цена текущая.
    """
    path = Path(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    result = EstimateData(source_file=str(path))

    # Пытаемся найти название объекта / работ в шапке
    for row in range(1, min(40, ws.max_row + 1)):
        for col in range(1, 20):
            val = ws.cell(row=row, column=col).value
            if not val or not isinstance(val, str):
                continue
            v = val.strip()
            if "наименование объекта" in v.lower() or "наименование стройки" in v.lower():
                # следующее непустое значение в строке или ниже
                for c2 in range(col + 1, 25):
                    v2 = ws.cell(row=row, column=c2).value
                    if v2 and str(v2).strip():
                        result.object_name = str(v2).strip()
                        break
            if "наименование работ" in v.lower() or "система для передачи" in v.lower():
                result.work_name = v if "система" in v.lower() else result.work_name
            # прямое название работ
            if v.startswith("Система") or "передача данных" in v.lower():
                result.work_name = v

    # Основной проход по позициям
    for row in range(1, ws.max_row + 1):
        code = ws.cell(row=row, column=2).value
        name = ws.cell(row=row, column=3).value
        unit = ws.cell(row=row, column=8).value
        qty = ws.cell(row=row, column=11).value
        price = ws.cell(row=row, column=14).value

        if not _is_material_row(code, name, unit, qty, price):
            continue

        result.items.append(MaterialItem(
            name=str(name).strip(),
            unit=str(unit).strip(),
            qty=float(qty),
            price=float(price),
            code=str(code).strip(),
            source_row=row,
        ))

    return result


def parse_simple_spec(path: str | Path, sheet_name: Optional[str] = None) -> EstimateData:
    """
    Разбор упрощённой спецификации.
    Ищет строку-заголовок с колонками вида: Наименование | Ед. | Кол-во | Цена
    (регистр не важен, допускаются варианты «Количество», «Цена за ед.» и т.п.).
    """
    path = Path(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    result = EstimateData(source_file=str(path))

    # Поиск заголовка
    header_row = None
    col_map = {}  # logical -> column index (1-based)

    name_keys = ("наименование", "название", "товар", "материал", "позиция")
    unit_keys = ("ед", "единица", "изм")
    qty_keys = ("кол", "объем", "количество", "кол-во", "qty")
    price_keys = ("цена", "стоимость", "price")

    for row in range(1, min(50, ws.max_row + 1)):
        cells = {}
        for col in range(1, min(30, ws.max_column + 1)):
            val = ws.cell(row=row, column=col).value
            if val and isinstance(val, str):
                cells[col] = val.strip().lower()
        if not cells:
            continue

        found = {}
        for col, text in cells.items():
            if any(k in text for k in name_keys) and "name" not in found:
                found["name"] = col
            elif any(k in text for k in unit_keys) and "unit" not in found:
                found["unit"] = col
            elif any(k in text for k in qty_keys) and "qty" not in found:
                found["qty"] = col
            elif any(k in text for k in price_keys) and "price" not in found:
                found["price"] = col

        if "name" in found and "qty" in found and "price" in found:
            header_row = row
            col_map = found
            if "unit" not in col_map:
                col_map["unit"] = None
            break

    if header_row is None:
        # fallback: пробуем как Гранд-Смету
        return parse_grand_smeta_lsr(path, sheet_name)

    for row in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(row=row, column=col_map["name"]).value
        qty = ws.cell(row=row, column=col_map["qty"]).value
        price = ws.cell(row=row, column=col_map["price"]).value
        unit = ws.cell(row=row, column=col_map["unit"]).value if col_map.get("unit") else "шт"

        if name is None or qty is None or price is None:
            continue
        try:
            q = float(qty)
            p = float(price)
        except (TypeError, ValueError):
            continue
        if q <= 0:
            continue
        name_str = str(name).strip()
        if not name_str or name_str.lower() in ("итого", "всего"):
            continue

        result.items.append(MaterialItem(
            name=name_str,
            unit=str(unit).strip() if unit else "шт",
            qty=q,
            price=p,
            source_row=row,
        ))

    return result


def parse_estimate(path: str | Path, sheet_name: Optional[str] = None) -> EstimateData:
    """
    Универсальный вход: сначала пробуем как упрощённую спецификацию,
    при неудаче — как ЛСР Гранд-Сметы.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")

    # Быстрая эвристика по имени
    name_lower = path.name.lower()
    if "лср" in name_lower or "смета" in name_lower or "рим" in name_lower:
        data = parse_grand_smeta_lsr(path, sheet_name)
        if data.items:
            return data

    data = parse_simple_spec(path, sheet_name)
    if data.items:
        return data

    # Последняя попытка — Гранд-Смета
    return parse_grand_smeta_lsr(path, sheet_name)
