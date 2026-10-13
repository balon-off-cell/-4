# -*- coding: utf-8 -*-
"""
Веб-интерфейс Генератора коммерческих предложений и конъюнктурного анализа.
Запуск: streamlit run app.py
"""

from __future__ import annotations

import io
import tempfile
from datetime import date
from pathlib import Path

import streamlit as st
import pandas as pd

from estimate_parser import parse_estimate, EstimateData
from ka_builder import build_ka_data, write_ka_excel, SUPPLIERS
from kp_builder import build_kp_pdf
from blank_extractor import extract_company_assets, CompanyAssets

# ---------------------------------------------------------------------------
# Настройки страницы
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Генератор КП и Конъюнктурного анализа",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📋 Генератор коммерческих предложений и конъюнктурного анализа")
st.caption(
    "Загрузите смету / спецификацию → получите конъюнктурный анализ (3 предложения на позицию) "
    "и готовые КП в PDF."
)

# ---------------------------------------------------------------------------
# Боковая панель — параметры
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Параметры")
    base_markup = st.number_input(
        "Базовый коэффициент наценки",
        min_value=0.5,
        max_value=3.0,
        value=1.0,
        step=0.05,
        help="Применяется ко всем ценам. 1.0 = без дополнительной наценки.",
    )
    object_name = st.text_input("Название объекта", value="")
    work_name = st.text_input("Наименование работ", value="")
    agreed_by = st.text_input("СОГЛАСОВАНО (ФИО / должность)", value="")
    approved_by = st.text_input("УТВЕРЖДАЮ (ФИО / должность)", value="")

    st.divider()
    st.subheader("Поставщики (эталонные)")
    for s in SUPPLIERS:
        st.markdown(f"**{s['short']}** — {s['name']}  \nИНН {s['inn']} · ×{s['markup_factor']}")

# ---------------------------------------------------------------------------
# Основная зона — загрузка файла
# ---------------------------------------------------------------------------
st.subheader("1. Загрузите смету / спецификацию (Excel)")
uploaded_file = st.file_uploader(
    "Поддерживаются ЛСР Гранд-Смета и простые спецификации",
    type=["xlsx", "xls"],
    accept_multiple_files=False,
)

if uploaded_file is None:
    st.info("👆 Загрузите Excel-файл со сметой или спецификацией, чтобы начать.")
    st.stop()

# Сохраняем во временный файл
with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
    tmp.write(uploaded_file.getvalue())
    estimate_path = Path(tmp.name)

try:
    with st.spinner("Читаю смету..."):
        estimate = parse_estimate(estimate_path)
except Exception as e:
    st.error(f"Ошибка при разборе файла: {e}")
    st.stop()

if not estimate.items:
    st.error("В файле не найдено ни одной позиции материалов/оборудования.")
    st.stop()

st.success(f"Найдено позиций: **{len(estimate.items)}**")
if estimate.object_name:
    st.write(f"Объект: {estimate.object_name}")
if estimate.work_name:
    st.write(f"Работы: {estimate.work_name}")

# Превью позиций
with st.expander("Показать позиции из сметы", expanded=False):
    df_preview = pd.DataFrame(
        [
            {
                "№": i + 1,
                "Наименование": it.name[:80],
                "Ед.": it.unit,
                "Кол-во": it.qty,
                "Цена": it.price,
                "Сумма": it.total,
            }
            for i, it in enumerate(estimate.items)
        ]
    )
    st.dataframe(df_preview, use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# 2. Генерация конъюнктурного анализа
# ---------------------------------------------------------------------------
st.subheader("2. Конъюнктурный анализ (3 предложения на позицию)")

if st.button("🚀 Сформировать конъюнктурный анализ", type="primary"):
    with st.spinner("Формирую КА..."):
        positions = build_ka_data(estimate, base_markup=base_markup)

        ka_path = Path(tempfile.gettempdir()) / f"KA_{date.today().strftime('%Y%m%d')}.xlsx"
        write_ka_excel(
            positions,
            ka_path,
            object_name=object_name or estimate.object_name,
            work_name=work_name or estimate.work_name,
            agreed_by=agreed_by,
            approved_by=approved_by,
        )

        st.session_state["ka_path"] = str(ka_path)
        st.session_state["positions"] = positions

    st.success("Конъюнктурный анализ сформирован!")

    # Превью таблицы
    rows = []
    for pos in positions:
        for o in pos.offers:
            rows.append(
                {
                    "Выбор": "➕" if o.is_selected else "",
                    "№": pos.number,
                    "Наименование": pos.name[:60],
                    "Ед.": pos.unit,
                    "Кол-во": pos.qty,
                    "Поставщик": o.supplier["short"],
                    "Цена с НДС": o.unit_price_with_vat,
                    "Сумма с НДС": o.total_with_vat,
                    "Цена без НДС": o.unit_price_without_vat,
                }
            )
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # Скачать
    with open(ka_path, "rb") as f:
        st.download_button(
            label="📥 Скачать КА.xlsx",
            data=f,
            file_name=f"Конъюнктурный_анализ_{date.today().strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---------------------------------------------------------------------------
# 3. Генерация КП (опционально)
# ---------------------------------------------------------------------------
st.subheader("3. Генерация коммерческих предложений (PDF)")

st.markdown(
    "Можно сформировать **3 КП** (по одному на каждого поставщика) "
    "или один КП только по выбранным (самым дешёвым) позициям."
)

col1, col2 = st.columns(2)

with col1:
    gen_three = st.button("Сформировать 3 КП (все поставщики)")

with col2:
    gen_selected = st.button("Сформировать 1 КП (только выбранные +)")

if gen_three or gen_selected:
    if "positions" not in st.session_state:
        st.warning("Сначала сформируйте конъюнктурный анализ.")
    else:
        positions = st.session_state["positions"]
        outputs = []
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            if gen_three:
                for s_idx, supplier in enumerate(SUPPLIERS):
                    items_for_kp = []
                    for pos in positions:
                        offer = pos.offers[s_idx]
                        from estimate_parser import MaterialItem
                        mi = MaterialItem(
                            name=pos.name,
                            unit=pos.unit,
                            qty=pos.qty,
                            price=offer.unit_price_with_vat,
                            code=pos.code,
                        )
                        items_for_kp.append(mi)

                    est = EstimateData(
                        items=items_for_kp,
                        object_name=object_name or estimate.object_name,
                        work_name=work_name or estimate.work_name,
                    )
                    # Реквизиты поставщика
                    assets = CompanyAssets(
                        company_name=supplier["name"],
                        company_details=[
                            f"г. {supplier['city']}",
                            f"ИНН {supplier['inn']}  КПП {supplier.get('kpp', '')}",
                            f"Сайт: {supplier.get('site', '')}",
                        ],
                        director_name="________________",
                    )
                    out_pdf = tmpdir / f"KP_{supplier['short']}.pdf"
                    build_kp_pdf(
                        est,
                        assets,
                        out_pdf,
                        markup=1.0,
                        ref_number=f"№КП-{supplier['short']}-{date.today().strftime('%d%m')}",
                        subject=work_name or estimate.work_name or "поставку материалов",
                        director_name="________________",
                    )
                    outputs.append((f"КП_{supplier['short']}.pdf", out_pdf.read_bytes()))

            if gen_selected:
                items_for_kp = []
                selected_supplier = SUPPLIERS[0]
                for pos in positions:
                    for o in pos.offers:
                        if o.is_selected:
                            from estimate_parser import MaterialItem
                            mi = MaterialItem(
                                name=pos.name,
                                unit=pos.unit,
                                qty=pos.qty,
                                price=o.unit_price_with_vat,
                                code=pos.code,
                            )
                            items_for_kp.append(mi)
                            selected_supplier = o.supplier
                            break

                est = EstimateData(
                    items=items_for_kp,
                    object_name=object_name or estimate.object_name,
                    work_name=work_name or estimate.work_name,
                )
                assets = CompanyAssets(
                    company_name=selected_supplier["name"],
                    company_details=[
                        f"г. {selected_supplier['city']}",
                        f"ИНН {selected_supplier['inn']}",
                    ],
                    director_name="________________",
                )
                out_pdf = tmpdir / "KP_selected.pdf"
                build_kp_pdf(
                    est,
                    assets,
                    out_pdf,
                    markup=1.0,
                    ref_number=f"№КП-выбранные-{date.today().strftime('%d%m')}",
                    subject=work_name or estimate.work_name or "поставку материалов",
                    director_name="________________",
                )
                outputs.append(("KP_выбранные.pdf", out_pdf.read_bytes()))

        st.success(f"Сформировано файлов: {len(outputs)}")
        for name, data in outputs:
            st.download_button(
                label=f"📥 Скачать {name}",
                data=data,
                file_name=name,
                mime="application/pdf",
                key=name,
            )

# ---------------------------------------------------------------------------
# Подвал
# ---------------------------------------------------------------------------
st.divider()
st.caption(
    "Генератор КП и КА · версии для браузера (Streamlit) · "
    "Цены формируются автоматически (ЛУИС+ — базовая, ЭТМ +3.5 %, Машпромлизинг +7 %)."
)
