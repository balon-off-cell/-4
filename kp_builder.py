# -*- coding: utf-8 -*-
"""
Сборка коммерческого предложения (КП) в PDF.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import List, Optional

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from estimate_parser import EstimateData, MaterialItem
from blank_extractor import CompanyAssets

# Шрифты с кириллицей
_FONT_REG = "Helvetica"
_FONT_BOLD = "Helvetica-Bold"
_FONTS_REGISTERED = False


def _register_fonts():
    """Регистрирует TTF-шрифты с кириллицей. Скачивает DejaVu, если нет в проекте."""
    global _FONT_REG, _FONT_BOLD, _FONTS_REGISTERED
    if _FONTS_REGISTERED:
        return

    base = Path(__file__).resolve().parent
    fonts_dir = base / "fonts"
    reg_path = fonts_dir / "DejaVuSans.ttf"
    bold_path = fonts_dir / "DejaVuSans-Bold.ttf"

    # Если шрифтов нет — пробуем скачать (Streamlit Cloud / первый запуск)
    if not reg_path.exists() or not bold_path.exists():
        try:
            import urllib.request
            import zipfile
            import io
            fonts_dir.mkdir(parents=True, exist_ok=True)
            url = (
                "https://github.com/dejavu-fonts/dejavu-fonts/releases/"
                "download/version_2_37/dejavu-fonts-ttf-2.37.zip"
            )
            data = urllib.request.urlopen(url, timeout=30).read()
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for name in z.namelist():
                    if name.endswith("DejaVuSans.ttf"):
                        reg_path.write_bytes(z.read(name))
                    elif name.endswith("DejaVuSans-Bold.ttf"):
                        bold_path.write_bytes(z.read(name))
        except Exception:
            pass

    candidates = [
        (reg_path, bold_path),
        (base / "DejaVuSans.ttf", base / "DejaVuSans-Bold.ttf"),
        (Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
         Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")),
        (Path("/usr/share/fonts/TTF/DejaVuSans.ttf"),
         Path("/usr/share/fonts/TTF/DejaVuSans-Bold.ttf")),
        (Path("C:/Windows/Fonts/arial.ttf"), Path("C:/Windows/Fonts/arialbd.ttf")),
    ]

    for r, b in candidates:
        if r.exists() and b.exists():
            try:
                pdfmetrics.registerFont(TTFont("DejaVu", str(r)))
                pdfmetrics.registerFont(TTFont("DejaVuBold", str(b)))
                _FONT_REG = "DejaVu"
                _FONT_BOLD = "DejaVuBold"
                _FONTS_REGISTERED = True
                return
            except Exception:
                continue

    # Fallback без кириллицы
    _FONT_REG = "Helvetica"
    _FONT_BOLD = "Helvetica-Bold"
    _FONTS_REGISTERED = True


def _fmt_num(n: float, d: int = 2) -> str:
    return f"{n:,.{d}f}".replace(",", " ")


def _fmt_qty(q: float) -> str:
    if abs(q - round(q)) < 1e-9:
        return str(int(round(q)))
    return f"{q:.4f}".rstrip("0").rstrip(".")


def build_kp_pdf(
    estimate: EstimateData,
    assets: CompanyAssets,
    output_path: str | Path,
    *,
    markup: float = 1.3,
    ref_number: str = "",
    subject: str = "",
    delivery_days: int = 30,
    advance_percent: int = 80,
    validity_days: int = 30,
    director_name: str = "",
    include_vat_note: bool = True,
) -> Path:
    """
    Формирует PDF коммерческого предложения.

    :param estimate: данные из сметы/спецификации
    :param assets: логотип, печать, реквизиты из бланка
    :param output_path: куда сохранить
    :param markup: коэффициент наценки (1.3 = +30%)
    :param ref_number: исходящий номер (если пусто — генерируется)
    :param subject: тема / описание объекта
    :param director_name: ФИО директора (если не извлечено из бланка)
    """
    _register_fonts()
    output_path = Path(output_path)

    if not ref_number:
        ref_number = f"№КП-{date.today().strftime('%d%m%y')}"

    if not subject:
        parts = []
        if estimate.work_name:
            parts.append(estimate.work_name)
        if estimate.object_name:
            parts.append(estimate.object_name)
        subject = " / ".join(parts) if parts else "поставку материалов и оборудования"

    dir_name = director_name or assets.director_name or "________________"

    # --- стили ---
    style_company_name = ParagraphStyle(
        "CN", fontName=_FONT_BOLD, fontSize=11, alignment=TA_CENTER, spaceAfter=0.5 * mm
    )
    style_company = ParagraphStyle(
        "C", fontName=_FONT_REG, fontSize=8, alignment=TA_CENTER, leading=10
    )
    style_ref = ParagraphStyle(
        "R", fontName=_FONT_REG, fontSize=9, alignment=TA_LEFT, spaceBefore=2 * mm, spaceAfter=1 * mm
    )
    style_title = ParagraphStyle(
        "T", fontName=_FONT_BOLD, fontSize=12, alignment=TA_CENTER, spaceBefore=2 * mm, spaceAfter=3 * mm
    )
    style_body = ParagraphStyle(
        "B", fontName=_FONT_REG, fontSize=9, alignment=TA_JUSTIFY, leading=11.5, spaceAfter=2 * mm
    )
    style_body_bold = ParagraphStyle(
        "BB", fontName=_FONT_BOLD, fontSize=9, alignment=TA_LEFT, leading=11.5, spaceAfter=1.5 * mm
    )
    style_th = ParagraphStyle("TH", fontName=_FONT_BOLD, fontSize=6.5, alignment=TA_CENTER, leading=8)
    style_tc = ParagraphStyle("TC", fontName=_FONT_REG, fontSize=6, leading=7.5, alignment=TA_LEFT)
    style_tn = ParagraphStyle("TN", fontName=_FONT_REG, fontSize=6, leading=7.5, alignment=TA_RIGHT)
    style_tcen = ParagraphStyle("TCen", fontName=_FONT_REG, fontSize=6, leading=7.5, alignment=TA_CENTER)
    style_total = ParagraphStyle(
        "Tot", fontName=_FONT_REG, fontSize=9, alignment=TA_JUSTIFY, leading=11.5, spaceBefore=2.5 * mm
    )
    style_sign = ParagraphStyle("S", fontName=_FONT_REG, fontSize=9, alignment=TA_LEFT, leading=12)
    style_sign_r = ParagraphStyle("SR", fontName=_FONT_REG, fontSize=9, alignment=TA_RIGHT, leading=12)

    story = []

    # Логотип
    if assets.logo_path and Path(assets.logo_path).exists():
        logo = Image(str(assets.logo_path), width=20 * mm, height=20 * mm)
        logo.hAlign = "CENTER"
        story.append(logo)
        story.append(Spacer(1, 1 * mm))

    # Реквизиты компании
    company_title = assets.company_name or "ООО «Компания»"
    story.append(Paragraph(company_title, style_company_name))

    if assets.company_details:
        details_html = "<br/>".join(assets.company_details)
    else:
        # Минимальные реквизиты, если бланк не передан
        details_html = "г. Санкт-Петербург"
    story.append(Paragraph(details_html, style_company))
    story.append(Spacer(1, 2.5 * mm))

    story.append(Paragraph(f"Исх. {ref_number} от {date.today().strftime('%d.%m.%Y')} года", style_ref))
    story.append(Paragraph("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", style_title))

    story.append(Paragraph(
        f"На Ваш запрос направляем коммерческое предложение на поставку материалов "
        f"и оборудования ({subject}).",
        style_body,
    ))
    story.append(Paragraph(
        f"Срок поставки материалов и оборудования составляет до {delivery_days} календарных дней "
        f"при условии предоставления Заказчиком всех необходимых исходных данных "
        f"и согласования спецификации.",
        style_body,
    ))
    story.append(Paragraph(
        "В случае несвоевременного предоставления со стороны Заказчика исходных данных "
        "или согласования новых технических решений, срок поставки будет прямо пропорционально "
        "сдвигаться.",
        style_body,
    ))
    story.append(Paragraph(
        f"Необходимое авансирование для выполнения поставки — до {advance_percent}% стоимости, "
        f"после подписания Договора.",
        style_body,
    ))
    story.append(Paragraph(
        f"Срок действия коммерческого предложения: {validity_days} календарных дней.",
        style_body,
    ))
    story.append(Paragraph("Спецификация товара (работ, услуг):", style_body_bold))
    story.append(Spacer(1, 1 * mm))

    # Таблица
    header = [
        Paragraph("№ п/п", style_th),
        Paragraph("Наименование, характеристики товара (работ, услуг)", style_th),
        Paragraph("Ед.изм.", style_th),
        Paragraph("Объем товаров (работ, услуг)", style_th),
        Paragraph(f"Цена за единицу, руб. (с наценкой {markup})", style_th),
        Paragraph("Сумма, руб.", style_th),
        Paragraph("Страна происхождения товара", style_th),
    ]
    data = [header]
    grand_total = 0.0

    for i, item in enumerate(estimate.items, 1):
        unit_price, amount = item.with_markup(markup)
        grand_total += amount
        name = item.name if len(item.name) <= 105 else item.name[:102] + "..."
        data.append([
            Paragraph(str(i), style_tcen),
            Paragraph(name, style_tc),
            Paragraph(item.unit, style_tcen),
            Paragraph(_fmt_qty(item.qty), style_tn),
            Paragraph(_fmt_num(unit_price), style_tn),
            Paragraph(_fmt_num(amount), style_tn),
            Paragraph("РФ", style_tcen),
        ])

    data.append([
        "",
        Paragraph("<b>ИТОГО</b>", style_th),
        "", "", "",
        Paragraph(f"<b>{_fmt_num(grand_total)}</b>", style_th),
        "",
    ])

    col_widths = [9 * mm, 80 * mm, 13 * mm, 17 * mm, 28 * mm, 23 * mm, 20 * mm]
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.91, 0.91, 0.94)),
        ("BACKGROUND", (0, -1), (-1, -1), colors.Color(0.92, 0.92, 0.92)),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.Color(0.35, 0.35, 0.35)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 1.2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1.2),
    ]))
    story.append(table)
    story.append(Spacer(1, 3 * mm))

    vat_note = (
        " НДС начисляется дополнительно в соответствии с действующим законодательством."
        if include_vat_note else ""
    )
    story.append(Paragraph(
        f"Всего {_fmt_num(grand_total)} руб. "
        f"(цены указаны с наценкой коэффициента {markup} к сметным ценам;{vat_note})",
        style_total,
    ))

    story.append(Spacer(1, 10 * mm))

    # Блок подписи и печати
    left = Paragraph(
        "С уважением,<br/>Генеральный директор<br/><br/>________________",
        style_sign,
    )
    right = Paragraph(f"<br/><br/><br/>{dir_name}", style_sign_r)

    # Печать: либо из бланка, либо плейсхолдер «М.П.»
    if assets.stamp_path and Path(assets.stamp_path).exists():
        stamp_flowable = Image(str(assets.stamp_path), width=35 * mm, height=30 * mm)
    else:
        # Круглый плейсхолдер места печати
        stamp_style = ParagraphStyle(
            "StampPh",
            fontName=_FONT_REG,
            fontSize=9,
            alignment=TA_CENTER,
            leading=12,
            textColor=colors.Color(0.45, 0.45, 0.45),
        )
        stamp_flowable = Paragraph(
            "┌──────────┐<br/>│&nbsp;&nbsp;&nbsp;М.П.&nbsp;&nbsp;&nbsp;│<br/>└──────────┘",
            stamp_style,
        )

    sig_data = [[left, stamp_flowable, right]]
    col_w = [60 * mm, 50 * mm, 60 * mm]

    sig_table = Table(sig_data, colWidths=col_w)
    sig_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, 0), "LEFT"),
        ("ALIGN", (1, 0), (1, 0), "CENTER"),
        ("ALIGN", (2, 0), (2, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(sig_table)

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=8 * mm,
        bottomMargin=10 * mm,
    )
    doc.build(story)
    return output_path
