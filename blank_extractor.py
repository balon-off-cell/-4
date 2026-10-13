# -*- coding: utf-8 -*-
"""
Извлечение логотипа, печати и подписи из бланка (PDF) компании.
"""

from __future__ import annotations

import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from PIL import Image


@dataclass
class CompanyAssets:
    """Ресурсы, вырезанные из бланка."""
    logo_path: Optional[Path] = None
    stamp_path: Optional[Path] = None          # печать (круглая)
    signature_path: Optional[Path] = None      # подпись (если отдельным изображением)
    company_name: str = ""
    company_details: List[str] = None          # адрес, ОГРН, ИНН и т.д.
    director_name: str = ""
    raw_images: List[Path] = None

    def __post_init__(self):
        if self.company_details is None:
            self.company_details = []
        if self.raw_images is None:
            self.raw_images = []


def extract_images_from_pdf(pdf_path: str | Path, out_dir: str | Path) -> List[Path]:
    """
    Извлекает все встроенные изображения из PDF через pdfimages (poppler).
    Возвращает список путей к файлам.
    """
    import subprocess
    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    prefix = out_dir / "img"
    # -all сохраняет jpeg/png как есть
    subprocess.run(
        ["pdfimages", "-all", str(pdf_path), str(prefix)],
        check=False,
        capture_output=True,
    )
    files = sorted(out_dir.glob("img-*"))
    return list(files)


def _apply_smask(rgb_path: Path, mask_path: Path, out_path: Path) -> Path:
    """Накладывает soft-mask (альфа) на RGB-изображение."""
    rgb = Image.open(rgb_path).convert("RGBA")
    mask = Image.open(mask_path).convert("L")
    if mask.size != rgb.size:
        mask = mask.resize(rgb.size, Image.Resampling.LANCZOS)
    rgb.putalpha(mask)
    rgb.save(out_path)
    return out_path


def extract_company_assets(
    blank_pdf: str | Path,
    work_dir: Optional[str | Path] = None,
) -> CompanyAssets:
    """
    Главная функция: вырезает логотип и печать из бланка КП.

    Эвристики:
    - логотип обычно первое крупное квадратное изображение (на 1-й странице);
    - печать — изображение с soft-mask, часто круглое / синее, на последней странице.
    """
    blank_pdf = Path(blank_pdf)
    if work_dir is None:
        work_dir = Path(tempfile.mkdtemp(prefix="kp_blank_"))
    else:
        work_dir = Path(work_dir)
        work_dir.mkdir(parents=True, exist_ok=True)

    assets = CompanyAssets()
    images = extract_images_from_pdf(blank_pdf, work_dir / "raw")
    assets.raw_images = images

    if not images:
        return assets

    # Собираем пары image + smask (poppler даёт img-000.jpg + img-001.png и т.д.)
    pairs: List[Tuple[Path, Optional[Path]]] = []
    i = 0
    while i < len(images):
        img = images[i]
        mask = None
        if i + 1 < len(images):
            nxt = images[i + 1]
            # smask обычно grayscale png рядом
            if nxt.suffix.lower() == ".png":
                try:
                    m = Image.open(nxt)
                    if m.mode in ("L", "LA", "1"):
                        mask = nxt
                        i += 1
                except Exception:
                    pass
        pairs.append((img, mask))
        i += 1

    prepared: List[Tuple[Path, int, int, bool]] = []  # path, w, h, has_alpha
    for idx, (img_path, mask_path) in enumerate(pairs):
        try:
            im = Image.open(img_path)
            w, h = im.size
        except Exception:
            continue
        out = work_dir / f"asset_{idx}.png"
        has_alpha = False
        if mask_path:
            _apply_smask(img_path, mask_path, out)
            has_alpha = True
        else:
            im.convert("RGBA").save(out)
        prepared.append((out, w, h, has_alpha))

    if not prepared:
        return assets

    # Логотип: первое относительно квадратное изображение без сильного альфа-приоритета
    # или просто первое
    logo_candidate = prepared[0]
    for p, w, h, alpha in prepared:
        ratio = w / max(h, 1)
        if 0.7 <= ratio <= 1.4 and w >= 80:
            logo_candidate = (p, w, h, alpha)
            break
    assets.logo_path = logo_candidate[0]

    # Печать: предпочитаем изображение с альфа-каналом (типично для печати)
    # и не слишком маленькое; обычно последнее
    stamp_candidates = [x for x in prepared if x[3] and x[1] >= 100]
    if not stamp_candidates:
        stamp_candidates = [x for x in prepared if x[1] >= 120]
    if stamp_candidates:
        # берём последнее (печать в конце документа)
        assets.stamp_path = stamp_candidates[-1][0]
    elif len(prepared) > 1:
        assets.stamp_path = prepared[-1][0]

    # Текстовые реквизиты из PDF (опционально)
    try:
        import pdfplumber
        import re
        with pdfplumber.open(blank_pdf) as pdf:
            if pdf.pages:
                page_text = pdf.pages[0].extract_text() or ""
                lines = [ln.strip() for ln in page_text.splitlines() if ln.strip()]
                # Название компании — короткая строка вида «ООО «…»»
                for ln in lines[:12]:
                    low = ln.lower()
                    if re.match(r"^(ооо|ао|ип|пао|зао)\b", low) and len(ln) < 80:
                        assets.company_name = ln
                        break
                for ln in lines[:20]:
                    low = ln.lower()
                    if any(k in low for k in ("огрн", "инн", "кпп", "тел", "@")):
                        if ln not in assets.company_details:
                            assets.company_details.append(ln)
                    elif re.match(r"^\d{6},?\s*г\.", low) or "улица" in low or "пр-т" in low:
                        if ln not in assets.company_details:
                            assets.company_details.insert(0, ln)
                last_text = pdf.pages[-1].extract_text() or ""
                m = re.search(r"([А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.)", last_text)
                if m:
                    assets.director_name = m.group(1)
    except Exception:
        pass

    return assets
