# -*- coding: utf-8 -*-
"""توليد ملفات Word و PDF مع دعم الاتجاهين وكتل الكود وصفحة المصادر."""

import io
import os
import re

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from config import (
    logger, UNIVERSITY, COLLEGE, DEPARTMENT, STAGE, academic_year,
    ARABIC_FONT, LATIN_FONT, MONO_FONT, AR_TTF, AR_TTF_BOLD, MONO_TTF,
)
from text_utils import iter_doc_blocks, has_arabic

try:
    from fpdf import FPDF
    FPDF_OK = True
except Exception:
    FPDF_OK = False

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    BIDI_OK = True
except Exception:
    BIDI_OK = False


_REF_HEADINGS = {
    "المصادر", "المراجع", "المصادر والمراجع", "المراجع والمصادر",
    "قائمة المصادر", "قائمة المراجع", "قائمة المصادر والمراجع",
    "references", "sources", "bibliography", "references and sources",
}


def _is_references_heading(text: str) -> bool:
    """يتحقق مما إذا كان العنوان الفرعي هو عنوان صفحة المصادر."""
    cleaned = re.sub(r"[:：\-–—\.\s]+", " ", (text or "").strip()).strip().lower()
    return cleaned in _REF_HEADINGS


# ══════════════════════════════ ملف Word ══════════════════════════════

def _set_run_style(run, lang: str, size: int = 13, bold: bool = False,
                   color: tuple[int, int, int] | None = None,
                   font_name: str | None = None):
    name = font_name or (ARABIC_FONT if lang == "ar" else LATIN_FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name)
    if lang == "ar":
        rtl = OxmlElement("w:rtl")
        rtl.set(qn("w:val"), "1")
        rPr.append(rtl)
        szcs = OxmlElement("w:szCs")
        szcs.set(qn("w:val"), str(size * 2))
        rPr.append(szcs)


def _set_para_dir(paragraph, lang: str):
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1" if lang == "ar" else "0")
    pPr.append(bidi)


def _add_par(doc, text: str, lang: str, size: int = 13, bold: bool = False,
             align=None, space_after: int = 8, color=None,
             font_name: str | None = None, line_spacing: float = 1.5):
    p = doc.add_paragraph()
    _set_para_dir(p, lang)
    p.alignment = align if align is not None else WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    run = p.add_run(text)
    _set_run_style(run, lang, size=size, bold=bold, color=color, font_name=font_name)
    return p


def _add_code_block(doc, lines: list[str]) -> None:
    """كتلة كود: خط ثابت العرض، يسار، سطر واحد، والإزاحة كما هي."""
    for ln in lines:
        _add_par(doc, ln if ln.strip() else "\u00a0", "en", size=10,
                 align=WD_ALIGN_PARAGRAPH.LEFT, space_after=0,
                 font_name=MONO_FONT, line_spacing=1.0)


def create_academic_docx(title: str, body: str, subject_name: str, lang: str,
                         with_cover: bool = False, student: str = "") -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1.1)
    sec.right_margin = Inches(1.1)

    center = WD_ALIGN_PARAGRAPH.CENTER
    head_align = WD_ALIGN_PARAGRAPH.RIGHT if lang == "ar" else WD_ALIGN_PARAGRAPH.LEFT

    if with_cover:
        doc.add_paragraph()
        _add_par(doc, UNIVERSITY, "ar", size=20, bold=True, align=center, space_after=6)
        _add_par(doc, COLLEGE, "ar", size=15, align=center, space_after=4)
        _add_par(doc, DEPARTMENT, "ar", size=14, align=center, space_after=40)
        _add_par(doc, title, lang, size=24, bold=True, align=center,
                 space_after=30, color=(0x1F, 0x3B, 0x73))
        _add_par(doc, subject_name, "ar", size=14, align=center, space_after=6)
        _add_par(doc, STAGE, "ar", size=13, align=center, space_after=6)
        if student:
            _add_par(doc, student, "ar", size=13, align=center, space_after=6)
        _add_par(doc, academic_year(), "ar", size=13, align=center, space_after=0)
        doc.add_page_break()
    else:
        _add_par(doc, title, lang, size=19, bold=True, align=center,
                 space_after=18, color=(0x1F, 0x3B, 0x73))

    has_content = False
    in_refs = False
    for kind, item in iter_doc_blocks(body):
        if kind == "code":
            _add_code_block(doc, item)
            has_content = True
        elif kind == "head":
            if _is_references_heading(item):
                if has_content:
                    doc.add_page_break()
                in_refs = True
                _add_par(doc, item, lang, size=16, bold=True,
                         align=center, space_after=14, color=(0x1F, 0x3B, 0x73))
            else:
                in_refs = False
                _add_par(doc, item, lang, size=15, bold=True,
                         align=head_align, space_after=6, color=(0x1F, 0x3B, 0x73))
            has_content = True
        else:
            _add_par(
                doc, item, lang,
                size=12 if in_refs else 13,
                align=head_align if in_refs else None,
                space_after=10 if in_refs else 8,
            )
            has_content = True

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ══════════════════════════════ ملف PDF ══════════════════════════════
# ملاحظة: الملف قد يخلط العربية واللاتينية (عنوان عربي مع متن إنكليزي)،
# لذلك يُسجَّل خط يونيكودي واحد لكل الأنماط، ويُحدَّد اتجاه كل فقرة وحدها.

PDF_FAMILY = "doc"
PDF_MONO = "mono"
_PDF_MAX_WORD = 42

# رموز شائعة قد لا يحتويها الخط، تُستبدل ببدائل آمنة
_PDF_MAP = {
    "\u2014": "-", "\u2013": "-", "\u2012": "-", "\u2212": "-", "\u2015": "-",
    "\u201c": '"', "\u201d": '"', "\u201e": '"',
    "\u2018": "'", "\u2019": "'", "\u2032": "'", "\u2033": '"',
    "\u2026": "...", "\u2022": "-", "\u25cf": "-", "\u25aa": "-", "\u25a0": "-",
    "\u2192": "->", "\u2190": "<-", "\u21d2": "=>", "\u21d4": "<=>",
    "\u221a": "sqrt", "\u221e": "inf", "\u2248": "~", "\u2260": "!=",
    "\u2264": "<=", "\u2265": ">=", "\u2211": "sum", "\u222b": "integral",
    "\u2208": "in", "\u2229": "and", "\u222a": "or", "\u00ac": "not",
    "\u03b1": "alpha", "\u03b2": "beta", "\u03b8": "theta", "\u03c0": "pi",
    "\u03bc": "mu", "\u03c3": "sigma", "\u03bb": "lambda", "\u0394": "Delta",
    "\u03a3": "Sigma", "\u03b3": "gamma", "\u03c1": "rho", "\u03c6": "phi",
    "\u00a0": " ", "\u202f": " ", "\u2009": " ", "\u200a": " ",
    "\u200b": "", "\u200c": "", "\u200d": "", "\u200e": "", "\u200f": "",
    "\u202a": "", "\u202b": "", "\u202c": "", "\ufeff": "",
    "\u2028": " ", "\u2029": " ", "\t": "    ",
}


def sanitize_for_pdf(text: str, keep_spaces: bool = False) -> str:
    """يستبدل الرموز غير المدعومة، يحذف الإيموجي، ويقطّع الكلمات الطويلة جداً.

    keep_spaces=True لكتل الكود: لا تُدمج الفراغات المتتالية فتبقى الإزاحة.
    """
    if not text:
        return ""
    for bad, good in _PDF_MAP.items():
        text = text.replace(bad, good)

    keep = []
    for ch in text:
        o = ord(ch)
        if ch in "\n\r":
            keep.append(ch)
        elif o < 0x20:
            keep.append(" ")
        elif 0x2300 <= o <= 0x2BFF or 0xE000 <= o <= 0xF8FF or o >= 0x1F000:
            continue                      # رسوم ورموز وإيموجي
        elif 0xFE00 <= o <= 0xFE0F:
            continue                      # محدّدات شكل الإيموجي
        else:
            keep.append(ch)
    text = "".join(keep)

    parts = []
    for word in text.split(" "):
        while len(word) > _PDF_MAX_WORD:
            parts.append(word[:_PDF_MAX_WORD])
            word = word[_PDF_MAX_WORD:]
        parts.append(word)
    text = " ".join(parts)
    if keep_spaces:
        return text
    return re.sub(r"[ ]{3,}", "  ", text)


def _hard_clean(text: str) -> str:
    """آخر خط دفاع: لاتيني أساسي وعربي فقط."""
    keep = []
    for ch in text:
        o = ord(ch)
        if ch in "\n\r" or 0x20 <= o <= 0x7E or 0x0600 <= o <= 0x06FF \
                or 0xFB50 <= o <= 0xFEFF:
            keep.append(ch)
    return "".join(keep) or " "


def _pdf_ready(lang: str = "ar") -> tuple[bool, str]:
    if not FPDF_OK:
        return False, "مكتبة fpdf2 غير مثبّتة على الخادم."
    if not os.path.exists(AR_TTF):
        return False, f"ملف الخط غير موجود في المسار {AR_TTF}."
    return True, ""


class _ReportPDF(FPDF):
    def __init__(self, lang: str):
        super().__init__(format="A4", unit="mm")
        self.lang = lang
        self.base_font = PDF_FAMILY
        self.set_margins(20, 20, 20)
        self.set_auto_page_break(True, margin=20)

        # خط واحد يونيكودي لكل الأنماط، فلا رجوع إلى helvetica مطلقاً
        bold_path = AR_TTF_BOLD if os.path.exists(AR_TTF_BOLD) else AR_TTF
        self.add_font(PDF_FAMILY, "", AR_TTF)
        self.add_font(PDF_FAMILY, "B", bold_path)
        self.add_font(PDF_FAMILY, "I", AR_TTF)
        self.add_font(PDF_FAMILY, "BI", bold_path)

        # خط ثابت العرض للكود إن وُجد، وإلا يُستخدم الخط الأساسي
        self.mono_ok = False
        if os.path.exists(MONO_TTF):
            try:
                self.add_font(PDF_MONO, "", MONO_TTF)
                self.add_font(PDF_MONO, "B", MONO_TTF)
                self.mono_ok = True
            except Exception as e:
                logger.warning("mono font not loaded: %s", e)

        self.shaped = False
        self._dir = None
        try:
            self.set_text_shaping(True, direction="ltr", script="latn", language="eng")
            self.shaped = True
            self._dir = "ltr"
        except Exception as e:
            logger.warning("text shaping unavailable: %s", e)

        self.use(12)

    def use(self, size: int, bold: bool = False):
        self.set_font(self.base_font, "B" if bold else "", size)

    def use_mono(self, size: int):
        self.set_font(PDF_MONO if self.mono_ok else self.base_font, "", size)

    def _set_dir(self, rtl: bool) -> None:
        if not self.shaped:
            return
        want = "rtl" if rtl else "ltr"
        if self._dir == want:
            return
        try:
            if rtl:
                self.set_text_shaping(True, direction="rtl",
                                      script="arab", language="ara")
            else:
                self.set_text_shaping(True, direction="ltr",
                                      script="latn", language="eng")
            self._dir = want
        except Exception as e:
            logger.warning("set direction failed: %s", e)
            self.shaped = False

    def _prep(self, text: str, rtl: bool) -> str:
        if not rtl or self.shaped:
            return text
        if BIDI_OK:
            try:
                return get_display(arabic_reshaper.reshape(text))
            except Exception:
                return text
        return text

    def _wrap(self, text: str, rtl: bool) -> list[str]:
        avail = self.w - self.l_margin - self.r_margin
        lines, cur = [], ""
        for word in text.split(" "):
            trial = word if not cur else cur + " " + word
            if self.get_string_width(self._prep(trial, rtl)) <= avail or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)
        return lines or [""]

    def _render(self, text: str, rtl: bool, center: bool, lh: float,
                code: bool = False, justify: bool = True) -> None:
        self._set_dir(False if code else rtl)
        if code:
            self.multi_cell(0, lh, text, align="L", new_x="LMARGIN", new_y="NEXT")
            return
        if self.shaped:
            align = "C" if center else ("R" if rtl else ("J" if justify else "L"))
            self.multi_cell(0, lh, text, align=align, new_x="LMARGIN", new_y="NEXT")
        else:
            align = "C" if center else ("R" if rtl else "L")
            for line in self._wrap(text, rtl):
                self.cell(0, lh, self._prep(line, rtl), align=align,
                          new_x="LMARGIN", new_y="NEXT")

    def par(self, text: str, size: int = 12, bold: bool = False,
            center: bool = False, gap: float = 3.0, rgb=None,
            code: bool = False, justify: bool = True):
        text = sanitize_for_pdf(text, keep_spaces=code)
        if not text.strip() and not code:
            return
        rtl = False if code else has_arabic(text)
        if code:
            self.use_mono(size)
        else:
            self.use(size, bold)
        self.set_text_color(*(rgb or (0, 0, 0)))
        lh = max(4.5, size * 0.62)
        try:
            self._render(text, rtl, center, lh, code=code, justify=justify)
        except Exception as e:
            logger.warning("pdf paragraph failed (%s) — retrying cleaned", e)
            try:
                self._render(_hard_clean(text), rtl, center, lh,
                             code=code, justify=justify)
            except Exception as e2:
                logger.warning("pdf paragraph dropped: %s", e2)
        self.ln(gap)
        self.set_text_color(0, 0, 0)


def create_academic_pdf(title: str, body: str, subject_name: str, lang: str,
                        with_cover: bool = False, student: str = "") -> bytes:
    pdf = _ReportPDF(lang)
    pdf.add_page()

    if with_cover:
        pdf.ln(25)
        pdf.par(UNIVERSITY, size=22, bold=True, center=True, gap=3)
        pdf.par(COLLEGE, size=15, center=True, gap=2)
        pdf.par(DEPARTMENT, size=14, center=True, gap=35)
        pdf.par(title, size=24, bold=True, center=True, gap=25, rgb=(31, 59, 115))
        pdf.par(subject_name, size=14, center=True, gap=3)
        pdf.par(STAGE, size=13, center=True, gap=3)
        if student:
            pdf.par(student, size=13, center=True, gap=3)
        pdf.par(academic_year(), size=13, center=True, gap=0)
        pdf.add_page()
    else:
        pdf.par(title, size=19, bold=True, center=True, gap=8, rgb=(31, 59, 115))

    has_content = False
    in_refs = False
    for kind, item in iter_doc_blocks(body):
        if kind == "code":
            pdf.ln(1)
            for ln in item:
                pdf.par(ln if ln.strip() else " ", size=9, gap=0, code=True)
            pdf.ln(3)
            has_content = True
        elif kind == "head":
            if _is_references_heading(item):
                if has_content:
                    pdf.add_page()
                in_refs = True
                pdf.ln(4)
                pdf.par(item, size=16, bold=True, center=True, gap=6, rgb=(31, 59, 115))
            else:
                in_refs = False
                pdf.ln(2)
                pdf.par(item, size=15, bold=True, gap=2, rgb=(31, 59, 115))
            has_content = True
        else:
            pdf.par(
                item,
                size=11 if in_refs else 12,
                gap=4 if in_refs else 3,
                justify=not in_refs,
            )
            has_content = True

    return bytes(pdf.output())
