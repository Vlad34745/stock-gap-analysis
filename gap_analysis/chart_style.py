"""
Shared chart styling helpers so every chart in the report matches the
brand theme (Segoe UI font, teal/green/red accents) instead of Excel's
plain defaults. Used by report.py for both the per-ticker gap charts
and the Summary comparison chart.
"""
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    Paragraph, ParagraphProperties, CharacterProperties, Font as DrawingFont, RichTextProperties
)

from .config import ACCENT_TEAL, GRID_GRAY


def _title_from_text(text, cp, pp):
    from openpyxl.chart.title import Title
    from openpyxl.chart.text import Text
    from openpyxl.drawing.text import RegularTextRun
    run = RegularTextRun(rPr=cp, t=text)
    para = Paragraph(pPr=pp, r=[run])
    rich = RichText(bodyPr=RichTextProperties(), p=[para])
    return Title(tx=Text(rich=rich))


def styled_title(text: str, size: int = 1400, color: str = ACCENT_TEAL):
    """Build a RichText chart title in the report's font/color instead of Excel's plain default."""
    cp = CharacterProperties(sz=size, b=True, solidFill=color, latin=DrawingFont(typeface="Segoe UI"))
    pp = ParagraphProperties(defRPr=cp)
    return _title_from_text(text, cp, pp)


def style_axis_text(axis, size: int = 900, color: str = "424242", rot: int = -2700000):
    """Style an axis's tick labels. `rot` is in 60,000ths of a degree (negative tilts up-left)."""
    cp = CharacterProperties(sz=size, solidFill=color, latin=DrawingFont(typeface="Segoe UI"))
    pp = ParagraphProperties(defRPr=cp)
    axis.txPr = RichText(bodyPr=RichTextProperties(rot=rot, vert="horz"), p=[Paragraph(pPr=pp, r=[])])


def style_axis_title(axis, text: str, size: int = 1000, color: str = ACCENT_TEAL):
    cp = CharacterProperties(sz=size, b=True, solidFill=color, latin=DrawingFont(typeface="Segoe UI"))
    pp = ParagraphProperties(defRPr=cp)
    axis.title = _title_from_text(text, cp, pp)


def light_gridlines():
    return ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=GRID_GRAY, w=6350)))
