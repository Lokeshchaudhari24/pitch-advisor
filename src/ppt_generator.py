from pathlib import Path
from copy import deepcopy
from io import BytesIO
import re
from urllib.request import Request, urlopen

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt, Inches


# -----------------------------------------------------------------------------
# Template-based PowerPoint generator for the Marsh AI Pitch Advisor.
#
# The template controls ALL visual design. This module only fills placeholders.
# It is intentionally defensive about long text so generated content stays
# inside its boxes instead of overflowing.
# -----------------------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "Marsh_AI_Pitch_Template.pptx"
LOGO_PATH = PROJECT_ROOT / "assets" / "marsh_logo.png"
BUSINESS_IMAGE_URL = (
    "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab"
    "?auto=format&fit=crop&w=1400&q=85"
)
_BUSINESS_IMAGE_BYTES = None

# Safe font-size ranges for dynamic content.
FONT_MIN = {
    "short": 10,
    "medium": 8.5,
    "long": 7.5,
}

MARSH_RED = RGBColor(0xC8, 0x0A, 0x2E)
INK = RGBColor(0x20, 0x2B, 0x35)
MUTED = RGBColor(0x6B, 0x75, 0x7E)
LINE = RGBColor(0xDB, 0xDE, 0xE1)
PALE_RED = RGBColor(0xFC, 0xF0, 0xF2)
PALE_GRAY = RGBColor(0xF5, 0xF6, 0xF7)


def _clean(value):
    """Convert a value to safe display text."""
    if value is None:
        return "—"
    text = str(value).strip()
    return text if text else "—"


def _shorten(text, max_chars):
    """Soft truncate only when a field is extremely long."""
    text = _clean(text)
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 1].rstrip() + "…"


def _display_document_name(value):
    text = str(value or "").strip()
    if not text:
        return ""
    path = Path(text)
    return path.stem if path.suffix.lower() == ".pdf" else path.name


def _set_font_size(shape, size):
    if not hasattr(shape, "text_frame"):
        return
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size)


def _fit_text(shape, text, min_size=8, max_size=None, bold=False):
    """Replace text and reduce font size for long dynamic content.

    The template has deliberately generous text boxes. We use a predictable
    length-based sizing strategy rather than changing the template geometry.
    """
    if not hasattr(shape, "text_frame"):
        return

    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True

    paragraph = tf.paragraphs[0]
    run = paragraph.add_run()
    run.text = _clean(text)
    run.font.name = "Aptos"
    run.font.bold = bold

    # Start from the existing template size when available.
    base_size = max_size or 11
    char_count = len(run.text)

    if char_count > 180:
        size = max(min_size, base_size - 3.0)
    elif char_count > 120:
        size = max(min_size, base_size - 2.0)
    elif char_count > 80:
        size = max(min_size, base_size - 1.0)
    else:
        size = base_size

    run.font.size = Pt(size)


def _replace_in_text_shape(shape, replacements):
    """Replace placeholders in a normal text shape while preserving its box."""
    if not hasattr(shape, "text_frame"):
        return False

    original = shape.text
    if not original:
        return False

    updated = original
    found = False
    for key, value in replacements.items():
        if key in updated:
            replacement = "" if value is None else str(value)
            updated = updated.replace(key, replacement)
            found = True

    if not found:
        return False

    # Rebuild text so placeholders that were split across runs are also handled.
    tf = shape.text_frame
    original_sizes = [
        run.font.size.pt
        for paragraph in tf.paragraphs
        for run in paragraph.runs
        if run.font.size is not None
    ]
    base_size = max(original_sizes, default=11)
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = updated
    r.font.name = "Aptos"

    # Preserve the template's intended hierarchy, shrinking only long text.
    if len(updated) > 160:
        size = max(FONT_MIN["long"], base_size - 3)
    elif len(updated) > 100:
        size = max(FONT_MIN["long"], base_size - 2)
    elif len(updated) > 65:
        size = max(FONT_MIN["long"], base_size - 1)
    else:
        size = base_size
    r.font.size = Pt(size)
    return True


def _replace_table_cells(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            original = cell.text
            original_sizes = [
                run.font.size.pt
                for paragraph in cell.text_frame.paragraphs
                for run in paragraph.runs
                if run.font.size is not None
            ]
            base_size = max(original_sizes, default=11)
            updated = original
            found = False
            for key, value in replacements.items():
                if key in updated:
                    replacement = "" if value is None else str(value)
                    updated = updated.replace(key, replacement)
                    found = True

            if not found:
                continue

            cell.text = updated
            tf = cell.text_frame
            tf.word_wrap = True
            for p in tf.paragraphs:
                for run in p.runs:
                    run.font.name = "Aptos"
                    run.font.size = Pt(
                        max(FONT_MIN["medium"], base_size - 1.5)
                        if len(updated) > 75
                        else base_size
                    )


def _hide_unused_risk_row(slide, risk_number):
    placeholders = {
        f"{{risk_{risk_number}}}",
        f"{{risk_{risk_number}_detail}}",
    }
    anchors = [
        shape
        for shape in slide.shapes
        if getattr(shape, "text", "") in placeholders
    ]
    if not anchors:
        return

    row_top = min(shape.top for shape in anchors) - Inches(0.1)
    row_bottom = max(shape.top + shape.height for shape in anchors) + Inches(0.1)
    for shape in list(slide.shapes):
        if row_top <= shape.top and shape.top + shape.height <= row_bottom:
            shape._element.getparent().remove(shape._element)


def _hide_unused_claim_row(slide, claim_number):
    """Remove an unused audit row, including its card and text shapes."""
    placeholders = {
        f"{{claim_{claim_number}_text}}",
        f"{{claim_{claim_number}_status}}",
        f"{{claim_{claim_number}_source}}",
        f"{{claim_{claim_number}_page}}",
    }
    anchors = [
        shape
        for shape in slide.shapes
        if any(token in getattr(shape, "text", "") for token in placeholders)
        or getattr(shape, "text", "").strip() == f"CLAIM {claim_number:02d}"
    ]
    if not anchors:
        return

    row_top = min(shape.top for shape in anchors) - Inches(0.1)
    row_bottom = max(shape.top + shape.height for shape in anchors) + Inches(0.1)
    for shape in list(slide.shapes):
        if row_top <= shape.top and shape.top + shape.height <= row_bottom:
            shape._element.getparent().remove(shape._element)


def _add_workflow_text(slide, text, x, y, width, height, size, color,
                       bold=False, font="Aptos", align=PP_ALIGN.LEFT):
    shape = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(width), Inches(height)
    )
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.margin_left = Pt(0)
    frame.margin_right = Pt(0)
    frame.margin_top = Pt(0)
    frame.margin_bottom = Pt(0)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return shape


def _add_workflow_card(slide, index, x, title, subtitle, description, icon):
    card_y = 2.68
    card_w = 2.31
    card_h = 2.85
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(card_y), Inches(card_w), Inches(card_h),
    )
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    card.line.color.rgb = LINE
    card.line.width = Pt(0.8)

    badge = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(x + card_w / 2 - 0.17), Inches(card_y - 0.17),
        Inches(0.34), Inches(0.34),
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = MARSH_RED
    badge.line.fill.background()
    _add_workflow_text(
        slide, f"{index:02d}", x + card_w / 2 - 0.17, card_y - 0.17,
        0.34, 0.34, 10, RGBColor(0xFF, 0xFF, 0xFF), True,
        align=PP_ALIGN.CENTER,
    )

    icon_size = 0.88
    icon_x = x + (card_w - icon_size) / 2
    icon_y = 2.96
    icon_back = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(icon_x), Inches(icon_y),
        Inches(icon_size), Inches(icon_size),
    )
    icon_back.fill.solid()
    icon_back.fill.fore_color.rgb = PALE_RED
    icon_back.line.fill.background()

    if icon in {"client", "policy", "audit"}:
        doc_x = icon_x + 0.29
        doc_y = icon_y + 0.22
        doc = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(doc_x), Inches(doc_y),
            Inches(0.30), Inches(0.42),
        )
        doc.fill.solid()
        doc.fill.fore_color.rgb = PALE_RED
        doc.line.color.rgb = INK
        doc.line.width = Pt(1.8)
        for line_index in range(3 if icon != "client" else 2):
            mark = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(doc_x + 0.06), Inches(doc_y + 0.11 + line_index * 0.09),
                Inches(0.17), Inches(0.018),
            )
            mark.fill.solid()
            mark.fill.fore_color.rgb = MARSH_RED
            mark.line.fill.background()
        if icon == "policy":
            back_doc = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(doc_x + 0.13), Inches(doc_y + 0.08),
                Inches(0.30), Inches(0.42),
            )
            back_doc.fill.solid()
            back_doc.fill.fore_color.rgb = PALE_RED
            back_doc.line.color.rgb = INK
            back_doc.line.width = Pt(1.8)
            back_doc._element.addprevious(doc._element)
        elif icon == "audit":
            lens = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(doc_x + 0.19), Inches(doc_y + 0.25),
                Inches(0.24), Inches(0.24),
            )
            lens.fill.solid()
            lens.fill.fore_color.rgb = PALE_RED
            lens.line.color.rgb = MARSH_RED
            lens.line.width = Pt(2)
            handle = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(doc_x + 0.38), Inches(doc_y + 0.46),
                Inches(0.12), Inches(0.035),
            )
            handle.rotation = 45
            handle.fill.solid()
            handle.fill.fore_color.rgb = MARSH_RED
            handle.line.fill.background()
    elif icon == "analysis":
        center = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(icon_x + 0.34), Inches(icon_y + 0.34),
            Inches(0.20), Inches(0.20),
        )
        center.fill.solid()
        center.fill.fore_color.rgb = MARSH_RED
        center.line.fill.background()
        for dx, dy in ((0.13, 0.13), (0.58, 0.13), (0.13, 0.58), (0.58, 0.58)):
            node = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(icon_x + dx), Inches(icon_y + dy),
                Inches(0.15), Inches(0.15),
            )
            node.fill.solid()
            node.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            node.line.color.rgb = INK
            node.line.width = Pt(1.5)
    else:
        head = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(icon_x + 0.35), Inches(icon_y + 0.19),
            Inches(0.18), Inches(0.18),
        )
        head.fill.solid()
        head.fill.fore_color.rgb = PALE_RED
        head.line.color.rgb = INK
        head.line.width = Pt(1.8)
        shoulders = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(icon_x + 0.25), Inches(icon_y + 0.43),
            Inches(0.38), Inches(0.23),
        )
        shoulders.fill.solid()
        shoulders.fill.fore_color.rgb = PALE_RED
        shoulders.line.color.rgb = INK
        shoulders.line.width = Pt(1.8)
        check = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(icon_x + 0.54), Inches(icon_y + 0.48),
            Inches(0.22), Inches(0.22),
        )
        check.fill.solid()
        check.fill.fore_color.rgb = MARSH_RED
        check.line.fill.background()
        _add_workflow_text(
            slide, "✓", icon_x + 0.54, icon_y + 0.47, 0.22, 0.22,
            11, RGBColor(0xFF, 0xFF, 0xFF), True, align=PP_ALIGN.CENTER,
        )

    _add_workflow_text(
        slide, title, x + 0.08, 3.98, card_w - 0.16, 0.34,
        14, INK, True, font="Cambria", align=PP_ALIGN.CENTER,
    )
    _add_workflow_text(
        slide, subtitle, x + 0.10, 4.32, card_w - 0.20, 0.43,
        10.5, MUTED, align=PP_ALIGN.CENTER,
    )

    description_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x + 0.10), Inches(4.81), Inches(card_w - 0.20), Inches(0.54),
    )
    description_box.fill.solid()
    description_box.fill.fore_color.rgb = PALE_GRAY
    description_box.line.fill.background()
    _add_workflow_text(
        slide, description, x + 0.19, 4.84, card_w - 0.38, 0.48,
        10, INK, align=PP_ALIGN.CENTER,
    )


def _redesign_workflow_slide(slide):
    """Replace the vertical workflow with a visual five-stage process."""
    for shape in list(slide.shapes):
        if Inches(1.6) <= shape.top < Inches(7.0):
            shape._element.getparent().remove(shape._element)

    _add_workflow_text(
        slide, "From client risk to confident advice", 0.5, 1.72,
        12.33, 0.34, 20, INK, True, font="Cambria",
    )
    _add_workflow_text(
        slide, "A human-in-the-loop workflow for evidence-grounded insurance pitches.",
        0.5, 2.08, 12.33, 0.28, 12, MUTED,
    )

    stages = [
        ("Client Context", "Company + key\nbusiness risks", "Understand the client's\nprofile and employee-health\nexposures", "client"),
        ("Policy Intelligence", "Policy documents", "Retrieve relevant evidence\nfrom selected policy\ndocuments", "policy"),
        ("AI Analysis", "Risk-to-benefit\nmapping", "Identify relevant benefits\ndirectly from policy\nevidence", "analysis"),
        ("Content Audit", "Claim → source → page", "Verify every claim with\nits exact source and\npage reference", "audit"),
        ("Advisor Decision", "Approve · Edit · Reject", "Human advisor reviews and\nfinalizes client-facing\ncontent", "advisor"),
    ]
    card_w = 2.31
    gap = 0.1875
    start_x = 0.5
    for index, (title, subtitle, description, icon) in enumerate(stages, 1):
        x = start_x + (index - 1) * (card_w + gap)
        _add_workflow_card(slide, index, x, title, subtitle, description, icon)
        if index < len(stages):
            arrow_x = x + card_w + 0.005
            _add_workflow_text(
                slide, "→", arrow_x, 3.32, gap - 0.01, 0.34,
                19, MARSH_RED, True, align=PP_ALIGN.CENTER,
            )

    band = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(6.02), Inches(12.33), Inches(0.82),
    )
    band.fill.solid()
    band.fill.fore_color.rgb = PALE_RED
    band.line.color.rgb = RGBColor(0xF4, 0xD9, 0xDE)
    band.line.width = Pt(0.8)

    gains = [
        ("FASTER", "Reduce manual pitch preparation\nand time to insight"),
        ("TRACEABLE", "Every supported claim links\nto policy evidence"),
        ("CONTROLLED", "Advisor retains final approval\nand control over client content"),
    ]
    band_col_w = 4.11
    for index, (heading, detail) in enumerate(gains):
        x = 0.5 + index * band_col_w
        if index:
            divider = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(x), Inches(6.17), Inches(0.01), Inches(0.52),
            )
            divider.fill.solid()
            divider.fill.fore_color.rgb = RGBColor(0xEE, 0xD2, 0xD7)
            divider.line.fill.background()
        marker = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(x + 0.24), Inches(6.17),
            Inches(0.48), Inches(0.48),
        )
        marker.fill.solid()
        marker.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        marker.line.fill.background()
        _add_workflow_text(
            slide, f"0{index + 1}", x + 0.24, 6.17, 0.48, 0.48,
            11, MARSH_RED, True, align=PP_ALIGN.CENTER,
        )
        _add_workflow_text(
            slide, heading, x + 0.88, 6.13, 2.95, 0.22,
            12, MARSH_RED, True,
        )
        _add_workflow_text(
            slide, detail, x + 0.88, 6.37, 3.00, 0.38,
            10, INK,
        )


def _profile_items(value):
    if isinstance(value, str):
        return [value.strip()] if value.strip() else []
    if isinstance(value, list):
        return [
            item.strip() if isinstance(item, str) else item
            for item in value
            if (isinstance(item, str) and item.strip()) or isinstance(item, dict)
        ]
    return []


def _add_profile_panel(slide, x, y, width, height, title):
    panel = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y), Inches(width), Inches(height),
    )
    panel.fill.solid()
    panel.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    panel.line.color.rgb = LINE
    panel.line.width = Pt(0.8)
    _add_workflow_text(
        slide, title, x + 0.20, y + 0.16, width - 0.40, 0.24,
        10, MARSH_RED, True,
    )
    return panel


def _add_business_image(slide, x, y, width, height):
    global _BUSINESS_IMAGE_BYTES

    if _BUSINESS_IMAGE_BYTES is None:
        try:
            request = Request(
                BUSINESS_IMAGE_URL,
                headers={"User-Agent": "MarshPitchAdvisor/1.0"},
            )
            with urlopen(request, timeout=4) as response:
                _BUSINESS_IMAGE_BYTES = response.read()
        except Exception:
            _BUSINESS_IMAGE_BYTES = b""

    if not _BUSINESS_IMAGE_BYTES:
        _add_profile_panel(slide, x, y, width, height, "BUSINESS CONTEXT")
        _add_workflow_text(
            slide, "Company operating context", x + 0.18, y + 0.45,
            width - 0.36, 0.32, 12, INK, True,
        )
        _add_workflow_text(
            slide, "See the overview and workforce risks on this slide.",
            x + 0.18, y + 0.85, width - 0.36, 0.60, 10, MUTED,
        )
        return

    inset = 0.04
    image_width = width - inset * 2
    image_height = height - inset * 2
    image = slide.shapes.add_picture(
        BytesIO(_BUSINESS_IMAGE_BYTES),
        Inches(x + inset),
        Inches(y + inset),
        width=Inches(image_width),
        height=Inches(image_height),
    )
    source_width, source_height = image.image.size
    source_ratio = source_width / source_height
    frame_ratio = image_width / image_height
    if source_ratio > frame_ratio:
        horizontal_crop = (1 - frame_ratio / source_ratio) / 2
        image.crop_left = horizontal_crop
        image.crop_right = horizontal_crop
    elif source_ratio < frame_ratio:
        vertical_crop = (1 - source_ratio / frame_ratio) / 2
        image.crop_top = vertical_crop
        image.crop_bottom = vertical_crop

    image._element.nvPicPr.cNvPr.set("descr", "Modern commercial office building")


def _redesign_company_overview_slide(slide, overview):
    """Render slide 1 from fields actually returned by the company profile."""
    for shape in slide.shapes:
        if getattr(shape, "text", "").strip() == "Company Overview":
            _replace_in_text_shape(shape, {"Company Overview": "Client Overview"})

    for shape in list(slide.shapes):
        if Inches(1.7) <= shape.top < Inches(7.0):
            shape._element.getparent().remove(shape._element)

    company = str(overview.get("company", "")).strip() or "Client name unavailable"
    industry = str(overview.get("industry", "")).strip() or "Not provided"
    company_size = str(overview.get("company_size", "")).strip() or "Not provided"
    business_overview = str(overview.get("business_overview", "")).strip()
    risks = []
    for item in _profile_items(overview.get("key_risks", [])):
        if isinstance(item, dict):
            title = str(
                item.get("title") or item.get("risk") or item.get("name") or ""
            ).strip()
            description = str(
                item.get("description") or item.get("short_description") or ""
            ).strip()
        else:
            title = str(item).strip()
            description = ""
        if title:
            risks.append({"title": title, "description": description})

    _add_workflow_text(
        slide, "Understanding the client, their business context and key risk exposures.",
        0.5, 1.70, 12.33, 0.28, 11.5, MUTED,
    )

    _add_profile_panel(slide, 0.42, 2.06, 3.20, 2.28, "CLIENT")
    company_size_font = (
        21 if len(company) < 18 else 18 if len(company) < 27
        else 14 if len(company) < 43 else 12
    )
    _add_workflow_text(
        slide, company, 0.65, 2.40, 2.74, 0.56,
        company_size_font, INK, True, font="Cambria",
    )
    separator = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.66), Inches(3.08),
        Inches(2.72), Inches(0.012),
    )
    separator.fill.solid()
    separator.fill.fore_color.rgb = LINE
    separator.line.fill.background()
    _add_workflow_text(
        slide, "INDUSTRY", 0.66, 3.12, 2.72, 0.18,
        8.5, MUTED, True,
    )
    _add_workflow_text(
        slide, industry, 0.66, 3.30, 2.72, 0.36,
        9 if len(industry) > 55 else 10, INK,
    )
    _add_workflow_text(
        slide, "COMPANY SIZE", 0.66, 3.68, 2.72, 0.18,
        8.5, MUTED, True,
    )
    _add_workflow_text(
        slide, company_size, 0.66, 3.86, 2.72, 0.43,
        8.5 if len(company_size) > 56 else 9.5, INK,
    )

    _add_profile_panel(slide, 3.82, 2.06, 6.05, 2.28, "BUSINESS OVERVIEW")
    overview_text = _shorten(
        business_overview or "No business overview was returned in the company profile.",
        540,
    )
    overview_font = 12.5 if len(overview_text) <= 320 else 11.5
    _add_workflow_text(
        slide, overview_text, 4.04, 2.46, 5.62, 1.72,
        overview_font, INK,
    )

    _add_business_image(slide, 10.07, 2.06, 2.84, 2.28)

    risk_panel = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.42), Inches(4.58), Inches(12.49), Inches(2.20),
    )
    risk_panel.fill.solid()
    risk_panel.fill.fore_color.rgb = RGBColor(0xFF, 0xF8, 0xF9)
    risk_panel.line.color.rgb = RGBColor(0xF4, 0xD9, 0xDE)
    risk_panel.line.width = Pt(0.8)
    _add_workflow_text(
        slide, "KEY EMPLOYEE-HEALTH EXPOSURES", 0.66, 4.73,
        5.00, 0.22, 10, MARSH_RED, True,
    )

    if not risks:
        _add_workflow_text(
            slide, "No specific employee-health risks were returned in the profile.",
            0.68, 5.20, 11.80, 0.42, 11, MUTED,
        )
        return

    visible_risks = risks[:6]
    count = len(visible_risks)
    columns = 3 if count > 4 else 2
    column_gap = 0.16
    row_gap = 0.10
    row_height = 0.64 if count > 1 else 1.10
    for index, risk in enumerate(visible_risks):
        row = index // columns
        column = index % columns
        row_columns = min(columns, count - row * columns)
        row_width = (12.01 - column_gap * (row_columns - 1)) / row_columns
        x = 0.66 + column * (row_width + column_gap)
        y = 5.08 + row * (row_height + row_gap)
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(y), Inches(row_width), Inches(row_height),
        )
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        card.line.color.rgb = RGBColor(0xF0, 0xD2, 0xD8)
        card.line.width = Pt(0.6)
        badge = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(x + 0.12), Inches(y + 0.12),
            Inches(0.34), Inches(0.34),
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = MARSH_RED
        badge.line.fill.background()
        _add_workflow_text(
            slide, f"{index + 1:02d}", x + 0.12, y + 0.12, 0.34, 0.34,
            9.5, RGBColor(0xFF, 0xFF, 0xFF), True,
            align=PP_ALIGN.CENTER,
        )
        _add_workflow_text(
            slide, _shorten(risk["title"], 48),
            x + 0.58, y + 0.07, row_width - 0.72, 0.24,
            9.5, INK, True,
        )
        if risk["description"]:
            _add_workflow_text(
                slide, _shorten(risk["description"], 92),
                x + 0.58, y + 0.32, row_width - 0.72, row_height - 0.36,
                8.3, MUTED,
            )

    if len(risks) > len(visible_risks):
        _add_workflow_text(
            slide, f"+ {len(risks) - len(visible_risks)} additional risks in the profile",
            0.68, 6.60, 11.90, 0.14, 7.5, MUTED,
        )


def _redesign_recommendation_slide(slide, overview, recommendation):
    """Render recommendation panels exclusively from profile/evidence data."""
    for shape in list(slide.shapes):
        if Inches(1.6) <= shape.top < Inches(7.0):
            shape._element.getparent().remove(shape._element)

    risks = []
    for item in _profile_items(overview.get("key_risks", [])):
        if isinstance(item, dict):
            title = str(
                item.get("title") or item.get("risk") or item.get("name") or ""
            ).strip()
            description = str(
                item.get("description") or item.get("short_description") or ""
            ).strip()
            if not title:
                title = str(item.get("detail") or item.get("reason") or "").strip()
        else:
            title = str(item).strip()
            description = ""
        if title:
            risks.append({"title": title, "description": description})
    benefits = [
        item for item in recommendation.get("supporting_benefits", []) or []
        if isinstance(item, dict) and str(item.get("benefit", "")).strip()
    ]
    policy = str(recommendation.get("recommended_policy", "")).strip()
    confidence = str(recommendation.get("confidence", "")).strip()

    _add_workflow_text(
        slide,
        "Client risks connected to verified policy evidence.",
        0.5, 1.68, 12.33, 0.27, 11.5, MUTED,
    )

    panel_y = 2.04
    panel_h = 3.66
    left_x, left_w = 0.42, 4.10
    middle_x, middle_w = 4.68, 3.20
    right_x, right_w = 8.04, 4.87
    for x, width, fill in (
        (left_x, left_w, RGBColor(0xFF, 0xF7, 0xF8)),
        (middle_x, middle_w, RGBColor(0xFF, 0xFF, 0xFF)),
        (right_x, right_w, RGBColor(0xF5, 0xFA, 0xFE)),
    ):
        panel = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(panel_y), Inches(width), Inches(panel_h),
        )
        panel.fill.solid()
        panel.fill.fore_color.rgb = fill
        panel.line.color.rgb = LINE
        panel.line.width = Pt(0.8)

    _add_workflow_text(
        slide, "CLIENT RISKS", left_x + 0.18, panel_y + 0.15,
        left_w - 0.36, 0.22, 10, MARSH_RED, True,
    )
    _add_workflow_text(
        slide, "Key areas of concern from the company profile.",
        left_x + 0.18, panel_y + 0.43, left_w - 0.36,
        0.30, 9, MUTED,
    )
    if risks:
        shown_risks = risks[:4]
        item_h = 0.62 if len(shown_risks) == 4 else 0.78 if len(shown_risks) == 3 else 0.95
        for index, risk in enumerate(shown_risks):
            y = panel_y + 0.88 + index * item_h
            badge = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(left_x + 0.20), Inches(y + 0.04),
                Inches(0.30), Inches(0.30),
            )
            badge.fill.solid()
            badge.fill.fore_color.rgb = PALE_RED
            badge.line.fill.background()
            _add_workflow_text(
                slide, f"{index + 1:02d}", left_x + 0.20, y + 0.04,
                0.30, 0.30, 8.5, MARSH_RED, True,
                align=PP_ALIGN.CENTER,
            )
            _add_workflow_text(
                slide, _shorten(risk["title"], 64),
                left_x + 0.62, y, left_w - 0.84, 0.27,
                9.5, INK, True,
            )
            if risk["description"]:
                _add_workflow_text(
                    slide, _shorten(risk["description"], 76),
                    left_x + 0.62, y + 0.27, left_w - 0.84,
                    max(item_h - 0.30, 0.25), 8.5, MUTED,
                )
        if len(risks) > len(shown_risks):
            _add_workflow_text(
                slide, f"+ {len(risks) - len(shown_risks)} additional profile risks",
                left_x + 0.20, panel_y + panel_h - 0.35,
                left_w - 0.40, 0.20, 8.5, MUTED,
            )
    else:
        _add_workflow_text(
            slide, "No client risks were supplied in the profile.",
            left_x + 0.20, panel_y + 1.00, left_w - 0.40,
            0.55, 10, MUTED,
        )

    _add_workflow_text(
        slide, "RECOMMENDED POLICY", middle_x + 0.18, panel_y + 0.15,
        middle_w - 0.36, 0.22, 10, MARSH_RED, True,
    )
    _add_workflow_text(
        slide, "Selected from supported benefits.",
        middle_x + 0.18, panel_y + 0.43, middle_w - 0.36,
        0.30, 9, MUTED,
    )
    policy_card_y = panel_y + 1.05
    policy_card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(middle_x + 0.16), Inches(policy_card_y),
        Inches(middle_w - 0.32), Inches(1.02),
    )
    policy_card.fill.solid()
    policy_card.fill.fore_color.rgb = PALE_GRAY
    policy_card.line.color.rgb = LINE
    policy_card.line.width = Pt(0.6)
    policy_text = _display_document_name(policy)
    policy_text = policy_text or "No sufficiently supported policy recommendation"
    policy_font = 12 if len(policy_text) <= 30 else 10 if len(policy_text) <= 54 else 8.5
    _add_workflow_text(
        slide, _shorten(policy_text, 100), middle_x + 0.32,
        policy_card_y + 0.13, middle_w - 0.64, 0.72,
        policy_font, INK, True, font="Cambria", align=PP_ALIGN.CENTER,
    )

    confidence_key = confidence.lower()
    if confidence_key in {"high", "medium", "low"}:
        confidence_color = (
            RGBColor(0x13, 0x8A, 0x55) if confidence_key == "high"
            else RGBColor(0xB5, 0x6A, 0x00) if confidence_key == "medium"
            else RGBColor(0xB4, 0x1F, 0x2B)
        )
        confidence_fill = (
            RGBColor(0xE8, 0xF5, 0xED) if confidence_key == "high"
            else RGBColor(0xFF, 0xF3, 0xDC) if confidence_key == "medium"
            else RGBColor(0xFB, 0xE8, 0xEA)
        )
        confidence_text = f"Confidence: {confidence.title()}"
    else:
        confidence_color = MUTED
        confidence_fill = PALE_GRAY
        confidence_text = f"Confidence: {confidence.title()}" if confidence else "Confidence unavailable"
    confidence_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(middle_x + 0.16), Inches(policy_card_y + 1.18),
        Inches(middle_w - 0.32), Inches(0.77),
    )
    confidence_box.fill.solid()
    confidence_box.fill.fore_color.rgb = confidence_fill
    confidence_box.line.fill.background()
    _add_workflow_text(
        slide, confidence_text, middle_x + 0.30,
        policy_card_y + 1.28, middle_w - 0.60, 0.23,
        10, confidence_color, True,
    )
    alignment_text = (
        f"{len(benefits)} verified benefit(s) support this option."
        if benefits else "No verified benefits support a policy selection."
    )
    _add_workflow_text(
        slide, alignment_text, middle_x + 0.30,
        policy_card_y + 1.54, middle_w - 0.60, 0.34,
        9, INK,
    )

    _add_workflow_text(
        slide, "KEY MATCHED BENEFITS", right_x + 0.18, panel_y + 0.15,
        right_w - 0.36, 0.22, 10, MARSH_RED, True,
    )
    _add_workflow_text(
        slide, "Verified benefits addressing the listed risks.",
        right_x + 0.18, panel_y + 0.43, right_w - 0.36,
        0.30, 9, MUTED,
    )
    risk_descriptions = {}
    for item in _profile_items(overview.get("key_risks", [])):
        if not isinstance(item, dict):
            continue
        title = str(item.get("title") or item.get("risk") or item.get("name") or "").strip()
        description = str(item.get("description") or item.get("short_description") or "").strip()
        if title and description:
            risk_descriptions[title.casefold()] = description

    benefit_groups = {}
    for benefit in benefits:
        risk = str(benefit.get("risk", "")).strip()
        address = str(
            benefit.get("risk_description")
            or risk_descriptions.get(risk.casefold())
            or risk
        ).strip()
        group_key = address.casefold() or f"benefit-{len(benefit_groups)}"
        group = benefit_groups.setdefault(group_key, {"address": address, "benefits": []})
        benefit_text = str(benefit.get("benefit", "")).strip()
        benefit_row = {
            "text": benefit_text,
            "page": benefit.get("page"),
        }
        if benefit_text and benefit_row not in group["benefits"]:
            group["benefits"].append(benefit_row)

    if benefit_groups:
        max_benefits = 4
        rendered_benefits = 0
        y = panel_y + 0.84
        for group in benefit_groups.values():
            if rendered_benefits >= max_benefits:
                break
            _add_workflow_text(
                slide, f"Addresses: {_shorten(group['address'], 76)}",
                right_x + 0.20, y, right_w - 0.40, 0.20,
                8.4, MUTED, True,
            )
            y += 0.21
            for benefit in group["benefits"]:
                if rendered_benefits >= max_benefits:
                    break
                _add_workflow_text(
                    slide, _shorten(benefit["text"], 110),
                    right_x + 0.20, y, right_w - 1.12, 0.36,
                    8.8, INK, True,
                )
                if benefit["page"] not in (None, ""):
                    _add_workflow_text(
                        slide, f"p. {benefit['page']}",
                        right_x + right_w - 0.90, y + 0.02,
                        0.68, 0.24, 8.2, MARSH_RED, True,
                        align=PP_ALIGN.RIGHT,
                    )
                rendered_benefits += 1
                y += 0.37
            y += 0.04
        if len(benefits) > rendered_benefits:
            _add_workflow_text(
                slide, f"+ {len(benefits) - rendered_benefits} more verified benefit(s)",
                right_x + 0.20, panel_y + panel_h - 0.20,
                right_w - 0.40, 0.15, 7.8, MUTED,
            )
    else:
        _add_workflow_text(
            slide, "No verified policy benefits were available for this recommendation.",
            right_x + 0.20, panel_y + 1.00, right_w - 0.40,
            0.60, 10, MUTED,
        )

    source_groups = {}
    for benefit in benefits:
        source = str(benefit.get("source", "")).strip()
        if not source:
            continue
        source_groups.setdefault(_display_document_name(source), None)

    source_panel = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.42), Inches(5.88), Inches(12.49), Inches(0.92),
    )
    source_panel.fill.solid()
    source_panel.fill.fore_color.rgb = RGBColor(0xFF, 0xF8, 0xF9)
    source_panel.line.color.rgb = RGBColor(0xF4, 0xD9, 0xDE)
    source_panel.line.width = Pt(0.8)
    _add_workflow_text(
        slide, "POLICY SOURCES", 0.62, 6.02,
        2.22, 0.21, 9.5, MARSH_RED, True,
    )
    if source_groups:
        shown_sources = list(source_groups)[:3]
        source_w = 3.10
        for index, source in enumerate(shown_sources):
            x = 2.88 + index * 3.25
            _add_workflow_text(
                slide, _shorten(source, 48), x, 6.08, source_w, 0.30,
                8.8, INK, True,
            )
        if len(source_groups) > len(shown_sources):
            _add_workflow_text(
                slide, f"+ {len(source_groups) - len(shown_sources)} more source(s)",
                11.18, 6.49, 1.90, 0.17, 7.5, MUTED,
                align=PP_ALIGN.RIGHT,
            )
    else:
        _add_workflow_text(
            slide, "No source citations are available from verified benefits.",
            2.88, 6.07, 8.9, 0.40, 9.5, MUTED,
        )


def _redesign_audit_slide(slide, audit_data):
    """Render audit metrics, claim evidence, and cited source documents."""
    for shape in list(slide.shapes):
        if Inches(1.6) <= shape.top < Inches(7.0):
            shape._element.getparent().remove(shape._element)

    claims = [
        item for item in audit_data.get("claim_results", []) or []
        if isinstance(item, dict)
    ]
    counts = {
        "SUPPORTED": 0,
        "PARTIALLY_SUPPORTED": 0,
        "UNSUPPORTED": 0,
        "NOT_FOUND": 0,
    }
    for claim in claims:
        status = str(claim.get("status", "")).strip().upper().replace(" ", "_")
        if status in counts:
            counts[status] += 1
    total = len(claims)
    pass_rate = round(counts["SUPPORTED"] / total * 100, 1) if total else 0
    requires_review = any(
        counts[key] for key in ("PARTIALLY_SUPPORTED", "UNSUPPORTED", "NOT_FOUND")
    )
    summary_title = "REVIEW REQUIRED" if requires_review else "AUDIT PASSED" if total else "NO CLAIMS TO AUDIT"
    summary_color = RGBColor(0xB4, 0x1F, 0x2B) if requires_review else RGBColor(0x13, 0x8A, 0x55)
    summary_fill = RGBColor(0xFB, 0xE8, 0xEA) if requires_review else RGBColor(0xE8, 0xF5, 0xED)

    _add_workflow_text(
        slide,
        "Claim-level evidence, status, and source traceability.",
        0.5, 1.68, 12.33, 0.27, 11.5, MUTED,
    )

    panel_y, panel_h = 2.02, 4.82
    summary_x, summary_w = 0.42, 2.75
    claims_x, claims_w = 3.34, 6.32
    sources_x, sources_w = 9.83, 3.08
    panel_specs = (
        (summary_x, summary_w, RGBColor(0xFF, 0xF7, 0xF8)),
        (claims_x, claims_w, RGBColor(0xFF, 0xFF, 0xFF)),
        (sources_x, sources_w, RGBColor(0xF5, 0xFA, 0xFE)),
    )
    for x, width, fill in panel_specs:
        panel = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(panel_y), Inches(width), Inches(panel_h),
        )
        panel.fill.solid()
        panel.fill.fore_color.rgb = fill
        panel.line.color.rgb = LINE
        panel.line.width = Pt(0.8)

    _add_workflow_text(
        slide, "AUDIT SUMMARY", summary_x + 0.17, panel_y + 0.15,
        summary_w - 0.34, 0.22, 10, MARSH_RED, True,
    )
    _add_workflow_text(
        slide, f"{counts['SUPPORTED']} / {total}", summary_x + 0.17,
        panel_y + 0.58, summary_w - 0.34, 0.60, 29, INK, True,
        font="Cambria",
    )
    _add_workflow_text(
        slide, "claims supported", summary_x + 0.17,
        panel_y + 1.16, summary_w - 0.34, 0.22, 10, MUTED,
    )
    _add_workflow_text(
        slide, f"{pass_rate:g}% supported", summary_x + 0.17,
        panel_y + 1.58, summary_w - 0.34, 0.22, 9.5, INK, True,
    )
    bar_x, bar_y, bar_w, bar_h = summary_x + 0.17, panel_y + 1.88, summary_w - 0.34, 0.12
    bar_bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(bar_x), Inches(bar_y), Inches(bar_w), Inches(bar_h),
    )
    bar_bg.fill.solid()
    bar_bg.fill.fore_color.rgb = RGBColor(0xE8, 0xEB, 0xEE)
    bar_bg.line.fill.background()
    if pass_rate:
        bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(bar_x), Inches(bar_y), Inches(bar_w * pass_rate / 100), Inches(bar_h),
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(0x13, 0x8A, 0x55)
        bar.line.fill.background()

    metric_items = [
        ("Supported", counts["SUPPORTED"], RGBColor(0xE8, 0xF5, 0xED), RGBColor(0x13, 0x8A, 0x55)),
        ("Partial", counts["PARTIALLY_SUPPORTED"], RGBColor(0xFF, 0xF3, 0xDC), RGBColor(0xB5, 0x6A, 0x00)),
        ("Unsupported", counts["UNSUPPORTED"], RGBColor(0xFB, 0xE8, 0xEA), RGBColor(0xB4, 0x1F, 0x2B)),
        ("Needs review", counts["NOT_FOUND"], RGBColor(0xE9, 0xF2, 0xFA), RGBColor(0x2C, 0x72, 0xA8)),
    ]
    tile_w, tile_h = 1.10, 0.70
    tile_gap_x, tile_gap_y = 0.20, 0.14
    tile_start_x, tile_start_y = summary_x + 0.16, panel_y + 2.20
    for index, (label, value, fill, accent) in enumerate(metric_items):
        col, row = index % 2, index // 2
        x = tile_start_x + col * (tile_w + tile_gap_x)
        y = tile_start_y + row * (tile_h + tile_gap_y)
        tile = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(y), Inches(tile_w), Inches(tile_h),
        )
        tile.fill.solid()
        tile.fill.fore_color.rgb = fill
        tile.line.fill.background()
        _add_workflow_text(
            slide, str(value), x + 0.10, y + 0.06,
            tile_w - 0.20, 0.30, 16, accent, True, font="Cambria",
        )
        _add_workflow_text(
            slide, label, x + 0.10, y + 0.38,
            tile_w - 0.20, 0.22, 8, INK,
        )

    status_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(summary_x + 0.16), Inches(panel_y + 4.00),
        Inches(summary_w - 0.32), Inches(0.55),
    )
    status_box.fill.solid()
    status_box.fill.fore_color.rgb = summary_fill
    status_box.line.fill.background()
    _add_workflow_text(
        slide, summary_title, summary_x + 0.30, panel_y + 4.08,
        summary_w - 0.60, 0.20, 9.3, summary_color, True,
    )
    supporting_text = (
        "All generated claims are supported by cited evidence."
        if total and not requires_review
        else "Review claims marked partial, unsupported, or not found."
        if total
        else "No claim results were returned for this run."
    )
    _add_workflow_text(
        slide, supporting_text, summary_x + 0.30, panel_y + 4.29,
        summary_w - 0.60, 0.24, 7.5, INK,
    )

    _add_workflow_text(
        slide, "CLAIM-LEVEL VERIFICATION", claims_x + 0.17,
        panel_y + 0.15, claims_w - 0.34, 0.22, 10, MARSH_RED, True,
    )
    _add_workflow_text(
        slide, "Claims and verification metadata from this audit.",
        claims_x + 0.17, panel_y + 0.43,
        claims_w - 0.34, 0.24, 8.5, MUTED,
    )
    table_x = claims_x + 0.14
    table_y = panel_y + 0.82
    table_w = claims_w - 0.28
    header = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(table_x), Inches(table_y), Inches(table_w), Inches(0.30),
    )
    header.fill.solid()
    header.fill.fore_color.rgb = PALE_GRAY
    header.line.fill.background()
    columns = (
        ("#", 0.28),
        ("GENERATED CLAIM", 2.10),
        ("STATUS", 0.95),
        ("CONF.", 0.64),
        ("SOURCE / PAGE", 1.84),
    )
    column_x = table_x + 0.08
    column_positions = []
    for label, width in columns:
        column_positions.append((column_x, width))
        _add_workflow_text(
            slide, label, column_x, table_y + 0.04, width, 0.21,
            7.2, INK, True,
        )
        column_x += width + 0.04

    max_rows = 8
    shown_claims = claims[:max_rows]
    row_start = table_y + 0.34
    row_area_h = 3.28
    row_h = row_area_h / max(len(shown_claims), 1)
    status_styles = {
        "SUPPORTED": ("Supported", RGBColor(0x13, 0x8A, 0x55)),
        "PARTIALLY_SUPPORTED": ("Partial", RGBColor(0xB5, 0x6A, 0x00)),
        "UNSUPPORTED": ("Unsupported", RGBColor(0xB4, 0x1F, 0x2B)),
        "NOT_FOUND": ("Not found", RGBColor(0x2C, 0x72, 0xA8)),
    }
    if shown_claims:
        for index, claim in enumerate(shown_claims):
            y = row_start + index * row_h
            if index % 2:
                row_fill = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, Inches(table_x), Inches(y),
                    Inches(table_w), Inches(row_h),
                )
                row_fill.fill.solid()
                row_fill.fill.fore_color.rgb = RGBColor(0xF8, 0xF9, 0xFA)
                row_fill.line.fill.background()
            raw_status = str(claim.get("status", "")).strip().upper().replace(" ", "_")
            status_text, status_color = status_styles.get(
                raw_status, ("Not recorded", MUTED)
            )
            confidence = str(claim.get("confidence", "")).strip()
            source = str(claim.get("source", "")).strip()
            source_name = _display_document_name(source) if source else "Not recorded"
            page = claim.get("page")
            source_page = f"{source_name}\nPage {page}" if page not in (None, "") else source_name
            values = (
                (str(index + 1), 8.0, INK, True),
                (_shorten(str(claim.get("claim", "")).strip() or "Claim text unavailable", 92), 7.8, INK, False),
                (status_text, 7.4, status_color, True),
                (confidence.title() if confidence else "—", 7.6, MUTED, True),
                (_shorten(source_page, 40), 7.5, INK, False),
            )
            for (x, width), (text, font_size, color, bold) in zip(column_positions, values):
                _add_workflow_text(
                    slide, text, x, y + 0.02, width, max(row_h - 0.04, 0.18),
                    font_size, color, bold,
                )
    else:
        _add_workflow_text(
            slide, "No claim results were returned for this audit.",
            table_x + 0.12, row_start + 0.30, table_w - 0.24,
            0.45, 10, MUTED,
        )
    if len(claims) > max_rows:
        _add_workflow_text(
            slide, f"{len(claims) - max_rows} additional claim(s) are in the audit report.",
            table_x + 0.08, panel_y + panel_h - 0.22,
            table_w - 0.16, 0.16, 7.2, MUTED,
        )

    _add_workflow_text(
        slide, "REFERENCE DOCUMENTS", sources_x + 0.17,
        panel_y + 0.15, sources_w - 0.34, 0.22, 10, MARSH_RED, True,
    )
    _add_workflow_text(
        slide, "Sources cited in claim verification.",
        sources_x + 0.17, panel_y + 0.43,
        sources_w - 0.34, 0.35, 8.5, MUTED,
    )
    references = {}
    for claim in claims:
        source = str(claim.get("source", "")).strip()
        if not source:
            continue
        source_name = _display_document_name(source)
        pages = references.setdefault(source_name, set())
        page = claim.get("page")
        if page not in (None, ""):
            pages.add(str(page))
    shown_sources = list(references.items())[:4]
    if shown_sources:
        source_row_h = 0.83
        for index, (source_name, pages) in enumerate(shown_sources):
            y = panel_y + 0.92 + index * 0.91
            source_card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(sources_x + 0.14), Inches(y),
                Inches(sources_w - 0.28), Inches(source_row_h),
            )
            source_card.fill.solid()
            source_card.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            source_card.line.color.rgb = LINE
            source_card.line.width = Pt(0.6)
            _add_workflow_text(
                slide, _shorten(source_name, 38), sources_x + 0.28,
                y + 0.12, sources_w - 0.56, 0.28,
                8.8, INK, True,
            )
            page_text = ", ".join(sorted(pages, key=lambda value: (not value.isdigit(), value)))
            citation = f"Pages {page_text}" if page_text else "Page not recorded"
            _add_workflow_text(
                slide, citation, sources_x + 0.28,
                y + 0.46, sources_w - 0.56, 0.20,
                8, MUTED,
            )
        if len(references) > len(shown_sources):
            _add_workflow_text(
                slide, f"+ {len(references) - len(shown_sources)} more document(s)",
                sources_x + 0.20, panel_y + panel_h - 0.24,
                sources_w - 0.40, 0.16, 7.5, MUTED,
            )
    else:
        _add_workflow_text(
            slide, "No source documents were recorded in this audit.",
            sources_x + 0.20, panel_y + 1.00,
            sources_w - 0.40, 0.55, 9.5, MUTED,
        )


def _add_review_stage(slide, x, title, subtitle, bullet_lines, icon):
    card_y = 2.48
    card_w = 4.0
    card_h = 1.84
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(card_y), Inches(card_w), Inches(card_h),
    )
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    card.line.color.rgb = LINE
    card.line.width = Pt(0.8)

    badge = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(x + 0.39), Inches(card_y - 0.06),
        Inches(0.38), Inches(0.38),
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = MARSH_RED
    badge.line.fill.background()
    _add_workflow_text(
        slide, f"{icon[1]:02d}", x + 0.39, card_y - 0.06, 0.38, 0.38,
        11, RGBColor(0xFF, 0xFF, 0xFF), True, align=PP_ALIGN.CENTER,
    )

    icon_x = x + 0.22
    icon_y = card_y + 0.34
    background = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(icon_x), Inches(icon_y),
        Inches(1.10), Inches(1.10),
    )
    background.fill.solid()
    background.fill.fore_color.rgb = PALE_RED
    background.line.fill.background()

    if icon[0] in {"output", "evidence"}:
        doc_x = icon_x + 0.34
        doc_y = icon_y + 0.20
        document = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(doc_x), Inches(doc_y),
            Inches(0.40), Inches(0.58),
        )
        document.fill.solid()
        document.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        document.line.color.rgb = INK
        document.line.width = Pt(2)
        for line_index in range(3):
            mark = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(doc_x + 0.08), Inches(doc_y + 0.16 + line_index * 0.12),
                Inches(0.23), Inches(0.025),
            )
            mark.fill.solid()
            mark.fill.fore_color.rgb = MARSH_RED
            mark.line.fill.background()
        if icon[0] == "evidence":
            lens = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(doc_x + 0.26), Inches(doc_y + 0.31),
                Inches(0.31), Inches(0.31),
            )
            lens.fill.solid()
            lens.fill.fore_color.rgb = PALE_RED
            lens.line.color.rgb = INK
            lens.line.width = Pt(2)
            handle = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(doc_x + 0.52), Inches(doc_y + 0.58),
                Inches(0.14), Inches(0.04),
            )
            handle.rotation = 45
            handle.fill.solid()
            handle.fill.fore_color.rgb = INK
            handle.line.fill.background()
        else:
            sparkle = slide.shapes.add_shape(
                MSO_SHAPE.STAR_5_POINT,
                Inches(doc_x + 0.34), Inches(doc_y - 0.04),
                Inches(0.16), Inches(0.16),
            )
            sparkle.fill.solid()
            sparkle.fill.fore_color.rgb = MARSH_RED
            sparkle.line.fill.background()
    else:
        head = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(icon_x + 0.43), Inches(icon_y + 0.20),
            Inches(0.24), Inches(0.24),
        )
        head.fill.solid()
        head.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        head.line.color.rgb = INK
        head.line.width = Pt(2)
        body = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(icon_x + 0.29), Inches(icon_y + 0.52),
            Inches(0.52), Inches(0.31),
        )
        body.fill.solid()
        body.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        body.line.color.rgb = INK
        body.line.width = Pt(2)
        check = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(icon_x + 0.68), Inches(icon_y + 0.57),
            Inches(0.25), Inches(0.25),
        )
        check.fill.solid()
        check.fill.fore_color.rgb = MARSH_RED
        check.line.fill.background()
        _add_workflow_text(
            slide, "✓", icon_x + 0.68, icon_y + 0.56, 0.25, 0.25,
            12, RGBColor(0xFF, 0xFF, 0xFF), True, align=PP_ALIGN.CENTER,
        )

    text_x = x + 1.48
    text_w = 2.30
    _add_workflow_text(
        slide, title, text_x, card_y + 0.25, text_w, 0.30,
        15, INK, True,
    )
    _add_workflow_text(
        slide, subtitle, text_x, card_y + 0.56, text_w, 0.25,
        10.5, MUTED,
    )
    for bullet_index, bullet in enumerate(bullet_lines):
        y = card_y + 0.91 + bullet_index * 0.37
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(text_x), Inches(y + 0.09),
            Inches(0.08), Inches(0.08),
        )
        dot.fill.solid()
        dot.fill.fore_color.rgb = MARSH_RED
        dot.line.fill.background()
        _add_workflow_text(
            slide, bullet, text_x + 0.20, y, text_w - 0.20, 0.34,
            9.5, INK,
        )


def _redesign_advisor_review_slide(slide, audit_data):
    """Build the advisor review slide from the current audit result."""
    for shape in list(slide.shapes):
        if Inches(1.6) <= shape.top < Inches(7.0):
            shape._element.getparent().remove(shape._element)

    _add_workflow_text(
        slide, "Human-in-the-loop control layer", 0.5, 1.70,
        12.33, 0.31, 19, INK, True, font="Cambria",
    )
    _add_workflow_text(
        slide,
        "AI assists the advisor, but final client-facing content remains subject to human review.",
        0.5, 2.02, 12.33, 0.25, 11.5, MUTED,
    )

    review_stages = [
        (0.37, "AI Output", "Generated claims", [
            "Claims are drafted from mapped policy benefits",
            "Source and page details travel with each claim",
        ], ("output", 1)),
        (4.67, "Evidence Check", "Source verification", [
            "Verify each claim against retrieved policy evidence",
            "Check the exact source and page reference",
        ], ("evidence", 2)),
        (8.97, "Advisor Review", "Human review", [
            "Review evidence and any audit concerns",
            "Approve, edit, or reject before client delivery",
        ], ("advisor", 3)),
    ]
    for stage in review_stages:
        _add_review_stage(slide, *stage)
    for arrow_x in (4.38, 8.68):
        _add_workflow_text(
            slide, "→", arrow_x, 3.13, 0.24, 0.38,
            20, MARSH_RED, True, align=PP_ALIGN.CENTER,
        )

    claims = audit_data.get("claim_results", []) or []
    usable_claims = [
        item for item in claims
        if isinstance(item, dict) and str(item.get("claim", "")).strip()
    ]
    status_priority = {
        "UNSUPPORTED": 0,
        "PARTIALLY_SUPPORTED": 1,
        "NOT_FOUND": 2,
        "SUPPORTED": 3,
    }
    if usable_claims:
        claim = min(
            usable_claims,
            key=lambda item: status_priority.get(
                str(item.get("status", "")).upper(), 2
            ),
        )
        claim_text = _shorten(claim.get("claim"), 105)
        source_text = _shorten(
            _display_document_name(claim.get("source")) or "Source not identified",
            42,
        )
        page_value = claim.get("page")
        page_text = f"p. {page_value}" if page_value not in (None, "") else "—"
        raw_status = str(claim.get("status", "Not assessed")).upper()
        status_text = raw_status.replace("_", " ").title()
        if raw_status == "SUPPORTED":
            status_color = RGBColor(0x13, 0x8A, 0x55)
            status_fill = RGBColor(0xE5, 0xF4, 0xEC)
        elif raw_status in {"PARTIALLY_SUPPORTED", "NOT_FOUND"}:
            status_color = RGBColor(0xB5, 0x6A, 0x00)
            status_fill = RGBColor(0xFF, 0xF2, 0xD8)
        else:
            status_color = RGBColor(0xB4, 0x1F, 0x2B)
            status_fill = RGBColor(0xFB, 0xE8, 0xEA)
    else:
        claim_text = "No audited claim is available for this run."
        source_text = "Not available"
        page_text = "—"
        status_text = "No audit data"
        status_color = MUTED
        status_fill = PALE_GRAY

    evidence = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.37), Inches(4.58), Inches(12.59), Inches(1.02),
    )
    evidence.fill.solid()
    evidence.fill.fore_color.rgb = PALE_GRAY
    evidence.line.color.rgb = LINE
    evidence.line.width = Pt(0.8)
    _add_workflow_text(
        slide, "EXAMPLE CLAIM (FROM AUDIT)", 0.55, 4.64,
        4.5, 0.20, 9.5, MARSH_RED, True,
    )
    _add_workflow_text(
        slide, f'“{claim_text}”', 0.92, 4.88,
        4.32, 0.57, 11, INK, True,
    )
    for x in (5.52, 8.06, 10.03):
        separator = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(x), Inches(4.78), Inches(0.01), Inches(0.62),
        )
        separator.fill.solid()
        separator.fill.fore_color.rgb = LINE
        separator.line.fill.background()
    _add_workflow_text(
        slide, "POLICY SOURCE", 5.75, 4.76, 2.15, 0.18,
        8.5, MUTED, True,
    )
    _add_workflow_text(
        slide, source_text, 5.75, 4.98, 2.15, 0.45,
        10, INK, True,
    )
    _add_workflow_text(
        slide, "PAGE", 8.28, 4.76, 1.40, 0.18,
        8.5, MUTED, True,
    )
    _add_workflow_text(
        slide, page_text, 8.28, 4.98, 1.45, 0.38,
        11, INK, True,
    )
    _add_workflow_text(
        slide, "AUDIT STATUS", 10.25, 4.76, 2.3, 0.18,
        8.5, MUTED, True,
    )
    status_badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(10.25), Inches(4.98), Inches(2.35), Inches(0.38),
    )
    status_badge.fill.solid()
    status_badge.fill.fore_color.rgb = status_fill
    status_badge.line.fill.background()
    _add_workflow_text(
        slide, status_text, 10.34, 4.99, 2.17, 0.36,
        10, status_color, True, align=PP_ALIGN.CENTER,
    )

    decision_panel = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.37), Inches(5.76), Inches(12.59), Inches(1.08),
    )
    decision_panel.fill.solid()
    decision_panel.fill.fore_color.rgb = RGBColor(0xFF, 0xFA, 0xFB)
    decision_panel.line.color.rgb = RGBColor(0xF4, 0xD9, 0xDE)
    decision_panel.line.width = Pt(0.8)
    _add_workflow_text(
        slide, "ADVISOR DECISION OPTIONS", 0.55, 5.78,
        4.5, 0.18, 9.5, MARSH_RED, True,
    )

    options = [
        ("APPROVE", "Claim is supported and appropriate.",
         RGBColor(0xE9, 0xF5, 0xEE), RGBColor(0x14, 0x8A, 0x55), "✓"),
        ("EDIT", "Claim needs clarification or modification.",
         RGBColor(0xFF, 0xF5, 0xE4), RGBColor(0xB5, 0x6A, 0x00), "✎"),
        ("REJECT", "Claim is not supported or appropriate.",
         RGBColor(0xFC, 0xEA, 0xEC), RGBColor(0xB4, 0x1F, 0x2B), "×"),
    ]
    option_w = 3.93
    option_gap = 0.20
    for index, (title, detail, fill_color, accent, symbol) in enumerate(options):
        x = 0.55 + index * (option_w + option_gap)
        option = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(6.00), Inches(option_w), Inches(0.70),
        )
        option.fill.solid()
        option.fill.fore_color.rgb = fill_color
        option.line.color.rgb = RGBColor(0xEA, 0xE5, 0xE6)
        option.line.width = Pt(0.6)
        icon = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(x + 0.15), Inches(6.10),
            Inches(0.48), Inches(0.48),
        )
        icon.fill.solid()
        icon.fill.fore_color.rgb = accent
        icon.line.fill.background()
        _add_workflow_text(
            slide, symbol, x + 0.15, 6.10, 0.48, 0.48,
            18, RGBColor(0xFF, 0xFF, 0xFF), True,
            align=PP_ALIGN.CENTER,
        )
        _add_workflow_text(
            slide, title, x + 0.78, 6.05, 2.95, 0.23,
            12, INK, True,
        )
        _add_workflow_text(
            slide, detail, x + 0.78, 6.29, 2.97, 0.31,
            9.5, INK,
        )


def _add_logo(slide):
    """Replace the template's text MARSH mark with the supplied logo image."""
    if not LOGO_PATH.exists():
        return

    # Remove the template MARSH wordmark only; keep the red decorative corner.
    targets = []
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip() == "MARSH":
            targets.append(shape)

    for shape in targets:
        left, top = shape.left, shape.top
        sp = shape._element
        sp.getparent().remove(sp)

        # Use the supplied logo without distortion. The template's original
        # MARSH text is replaced by the actual logo asset.
        slide.shapes.add_picture(
            str(LOGO_PATH),
            left,
            top - Inches(0.10),
            width=Inches(1.20),
            height=Inches(0.53),
        )

    # Keep the workspace label clear of the supplied logo.
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip() == "AI ADVISORY WORKSPACE":
            shape.top = Inches(0.72)


def _collect_slide_data(pitch):
    data = {item.get("title"): item.get("content", {}) for item in pitch}
    return data


def _format_status(value):
    """Make machine status values presentation-friendly."""
    if value is None:
        return ""
    return str(value).replace("_", " ").strip()


def _build_replacements(pitch):
    slides = _collect_slide_data(pitch)
    overview = slides.get("Company Overview", {})
    recommendation = slides.get("Evidence-Based Recommendation", {})
    audit = slides.get("AI Content Audit", {})

    replacements = {
        "{company_name}": overview.get("company", ""),
        "{industry}": overview.get("industry", ""),
        "{company_size}": overview.get("company_size", ""),
    }

    # Slide 1 risks.
    risks = overview.get("key_risks", []) or []
    for i in range(1, 6):
        risk = risks[i - 1] if i <= len(risks) else None
        if isinstance(risk, dict):
            title = risk.get("risk") or risk.get("title") or risk.get("name")
            detail = risk.get("detail") or risk.get("reason") or risk.get("description")
        else:
            title = risk
            detail = ""
        replacements[f"{{risk_{i}}}"] = _shorten(title, 100) if title else ""
        replacements[f"{{risk_{i}_detail}}"] = _shorten(detail, 155) if detail else ""

    # Slide 3 evidence table. Use up to three verified benefits because the
    # supplied template has three evidence rows.
    benefits = recommendation.get("supporting_benefits", []) or []
    for i in range(1, 4):
        item = benefits[i - 1] if i <= len(benefits) else {}
        replacements[f"{{evidence_{i}_risk}}"] = _shorten(item.get("risk", ""), 90)
        replacements[f"{{evidence_{i}_benefit}}"] = _shorten(item.get("benefit", ""), 125)
        source = item.get("source", "")
        page = item.get("page", "")
        source_text = f"{source}, p. {page}" if source and page else source
        replacements[f"{{evidence_{i}_source}}"] = _shorten(source_text, 80)

    replacements["{recommended_policy}"] = recommendation.get("recommended_policy", "")
    replacements["{confidence}"] = recommendation.get("confidence", "")
    replacements["{verified_benefits}"] = len(benefits)

    # Count unique mapped risks represented in the verified recommendation.
    mapped_risks = {b.get("risk") for b in benefits if b.get("risk")}
    replacements["{mapped_risks}"] = len(mapped_risks)

    # Slide 4 audit summary and per-claim traceability.
    replacements["{audit_total}"] = audit.get("total_claims", 0)
    replacements["{audit_supported}"] = audit.get("supported", 0)
    replacements["{audit_partial}"] = audit.get("partially_supported", 0)
    replacements["{audit_unsupported}"] = audit.get("unsupported", 0)
    replacements["{audit_status}"] = _format_status(audit.get("overall_status", "REVIEW_REQUIRED"))
    replacements["{audit_summary}"] = (
        f"Pass rate: {audit.get('pass_rate', 0)}% · "
        f"{audit.get('supported', 0)} supported · "
        f"{audit.get('partially_supported', 0)} partially supported · "
        f"{audit.get('unsupported', 0)} unsupported"
    )

    claims = audit.get("claim_results", []) or []
    for i in range(1, 6):
        result = claims[i - 1] if i <= len(claims) else None
        if result:
            replacements[f"{{claim_{i}_text}}"] = _shorten(result.get("claim", ""), 125)
            replacements[f"{{claim_{i}_status}}"] = _format_status(result.get("status", "NOT_FOUND"))
            replacements[f"{{claim_{i}_source}}"] = _shorten(result.get("source", ""), 55)
            replacements[f"{{claim_{i}_page}}"] = result.get("page", "—")
        else:
            replacements[f"{{claim_{i}_text}}"] = ""
            replacements[f"{{claim_{i}_status}}"] = ""
            replacements[f"{{claim_{i}_source}}"] = ""
            replacements[f"{{claim_{i}_page}}"] = ""

    return replacements, len(claims)


def create_pitch_ppt(pitch, output_path):
    """Populate the fixed Marsh template with generated pitch data."""
    template_path = Path(TEMPLATE_PATH)
    if not template_path.exists():
        raise FileNotFoundError(
            f"Marsh PPT template not found: {template_path}\n"
            "Place Marsh_AI_Pitch_Template.pptx inside templates/."
        )

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation(str(template_path))
    replacements, claim_count = _build_replacements(pitch)

    for slide_index, slide in enumerate(prs.slides, start=1):
        _add_logo(slide)

        if slide_index == 2:
            _redesign_workflow_slide(slide)
        elif slide_index == 5:
            _redesign_advisor_review_slide(
                slide,
                _collect_slide_data(pitch).get("AI Content Audit", {}),
            )
        elif slide_index == 3:
            slides_by_title = _collect_slide_data(pitch)
            _redesign_recommendation_slide(
                slide,
                slides_by_title.get("Company Overview", {}),
                slides_by_title.get("Evidence-Based Recommendation", {}),
            )
        elif slide_index == 4:
            _redesign_audit_slide(
                slide,
                _collect_slide_data(pitch).get("AI Content Audit", {}),
            )
        elif slide_index == 1:
            _redesign_company_overview_slide(
                slide,
                _collect_slide_data(pitch).get("Company Overview", {}),
            )
        elif slide_index == 4:
            for i in range(1, 6):
                if i > claim_count:
                    _hide_unused_claim_row(slide, i)

        # Tables first (Slide 3).
        for shape in slide.shapes:
            if getattr(shape, "has_table", False):
                _replace_table_cells(shape.table, replacements)

        # Normal text placeholders.
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                _replace_in_text_shape(shape, replacements)

    prs.save(str(output_path))
    print(f"PPT created from template: {output_path}")
    return str(output_path)


if __name__ == "__main__":
    print("Template-based ppt_generator.py loaded successfully.")
    print(f"Template: {TEMPLATE_PATH}")
    print(f"Logo: {LOGO_PATH}")
