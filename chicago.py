# chicago_dev.py
# Development-mode Chicago extractor.
# No OpenAI. No OpenAlex. No DOCX.
# Just: PDF  ->  footnotes  +  references  ->  Markdown text areas.

import re
import io
import fitz          # PyMuPDF
import streamlit as st
from datetime import datetime
from collections import Counter


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text(text):
    if not text:
        return ""
    text = text.replace("\u00ad", "")
    text = text.replace("‐", "-")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# ============================================================
# FOOTNOTE EXTRACTION — PyMuPDF geometry
# ============================================================

def extract_page_lines(page):
    data = page.get_text("dict")
    output = []

    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            pieces, sizes, span_data = [], [], []
            for span in spans:
                text = clean_text(span.get("text", ""))
                if not text:
                    continue
                size = float(span.get("size", 0))
                bbox = span.get("bbox", (0, 0, 0, 0))
                pieces.append(text)
                if size > 0:
                    sizes.append(size)
                span_data.append({
                    "text": text, "size": size,
                    "x0": float(bbox[0]), "y0": float(bbox[1]),
                    "x1": float(bbox[2]), "y1": float(bbox[3]),
                })
            if not pieces:
                continue
            text = clean_text(" ".join(pieces))
            x0, y0, x1, y1 = line["bbox"]
            output.append({
                "text": text,
                "x0": float(x0), "y0": float(y0),
                "x1": float(x1), "y1": float(y1),
                "min_size": min(sizes) if sizes else 0,
                "max_size": max(sizes) if sizes else 0,
                "avg_size": sum(sizes) / len(sizes) if sizes else 0,
                "spans": span_data,
            })
    output.sort(key=lambda line: (line["y0"], line["x0"]))
    return output


def merge_visual_lines(lines, y_tolerance=2.0):
    if not lines:
        return []
    sorted_lines = sorted(lines, key=lambda x: (x["y0"], x["x0"]))
    groups = []
    for line in sorted_lines:
        matched_group = None
        for group in reversed(groups[-4:]):
            if abs(line["y0"] - group["reference_y"]) <= y_tolerance:
                matched_group = group
                break
        if matched_group is None:
            groups.append({"reference_y": line["y0"], "lines": [line]})
        else:
            matched_group["lines"].append(line)
            normal_fragments = sorted(
                matched_group["lines"], key=lambda x: x["avg_size"], reverse=True,
            )
            matched_group["reference_y"] = normal_fragments[0]["y0"]

    merged = []
    for group in groups:
        fragments = sorted(group["lines"], key=lambda x: x["x0"])
        all_spans = []
        for fragment in fragments:
            all_spans.extend(fragment.get("spans", []))
        all_spans.sort(key=lambda s: s["x0"])
        text = clean_text(" ".join(
            span["text"] for span in all_spans if clean_text(span["text"])
        ))
        if not text:
            continue
        sizes = [span["size"] for span in all_spans if span["size"] > 0]
        merged.append({
            "text": text,
            "x0": min(f["x0"] for f in fragments),
            "y0": min(f["y0"] for f in fragments),
            "x1": max(f["x1"] for f in fragments),
            "y1": max(f["y1"] for f in fragments),
            "min_size": min(sizes) if sizes else 0,
            "max_size": max(sizes) if sizes else 0,
            "avg_size": sum(sizes) / len(sizes) if sizes else 0,
            "spans": all_spans,
        })
    merged.sort(key=lambda x: (x["y0"], x["x0"]))
    return merged


def find_bibliography_start(doc):
    headings = {"bibliography", "references", "daftar pustaka"}
    for page_index in range(len(doc)):
        page = doc[page_index]
        raw_lines = extract_page_lines(page)
        lines = merge_visual_lines(raw_lines, y_tolerance=2.0)
        for line in lines:
            if clean_text(line["text"]).lower() in headings:
                return page_index
    return len(doc)


def estimate_body_font_size(lines):
    sizes = []
    for line in lines:
        size = line["avg_size"]
        if size <= 0:
            continue
        sizes.append(round(size * 2) / 2)
    if not sizes:
        return None
    frequencies = {}
    for size in sizes:
        frequencies[size] = frequencies.get(size, 0) + 1
    return max(frequencies, key=frequencies.get)


def detect_footnote_number(line):
    spans = line.get("spans", [])
    if len(spans) < 2:
        return None
    first = spans[0]
    first_text = clean_text(first["text"])
    if not re.fullmatch(r"\d{1,3}[.]?", first_text):
        return None
    number_text = re.sub(r"\D", "", first_text)
    if not number_text:
        return None
    number = int(number_text)
    following_spans = [span for span in spans[1:] if clean_text(span["text"])]
    if not following_spans:
        return None
    second = following_spans[0]
    remainder = clean_text(" ".join(span["text"] for span in following_spans))
    if not remainder:
        return None
    marker_size = float(first["size"])
    text_size = float(second["size"])
    marker_y0 = float(first["y0"])
    text_y0 = float(second["y0"])
    smaller = marker_size < text_size - 0.3
    raised = marker_y0 < text_y0 - 0.5
    if not (smaller or raised):
        return None
    line_x0 = float(line["x0"])
    marker_x0 = float(first["x0"])
    if abs(marker_x0 - line_x0) > 3.0:
        return None
    return (number, remainder)


def is_running_footer(line, page_height):
    text = clean_text(line["text"])
    y1 = line["y1"]
    extreme_bottom = y1 > page_height * 0.965
    if extreme_bottom and re.fullmatch(r"\d{1,4}", text):
        return True
    if re.search(r"\b(?:ISSN|E-ISSN|P-ISSN)\b", text, re.I):
        return True
    if extreme_bottom and re.match(r"^(?:https?://|www\.)", text, re.I):
        return True
    return False


def is_journal_running_header(line):
    text = clean_text(line.get("text", ""))
    if not text:
        return False
    return bool(re.match(r"^\s*\d{1,4}\s+.+?\b(?:Vol\.?|Volume)\s*\d+", text, re.I))


def find_footnote_separator(page):
    page_width = page.rect.width
    separators = []
    try:
        drawings = page.get_drawings()
    except Exception:
        return None
    for drawing in drawings:
        for item in drawing.get("items", []):
            if not item or item[0] != "l":
                continue
            try:
                p1, p2 = item[1], item[2]
                x0, y0 = float(p1.x), float(p1.y)
                x1, y1 = float(p2.x), float(p2.y)
            except Exception:
                continue
            if abs(y1 - y0) > 2.0:
                continue
            width = abs(x1 - x0)
            if width < page_width * 0.20:
                continue
            y = (y0 + y1) / 2.0
            if y < page.rect.height * 0.15:
                continue
            if y > page.rect.height * 0.95:
                continue
            separators.append({"y": y, "width": width, "x0": min(x0, x1), "x1": max(x0, x1)})
    if not separators:
        return None
    separators.sort(key=lambda x: x["y"])
    return separators[-1]["y"]


def looks_like_footnote_continuation(line, body_size):
    text = clean_text(line.get("text", ""))
    if not text:
        return False
    small_font = (
        line.get("avg_size", 999) <= body_size - 0.25
        or line.get("min_size", 999) <= body_size - 0.25
    )
    if not small_font:
        return False
    if re.fullmatch(r"\d{1,4}", text):
        return False
    if re.search(r"\b(?:ISSN|E-ISSN|P-ISSN)\b", text, re.I):
        return False
    if is_journal_running_header(line):
        return False
    return True


def get_footnote_region_lines(page, lines, continuation_mode=False):
    page_height = page.rect.height
    if not lines:
        return []
    separator_y = find_footnote_separator(page)
    numbered_starts = [line for line in lines if detect_footnote_number(line) is not None]
    if not numbered_starts and not continuation_mode:
        return []
    if separator_y is not None:
        region_start_y = separator_y
    elif numbered_starts:
        first_start = min(numbered_starts, key=lambda x: x["y0"])
        region_start_y = first_start["y0"] - 2.0
    else:
        body_size = estimate_body_font_size(lines)
        if body_size is None:
            return []
        candidates = []
        for line in lines:
            if is_running_footer(line, page_height):
                continue
            if is_journal_running_header(line):
                continue
            if looks_like_footnote_continuation(line, body_size):
                candidates.append(line)
        candidates.sort(key=lambda line: (line["y0"], line["x0"]))
        return candidates

    candidates = []
    for line in lines:
        if line["y0"] <= region_start_y + 1.0:
            continue
        if is_running_footer(line, page_height):
            continue
        if not clean_text(line.get("text", "")):
            continue
        candidates.append(line)
    candidates.sort(key=lambda line: (line["y0"], line["x0"]))
    return candidates


def line_starts_numbered_note(line):
    if is_journal_running_header(line):
        return None
    strong = detect_footnote_number(line)
    if strong is not None:
        return strong
    text = clean_text(line.get("text", ""))
    m = re.match(r"^\s*(\d{1,3})[.)]?\s+(.+)$", text)
    if not m:
        return None
    return (int(m.group(1)), clean_text(m.group(2)))


def build_footnotes_from_document(page_candidates):
    notes = []
    current = None
    expected_number = None
    previous_page = None

    def save_current():
        nonlocal current
        if current is None:
            return
        current["text"] = clean_text(current["text"])
        notes.append(current)
        current = None

    for page_data in page_candidates:
        page_number = page_data["page"]
        footnote_lines = page_data["footnote_lines"]
        consumed_keys = set()

        if (current is not None and previous_page is not None
                and page_number == previous_page + 1):
            expected_line = None
            for line in footnote_lines:
                detected = line_starts_numbered_note(line)
                if detected is not None and detected[0] == expected_number:
                    expected_line = line
                    break
            if expected_line is not None:
                continuation_lines = []
                for line in footnote_lines:
                    if line["y0"] >= expected_line["y0"]:
                        continue
                    if line_starts_numbered_note(line) is not None:
                        continue
                    text = clean_text(line.get("text", ""))
                    if not text:
                        continue
                    continuation_lines.append(line)
                continuation_lines.sort(key=lambda x: (x["y0"], x["x0"]))
                for line in continuation_lines:
                    text = clean_text(line["text"])
                    consumed_keys.add((round(line["x0"], 2), round(line["y0"], 2), text))
                    current["text"] += " " + text
                    current["line_count"] += 1
                    current["end_page"] = page_number
                if continuation_lines:
                    current["cross_page"] = True
            else:
                continuation_lines = []
                for line in footnote_lines:
                    if line_starts_numbered_note(line) is not None:
                        continue
                    text = clean_text(line.get("text", ""))
                    if not text:
                        continue
                    continuation_lines.append(line)
                continuation_lines.sort(key=lambda x: (x["y0"], x["x0"]))
                for line in continuation_lines:
                    text = clean_text(line["text"])
                    consumed_keys.add((round(line["x0"], 2), round(line["y0"], 2), text))
                    current["text"] += " " + text
                    current["line_count"] += 1
                    current["end_page"] = page_number
                if continuation_lines:
                    current["cross_page"] = True

        for line in footnote_lines:
            text = clean_text(line.get("text", ""))
            key = (round(line["x0"], 2), round(line["y0"], 2), text)
            if key in consumed_keys:
                continue
            detected = line_starts_numbered_note(line)
            if current is None:
                if detected is None:
                    continue
                number, note_text = detected
                current = {
                    "number": number, "page": page_number, "end_page": page_number,
                    "y0": line["y0"], "text": note_text, "line_count": 1,
                    "cross_page": False, "layout_warning": False,
                    "layout_warning_message": "",
                }
                expected_number = number + 1
                continue
            if detected is not None:
                number, note_text = detected
                if number == expected_number:
                    save_current()
                    current = {
                        "number": number, "page": page_number, "end_page": page_number,
                        "y0": line["y0"], "text": note_text, "line_count": 1,
                        "cross_page": False, "layout_warning": False,
                        "layout_warning_message": "",
                    }
                    expected_number = number + 1
                    continue
            if text:
                current["text"] += " " + text
                current["line_count"] += 1
                current["end_page"] = page_number
        previous_page = page_number
    save_current()
    return notes


def clean_detected_footnotes(notes):
    cleaned = []
    for note in notes:
        text = clean_text(note["text"])
        if not text or len(text) < 5:
            continue
        note["text"] = text
        cleaned.append(note)
    return cleaned


def extract_chicago_footnotes(uploaded_file):
    uploaded_file.seek(0)
    pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    bibliography_page = find_bibliography_start(doc)
    page_candidates = []

    for page_index in range(bibliography_page):
        page = doc[page_index]
        page_number = page_index + 1
        raw_lines = extract_page_lines(page)
        lines = merge_visual_lines(raw_lines, y_tolerance=2.0)
        body_size = estimate_body_font_size(lines)
        candidates = get_footnote_region_lines(
            page, lines, continuation_mode=(page_index > 0),
        )
        page_candidates.append({
            "page": page_number,
            "footnote_lines": candidates,
            "all_lines": lines,
            "body_size": body_size,
        })

    all_notes = build_footnotes_from_document(page_candidates)
    doc.close()
    uploaded_file.seek(0)
    all_notes = clean_detected_footnotes(all_notes)
    return all_notes


# ============================================================
# BIBLIOGRAPHY EXTRACTION + SEGMENTATION
# ============================================================

def is_bibliography_heading(text):
    normalized = re.sub(r"\s+", " ", text).strip().lower()
    return normalized in {"bibliography", "references", "daftar pustaka"}


def extract_pdf_lines(page):
    data = page.get_text("dict")
    lines = []
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            text = "".join(span.get("text", "") for span in spans)
            text = re.sub(r"\s+", " ", text).strip()
            if not text:
                continue
            x0, y0, x1, y1 = line["bbox"]
            lines.append({
                "text": text,
                "x0": float(x0), "y0": float(y0),
                "x1": float(x1), "y1": float(y1),
            })
    return lines


def line_is_header_footer(line, page_height):
    text = re.sub(r"\s+", " ", line["text"]).strip()
    y0, y1 = line["y0"], line["y1"]
    in_header_zone = y0 < page_height * 0.12
    in_footer_zone = y1 > page_height * 0.93
    if re.fullmatch(r"\d{1,4}", text):
        if in_header_zone or in_footer_zone:
            return True
    if in_header_zone:
        if re.search(r"\bvol\.?\s*\d+", text, re.I):
            return True
        if re.search(r"\bvolume\s+\d+", text, re.I):
            return True
        if re.search(r"\bno\.?\s*\d+", text, re.I) and re.search(r"\b(19|20)\d{2}\b", text):
            return True
        if re.search(r"\s\d{2,4}\s*$", text):
            return True
        if "…" in text or "..." in text:
            return True
        if re.search(r"\b(journal|jurnal)\b", text, re.I) and re.search(r"\b(vol|volume|no|number)\b", text, re.I):
            return True
    if in_footer_zone:
        if re.match(r"^(https?://|www\.)", text, re.I):
            return True
        if re.search(r"\b(issn|e-issn|p-issn)\b", text, re.I):
            return True
    return False


def get_line_column(line, page_width):
    midpoint = page_width / 2
    center = (line["x0"] + line["x1"]) / 2
    return "LEFT" if center < midpoint else "RIGHT"


def extract_bibliography_lines(uploaded_file):
    uploaded_file.seek(0)
    pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    bibliography_found = False
    start_page = None
    start_column = None
    start_y = None
    ordered_lines = []

    for page_number, page in enumerate(doc, start=1):
        page_width = page.rect.width
        page_height = page.rect.height
        all_lines = extract_pdf_lines(page)

        heading = None
        if not bibliography_found:
            for line in all_lines:
                if is_bibliography_heading(line["text"]):
                    heading = line
                    break
            if heading is None:
                continue
            bibliography_found = True
            start_page = page_number
            start_column = get_line_column(heading, page_width)
            start_y = heading["y0"]

        lines = []
        for line in all_lines:
            if is_bibliography_heading(line["text"]):
                continue
            if line_is_header_footer(line, page_height):
                continue
            if line["y1"] > page_height * 0.94:
                continue
            lines.append(line)

        left, right = [], []
        for line in lines:
            column = get_line_column(line, page_width)
            line["page"] = page_number
            line["column"] = column
            if column == "LEFT":
                left.append(line)
            else:
                right.append(line)

        left.sort(key=lambda line: (line["y0"], line["x0"]))
        right.sort(key=lambda line: (line["y0"], line["x0"]))

        if page_number == start_page:
            if start_column == "RIGHT":
                right = [line for line in right if line["y0"] > start_y]
                ordered_lines.extend(right)
            else:
                left = [line for line in left if line["y0"] > start_y]
                ordered_lines.extend(left)
                ordered_lines.extend(right)
        else:
            ordered_lines.extend(left)
            ordered_lines.extend(right)

    doc.close()
    uploaded_file.seek(0)

    return {
        "found": bibliography_found,
        "start_page": start_page,
        "start_column": start_column,
        "lines": ordered_lines,
    }


_URL_OR_DOI_RE = re.compile(r"^(?:https?://|www\.|10\.\d{4,9}/|doi\s*:)", re.I)

_AUTHOR_START_RE = re.compile(
    r"^[A-Z\u00c0-\u00d6\u00d8-\u00dd][A-Za-z\u00c0-\u00ff'\u2019\-]+,\s+[A-Z]"
)

_CORPORATE_START_RE = re.compile(r"^[A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z]+){1,6}\.")

_QUOTED_START_RE = re.compile(r"^[\u201c\"]")

_NUMBERED_START_RE = re.compile(r"^\d{1,3}\.\s+[A-Z]")


def _looks_like_reference_start(fragment: str) -> bool:
    s = fragment.lstrip()
    if not s:
        return False
    if _AUTHOR_START_RE.match(s):
        return True
    if _QUOTED_START_RE.match(s):
        return True
    if _NUMBERED_START_RE.match(s):
        return True
    m = _CORPORATE_START_RE.match(s)
    if m:
        first_two = " ".join(m.group(0).split()[:2]).rstrip(".")
        if re.match(
            r"^(?:WSEAS|Journal|Jurnal|Review|International|Proceedings|"
            r"Transactions|Bulletin|Studies|Research)\b",
            first_two, re.I,
        ):
            return False
        return True
    return False


def _joined_lookahead(lines, start_index, max_lines=3):
    parts = []
    for j in range(start_index, min(start_index + max_lines, len(lines))):
        parts.append(lines[j]["text"].strip())
    return " ".join(parts)


def _starts_new_reference_at(lines, index):
    joined = _joined_lookahead(lines, index, max_lines=3)
    return _looks_like_reference_start(joined)


def _repair_single_segment(text: str) -> str:
    value = clean_text(text or "")
    if not value:
        return ""

    numeric_doi_wrap = re.compile(
        r"(10\.\d{4,9}/[-._;()/:A-Za-z0-9]*\d)\s+(\d{1,6})(?=(?:[.,;)]|\s|$))",
        re.I,
    )
    previous = None
    while value != previous:
        previous = value
        value = numeric_doi_wrap.sub(r"\1\2", value)

    value = re.sub(
        r"(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+\.)\s+(\d+(?:\.\d+)*)(?=[\s.,;)]|$)",
        r"\1\2", value, flags=re.I,
    )

    value = re.sub(
        r"((?:https?://|www\.)(?![^\s]*doi\.org/)[^\s]*[/-])\s+"
        r"([a-z0-9][A-Za-z0-9._~:/?#\[\]@!$&'()*+,;=%-]*)",
        r"\1\2",
        value,
    )

    value = re.sub(r"(\w+)-\s+([a-z]\w*)", r"\1-\2", value)

    value = re.sub(
        r"https?://(?:dx\.)?doi\.org/\s*https?://(?:dx\.)?doi\.org/",
        "https://doi.org/",
        value, flags=re.I,
    )
    return clean_text(value)


def _repair_bibliography_pdf_breaks(text: str) -> str:
    value = clean_text(text or "")
    if not value:
        return ""
    boundary_re = re.compile(r"(?<=[.?!])\s+(?=[A-Z\u00c0-\u00d6\u00d8-\u00dd])")
    parts = boundary_re.split(value)
    repaired = [_repair_single_segment(p) for p in parts]
    return clean_text(" ".join(p for p in repaired if p))


def _probable_bibliography_start(text: str) -> bool:
    s = clean_text(text or "")
    if not s or _URL_OR_DOI_RE.match(s):
        return False
    if re.match(
        r"^[A-Z\u00c0-\u00d6\u00d8-\u00dd][A-Za-z\u00c0-\u00ff'\u2019\-]+,\s*"
        r"[A-Z\u00c0-\u00d6\u00d8-\u00dd]",
        s,
    ):
        return True
    if re.match(r"^[A-Z\u00c0-\u00d6\u00d8-\u00dd][A-Za-z\u00c0-\u00ff'\u2019\-]+,\s*$", s):
        return True
    if re.match(
        r"^[A-Z\u00c0-\u00d6\u00d8-\u00dd][A-Za-z\u00c0-\u00ff'\u2019\-]{2,}\.\s+"
        r"[A-Z\u00c0-\u00d6\u00d8-\u00dd]",
        s,
    ):
        return True
    return _looks_like_reference_start(s)


def _column_start_margins(lines, cluster_tolerance=3.0):
    by_col = {"LEFT": [], "RIGHT": []}
    for line in lines:
        by_col.setdefault(line.get("column", "LEFT"), []).append(
            float(line.get("x0", 0.0))
        )

    left_min = min(by_col["LEFT"], default=None)
    right_min = min(by_col["RIGHT"], default=None)

    single_column = (
        left_min is not None
        and right_min is not None
        and abs(left_min - right_min) <= 10.0
    )

    def clusters(values):
        groups = []
        for value in sorted(values):
            best = None
            for group in groups:
                center = sum(group) / len(group)
                if abs(value - center) <= cluster_tolerance:
                    best = group
                    break
            if best is None:
                groups.append([value])
            else:
                best.append(value)
        return [{"x": sum(g) / len(g), "count": len(g)} for g in groups]

    if single_column:
        all_values = by_col["LEFT"] + by_col["RIGHT"]
        cs = clusters(all_values)
        if not cs:
            return {"LEFT": None, "RIGHT": None, "_unified": True}
        repeated = [c for c in cs if c["count"] >= 2]
        unified = min(c["x"] for c in (repeated or cs))
        return {"LEFT": unified, "RIGHT": unified, "_unified": True}

    margins = {"_unified": False}
    for col, values in by_col.items():
        cs = clusters(values)
        if not cs:
            margins[col] = None
            continue
        repeated = [c for c in cs if c["count"] >= 2]
        margins[col] = min(c["x"] for c in (repeated or cs))
    return margins


def split_references_from_lines(lines):
    if not lines:
        return []

    has_geometry = all(
        isinstance(x, dict) and float(x.get("x0", 0.0)) > 0.0
        for x in lines
    )

    normalised = []
    for item in lines:
        if isinstance(item, dict):
            text = clean_text(item.get("text", ""))
            if text:
                normalised.append({
                    "text": text,
                    "x0": float(item.get("x0", 0.0)),
                    "y0": float(item.get("y0", 0.0)),
                    "page": item.get("page"),
                    "column": item.get("column", "LEFT"),
                })
        else:
            text = clean_text(str(item))
            if text:
                normalised.append({
                    "text": text, "x0": 0.0, "y0": 0.0,
                    "page": None, "column": "LEFT",
                })

    if not normalised:
        return []

    # ---------------- HEURISTIC MODE ----------------
    if not has_geometry:
        refs, current = [], []
        for idx, line in enumerate(normalised):
            text = line["text"]
            if not current:
                current = [text]
                continue
            prev_ends = bool(re.search(r"[.?!]\s*$", current[-1]))
            starts_new = (
                prev_ends
                and not _URL_OR_DOI_RE.match(text)
                and (
                    _starts_new_reference_at(normalised, idx)
                    or _probable_bibliography_start(text)
                )
            )
            if starts_new:
                refs.append(_repair_bibliography_pdf_breaks(" ".join(current)))
                current = [text]
            else:
                current.append(text)
        if current:
            refs.append(_repair_bibliography_pdf_breaks(" ".join(current)))
        return [clean_text(r) for r in refs if clean_text(r)]

    # ---------------- GEOMETRY MODE ----------------
    margins = _column_start_margins(normalised)
    single_column_page = bool(margins.get("_unified"))
    margin_tolerance = 7.0

    refs, current = [], []
    current_column = None

    for idx, line in enumerate(normalised):
        text = line["text"]
        col = line["column"]
        base = margins.get(col)

        at_start_margin = (
            base is not None
            and abs(line["x0"] - base) <= margin_tolerance
        )
        url_or_doi = bool(_URL_OR_DOI_RE.match(text))
        content_start = (
            _starts_new_reference_at(normalised, idx)
            or _probable_bibliography_start(text)
        )
        prev_ends = (
            bool(current)
            and bool(re.search(r"[.?!]\s*$", current[-1]))
        )
        column_changed = (
            not single_column_page
            and current
            and current_column is not None
            and col != current_column
        )

        starts_new = bool(
            current
            and at_start_margin
            and not url_or_doi
            and content_start
            and prev_ends
        )
        if column_changed and at_start_margin and not url_or_doi and content_start:
            starts_new = True

        if starts_new:
            refs.append(_repair_bibliography_pdf_breaks(" ".join(current)))
            current = [text]
        else:
            current.append(text)

        current_column = col

    if current:
        refs.append(_repair_bibliography_pdf_breaks(" ".join(current)))

    return [clean_text(r) for r in refs if clean_text(r)]


# ============================================================
# MARKDOWN RENDERING
# ============================================================

def footnotes_to_markdown(footnotes):
    """
    One footnote per numbered Markdown list item, single blank line
    between notes.
    """
    if not footnotes:
        return "_No footnotes detected._"

    lines = []
    for note in footnotes:
        number = int(note.get("number", 0))
        text = clean_text(note.get("text", ""))
        pages = note.get("pages") or []
        if pages:
            page_hint = f"  _(p. {', '.join(map(str, pages))})_"
        else:
            page_hint = ""
        lines.append(f"{number}. {text}{page_hint}")
    # Blank line between entries so Markdown renders them as a
    # loose list — visually separated.
    return "\n\n".join(lines)


def references_to_markdown(references):
    """
    Each reference on its own line.

    Between references we insert TWO newlines so the rendered Markdown
    shows each reference as its own paragraph. That satisfies the
    'new reference = double enter' requirement.
    """
    if not references:
        return "_No references detected._"

    blocks = []
    for i, ref in enumerate(references, start=1):
        text = clean_text(ref)
        blocks.append(f"**[{i}]**  {text}")
    return "\n\n".join(blocks)


def references_to_markdown_raw(references):
    """
    Same as above but without the numbering prefix — pure one-per-line
    with a blank line between each. Useful for copy-paste into another
    tool.
    """
    if not references:
        return ""
    return "\n\n".join(clean_text(r) for r in references if clean_text(r))


# ============================================================
# RENDER — Streamlit development UI
# ============================================================

def render():
    st.title("OmniCite Chicago — DEV MODE (no AI)")

    uploaded_files = st.file_uploader(
        "Upload manuscript PDF(s)",
        type=["pdf"],
        accept_multiple_files=True,
        key="chicago_dev_uploader",
    )

    if not uploaded_files:
        return

    # Session state key per filename.
    state_key = "chicago_dev_results"
    if state_key not in st.session_state:
        st.session_state[state_key] = {}

    if st.button("Extract footnotes + references", type="primary",
                 use_container_width=True, key="chicago_dev_run"):
        progress = st.progress(0, text="Starting...")
        for i, uf in enumerate(uploaded_files, start=1):
            progress.progress(
                int((i - 1) / len(uploaded_files) * 100),
                text=f"[{i}/{len(uploaded_files)}] {uf.name}",
            )

            try:
                uf.seek(0)
                footnotes = extract_chicago_footnotes(uf)

                uf.seek(0)
                bib = extract_bibliography_lines(uf)

                if bib.get("found"):
                    references = split_references_from_lines(bib["lines"])
                else:
                    references = []

                st.session_state[state_key][uf.name] = {
                    "footnotes": footnotes,
                    "references": references,
                    "ref_found": bool(bib.get("found")),
                    "ref_count": len(references),
                    "fn_count": len(footnotes),
                }
            except Exception as exc:
                st.session_state[state_key][uf.name] = {
                    "error": f"{type(exc).__name__}: {exc}",
                }

        progress.progress(100, text="Done.")
        progress.empty()
        st.rerun()

    results = st.session_state.get(state_key, {})
    if not results:
        return

    names = list(results.keys())
    selected = st.selectbox(
        "Select manuscript",
        names,
        key="chicago_dev_selected",
    )

    item = results.get(selected) or {}
    if item.get("error"):
        st.error(item["error"])
        return

    footnotes = item.get("footnotes", [])
    references = item.get("references", [])

    col1, col2 = st.columns(2)
    col1.metric("Footnotes extracted", item.get("fn_count", len(footnotes)))
    col2.metric("References extracted", item.get("ref_count", len(references)))

    if not item.get("ref_found"):
        st.warning("No bibliography heading was found in this PDF.")

    st.subheader("1. Footnotes — Markdown")
    st.caption(
        "One footnote per numbered list item. Blank line between notes."
    )
    footnotes_md = footnotes_to_markdown(footnotes)
    st.code(footnotes_md, language="markdown")

    st.subheader("2. References — Markdown")
    st.caption(
        "One reference per line. Blank line (double enter) between "
        "references so they render as separate paragraphs."
    )
    references_md = references_to_markdown(references)
    st.code(references_md, language="markdown")

    st.subheader("3. References — Plain (copy/paste friendly)")
    st.caption("Same content without numbering. Blank line between each.")
    st.code(references_to_markdown_raw(references), language="text")