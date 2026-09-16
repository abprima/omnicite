# apa.py
import re
from datetime import datetime
import io
import os
import json
import traceback
import difflib

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from collections import Counter
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from openalex_config import get_openalex_api_key


# =========================================================
# SECRETS
# =========================================================

def _get_openalex_api_key():
    return get_openalex_api_key()


# =========================================================
# VERIFICATION THRESHOLDS
# =========================================================

TITLE_STRONG_MATCH    = 0.78
TITLE_REVIEW_MATCH    = 0.55
AUTHOR_STRONG_MATCH   = 0.60
AUTHOR_REVIEW_MATCH   = 0.35
SEVERE_TITLE_MISMATCH  = 0.30
SEVERE_AUTHOR_MISMATCH = 0.25

DOI_RESOLVER_URL     = "https://doi.org"
DOI_RESOLVER_TIMEOUT = 10


# =========================================================
# SAFE REGEX HELPERS
# =========================================================

def _safe_compile(pattern, flags=0):
    try:
        return re.compile(pattern, flags)
    except re.error:
        return None


def _safe_search(pattern, string, flags=0):
    compiled = _safe_compile(pattern, flags)
    if compiled is None or string is None:
        return None
    try:
        return compiled.search(string)
    except Exception:
        return None


def _safe_findall(pattern, string, flags=0):
    compiled = _safe_compile(pattern, flags)
    if compiled is None or string is None:
        return []
    try:
        return compiled.findall(string)
    except Exception:
        return []


def _safe_sub(pattern, repl, string, flags=0):
    compiled = _safe_compile(pattern, flags)
    if compiled is None or string is None:
        return string if string is not None else ""
    try:
        return compiled.sub(repl, string)
    except Exception:
        return string


def _safe_fullmatch(pattern, string, flags=0):
    compiled = _safe_compile(pattern, flags)
    if compiled is None or string is None:
        return None
    try:
        return compiled.fullmatch(string)
    except Exception:
        return None


def _esc(value):
    if value is None:
        return ""
    return re.escape(str(value))


# =========================================================
# PDF EXTRACTION + HEADER / FOOTER REMOVAL
# =========================================================

def normalize_running_text(text):
    text = re.sub(r"\s+", " ", text).strip().lower()
    text = re.sub(r"\b\d+\b", "<num>", text)
    return text


def extract_pdf_text(uploaded_file):
    pdf_bytes = uploaded_file.read()
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    page_blocks = []

    for page_number, page in enumerate(document):
        page_height = page.rect.height
        page_width = page.rect.width
        blocks = []
        for block in page.get_text("blocks"):
            x0, y0, x1, y1, text, *_ = block
            text = text.strip()
            if not text:
                continue
            blocks.append({
                "text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                "page_height": page_height, "page_width": page_width
            })
        blocks.sort(key=lambda b: (round(b["y0"], 1), round(b["x0"], 1)))
        page_blocks.append({"page": page_number + 1, "blocks": blocks})

    candidate_counter = Counter()
    for page_data in page_blocks:
        seen_on_page = set()
        for block in page_data["blocks"]:
            ph = block["page_height"]
            is_top = block["y0"] <= ph * 0.12
            is_bottom = block["y1"] >= ph * 0.90
            if not (is_top or is_bottom):
                continue
            normalized = normalize_running_text(block["text"])
            if not normalized:
                continue
            if normalized not in seen_on_page:
                candidate_counter[normalized] += 1
                seen_on_page.add(normalized)

    total_pages = len(page_blocks)
    minimum_pages = max(2, int(total_pages * 0.30))
    running_text_patterns = {
        text for text, count in candidate_counter.items()
        if count >= minimum_pages
    }

    pages = []
    removed_running_text = []
    for page_data in page_blocks:
        clean_blocks = []
        for block in page_data["blocks"]:
            text = block["text"]
            normalized = normalize_running_text(text)
            ph = block["page_height"]
            is_top = block["y0"] <= ph * 0.12
            is_bottom = block["y1"] >= ph * 0.90

            if (is_top or is_bottom) and normalized in running_text_patterns:
                removed_running_text.append({
                    "Page": page_data["page"], "Text": text,
                    "Reason": "Repeated header/footer"
                })
                continue

            if _safe_fullmatch(r"\s*(?:page\s*)?\d+\s*", text, flags=re.IGNORECASE):
                if is_top or is_bottom:
                    removed_running_text.append({
                        "Page": page_data["page"], "Text": text,
                        "Reason": "Page number"
                    })
                    continue
            clean_blocks.append(text)

        page_text = "\n".join(clean_blocks)
        pages.append({"page": page_data["page"], "text": page_text})

    document.close()
    full_text = "\n".join(
        f"<<<PAGE_BREAK:{p['page']}>>>\n{p['text']}" for p in pages
    )
    return full_text, pages, removed_running_text


# =========================================================
# PDF FONT / ITALIC EXTRACTION
# =========================================================

def extract_pdf_style_spans(uploaded_file):
    uploaded_file.seek(0)
    pdf_bytes = uploaded_file.read()
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    spans = []
    for page_number, page in enumerate(document, start=1):
        page_dict = page.get_text("dict")
        for block in page_dict.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if not text.strip():
                        continue
                    font = span.get("font", "")
                    flags = span.get("flags", 0)
                    italic = bool(flags & 2) or bool(
                        re.search(r"italic|oblique", font, re.I)
                    )
                    spans.append({
                        "page": page_number, "text": text, "font": font,
                        "flags": flags, "italic": italic, "bbox": span.get("bbox")
                    })
    document.close()
    uploaded_file.seek(0)
    return spans


def normalize_style_text(text):
    text = (text or "").replace("\u00ad", "")
    text = text.replace("‐", "-").replace("‑", "-").replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def _normalized_page_char_map(spans, page_number):
    chars, styles, fonts = [], [], []
    last_space = True
    for span in spans:
        if span["page"] != page_number:
            continue
        raw = (span.get("text") or "").replace("\u00ad", "")
        raw = raw.replace("‐", "-").replace("‑", "-").replace("–", "-").replace("—", "-")
        for ch in raw:
            if ch.isspace():
                if not last_space and chars:
                    chars.append(" "); styles.append(False); fonts.append("")
                last_space = True
            else:
                chars.append(ch.lower())
                styles.append(bool(span.get("italic")))
                fonts.append(span.get("font", ""))
                last_space = False
        if chars and not last_space:
            chars.append(" "); styles.append(False); fonts.append(""); last_space = True
    text = "".join(chars).strip()
    if len(text) < len(chars):
        chars = list(text); styles = styles[:len(text)]; fonts = fonts[:len(text)]
    return text, styles, fonts


def _style_for_exact_range(page_text, styles, fonts, start, end):
    positions = [i for i in range(start, end) if i < len(page_text) and not page_text[i].isspace()]
    if not positions:
        return None
    italic_count = sum(1 for i in positions if styles[i])
    ratio = italic_count / len(positions)
    used_fonts = sorted({fonts[i] for i in positions if fonts[i]})
    return {"found": True, "italic": ratio >= 0.80, "italic_ratio": ratio, "fonts": used_fonts}


def find_phrase_style(spans, phrase):
    target = normalize_style_text(phrase)
    if not target or len(target) < 2:
        return None
    for page in sorted({s["page"] for s in spans}):
        page_text, styles, fonts = _normalized_page_char_map(spans, page)
        idx = page_text.find(target)
        if idx >= 0:
            result = _style_for_exact_range(page_text, styles, fonts, idx, idx + len(target))
            if result:
                result["page"] = page
                return result
    return {"found": False, "italic": None, "italic_ratio": None, "fonts": [], "page": None}


def find_component_style_after_anchor(spans, anchor, component, max_distance=120):
    anchor_n = normalize_style_text(anchor)
    component_n = normalize_style_text(component)
    if not anchor_n or not component_n:
        return {"found": False, "italic": None, "italic_ratio": None, "fonts": [], "page": None}
    for page in sorted({s["page"] for s in spans}):
        page_text, styles, fonts = _normalized_page_char_map(spans, page)
        a = page_text.find(anchor_n)
        if a < 0:
            continue
        search_start = a + len(anchor_n)
        search_end = min(len(page_text), search_start + max_distance)
        region = page_text[search_start:search_end]
        m = _safe_search(r"(?<!\w)" + _esc(component_n) + r"(?!\w)", region)
        if m:
            start = search_start + m.start()
            end = search_start + m.end()
            result = _style_for_exact_range(page_text, styles, fonts, start, end)
            if result:
                result["page"] = page
                return result
    return {"found": False, "italic": None, "italic_ratio": None, "fonts": [], "page": None}


def extract_journal_format_parts(reference):
    pattern = re.compile(
        r"\.\s*(?P<journal>[^.]+?),\s*(?P<volume>\d+)\s*(?:\((?P<issue>[^)]+)\))?\s*,\s*"
        r"(?P<pages>\d+(?:\s*[–-]\s*\d+)?|[Ee]\d+|[Aa]rticle\s+\w+)(?=\s*[\.,]|\s*https?://|$)",
        re.I
    )
    matches = list(pattern.finditer(reference))
    if not matches:
        return {"is_journal": False, "journal": None, "volume": None, "issue": None, "pages": None}
    m = matches[-1]
    return {
        "is_journal": True, "journal": m.group("journal").strip(),
        "volume": m.group("volume"), "issue": m.group("issue"), "pages": m.group("pages")
    }


def extract_book_format_parts(reference):
    ym = re.search(r"\((?:(?:19|20)\d{2}[a-z]?|n\.d\.)\)\.\s*", reference, re.I)
    if not ym:
        return {"is_book": False, "title": None, "publisher": None}
    tail = reference[ym.end():].strip()
    if re.search(r",\s*\d+\s*(?:\([^)]+\))?\s*,\s*\d+", tail):
        return {"is_book": False, "title": None, "publisher": None}
    tail = re.sub(r"\s+https?://\S+\s*$", "", tail).strip()
    sentences = [x.strip() for x in re.split(r"(?<=\.)\s+", tail) if x.strip()]
    if len(sentences) < 2:
        return {"is_book": False, "title": None, "publisher": None}
    return {"is_book": True, "title": sentences[0].rstrip("."), "publisher": sentences[-1].rstrip(".")}


def detect_apa_source_type(reference):
    if extract_journal_format_parts(reference)["is_journal"]:
        return "Journal Article"
    low = reference.lower()
    if re.search(r"\b(?:in)\s+[A-Z].+\(eds?\.\)", reference) or re.search(r"\(eds?\.\)", reference, re.I):
        return "Book Chapter"
    if re.search(r"\b(proceedings|conference|symposium)\b", low):
        return "Conference Proceeding"
    if re.search(r"\b(report|technical report|working paper)\b", low):
        return "Report"
    if extract_book_format_parts(reference)["is_book"] and not re.search(r"https?://", reference):
        return "Book"
    if re.search(r"https?://", reference) and "doi.org" not in low:
        return "Webpage / Online Document"
    return "Other"


def check_reference_italics(reference, style_spans, source_type=None):
    source_type = source_type or detect_apa_source_type(reference)
    results = {
        "journal_title_italic": None, "volume_italic": None,
        "issue_not_italic": None, "book_title_italic": None, "style_notes": []
    }

    if source_type == "Journal Article":
        parts = extract_journal_format_parts(reference)
        journal, volume, issue = parts["journal"], parts["volume"], parts["issue"]
        js = find_phrase_style(style_spans, journal)
        if js and js.get("found"):
            results["journal_title_italic"] = js["italic"]
            if not js["italic"]:
                results["style_notes"].append("Journal title should be italic in APA 7.")
        else:
            results["style_notes"].append("Journal title formatting could not be located reliably in PDF spans.")
        vs = find_component_style_after_anchor(style_spans, journal, volume)
        if vs.get("found"):
            results["volume_italic"] = vs["italic"]
            if not vs["italic"]:
                results["style_notes"].append("Journal volume should be italic in APA 7.")
        else:
            results["style_notes"].append("Volume formatting could not be located reliably in PDF spans.")
        if issue:
            ins = find_component_style_after_anchor(style_spans, journal, f"({issue})")
            if ins.get("found"):
                results["issue_not_italic"] = not ins["italic"]
                if ins["italic"]:
                    results["style_notes"].append("Issue number should not be italic in APA 7.")
            else:
                results["style_notes"].append("Issue-number formatting could not be located reliably in PDF spans.")
    elif source_type == "Book":
        parts = extract_book_format_parts(reference)
        title = parts.get("title")
        if title:
            ts = find_phrase_style(style_spans, title)
            if ts and ts.get("found"):
                results["book_title_italic"] = ts["italic"]
                if not ts["italic"]:
                    results["style_notes"].append("Book title should be italic in APA 7.")
            else:
                results["style_notes"].append("Book-title formatting could not be located reliably in PDF spans.")
    return results


# =========================================================
# TEXT CLEANING
# =========================================================

def clean_text(text):
    text = text.replace("\u00ad", "")
    text = text.replace("\u2013", "–").replace("\u2014", "—")
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)
    return text


def strip_markdown_markers(text):
    if not text:
        return text
    text = _safe_sub(r"\*\*\*(.+?)\*\*\*", r"\1", text, flags=re.S)
    text = _safe_sub(r"\*\*(.+?)\*\*", r"\1", text, flags=re.S)
    text = _safe_sub(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"\1", text, flags=re.S)
    text = _safe_sub(r"(?<!_)_(?!\s)(.+?)(?<!\s)_(?!_)", r"\1", text, flags=re.S)
    text = re.sub(r"[ \t]{2,}", " ", text).strip()
    return text


# =========================================================
# REFERENCE SECTION DETECTION
# =========================================================

REFERENCE_HEADINGS = [
    "references", "reference", "daftar pustaka", "daftar rujukan",
    "rujukan", "bibliografi", "bibliography", "referensi",
]

POST_REFERENCE_HEADINGS = {
    "acknowledgement", "acknowledgements", "acknowledgment", "acknowledgments",
    "author contribution", "author contributions", "authors contribution",
    "authors contributions", "authors' contribution", "authors' contributions",
    "authors’ contribution", "authors’ contributions", "contribution", "contributions",
    "author profile", "authors profile", "author profiles", "authors profiles",
    "profile", "profiles", "biography", "biographies", "author biography",
    "author biographies", "conflict of interest", "conflicts of interest",
    "conflict of interests", "competing interest", "competing interests",
    "declaration", "declarations", "declaration of interest", "declaration of interests",
    "declarations of interest", "funding", "funding information", "funding statement",
    "data availability", "data availability statement", "availability of data",
    "ethical approval", "ethics approval", "ethics statement", "ethical statement",
    "informed consent", "consent for publication", "disclosure", "disclosures",
    "appendix", "appendices", "supplementary material", "supplementary materials",
    "supplemental material", "supplemental materials", "notes", "author note", "author notes",
}


def normalize_heading(text):
    text = text.strip().lower()
    text = re.sub(r"^\s*(?:\d+(?:\.\d+)*)[\.\s:-]+", "", text)
    text = re.sub(r"^[\*\#•\-\s]+|[\*\#•\-\s]+$", "", text)
    text = re.sub(r"\s+", " ", text).rstrip(":").strip()
    return text


def is_post_reference_heading(line):
    normalized = normalize_heading(line)
    if not normalized:
        return False
    if normalized in POST_REFERENCE_HEADINGS:
        return True
    for heading in POST_REFERENCE_HEADINGS:
        if normalized.startswith(heading + " "):
            return True
    return False


def find_reference_section(text):
    lines = text.splitlines()
    start_index = None
    end_index = None
    heading_found = None

    for i, line in enumerate(lines):
        if normalize_heading(line) in REFERENCE_HEADINGS:
            start_index = i
            heading_found = line.strip()
            break

    if start_index is None:
        return None, None, text

    for i in range(start_index + 1, len(lines)):
        line = lines[i].strip()
        if not line:
            continue
        if is_post_reference_heading(line):
            end_index = i
            break

    body_text = "\n".join(lines[:start_index])
    if end_index is not None:
        reference_text = "\n".join(lines[start_index + 1:end_index])
    else:
        reference_text = "\n".join(lines[start_index + 1:])
    return reference_text, heading_found, body_text


# =========================================================
# APA REFERENCE SPLITTING
# =========================================================

APA_DATE_RE = re.compile(r"\((?:(?:19|20)\d{2}[a-z]?|n\.d\.)\)", re.I)


def contains_apa_date(text):
    return bool(APA_DATE_RE.search(text or ""))


def is_page_break(line):
    return bool(re.fullmatch(r"<<<PAGE_BREAK:\d+>>>", line.strip()))


def is_strong_reference_start(line: str) -> bool:
    """
    Detect a strong APA reference beginning.

    Examples:
    Masita, Biswan, M., & Puspita, E. (2018).
    Rodiahwati, D. (2022).
    Kementerian Kesehatan RI. (2022).
    WHO. (2022).
    BKKBN. (2021).

    A line is considered a strong start when an author/organization
    is followed by a parenthesized year.
    """
    if not line:
        return False

    s = re.sub(r"\s+", " ", line).strip()

    # Do not treat obvious DOI/URL continuation lines as reference starts
    if re.match(r"^(https?://|www\.|doi\s*:)", s, re.I):
        return False

    # ── NEW: Reject lines that are just initials + year, e.g.
    #    "E. (2024). ..."  or  "D. W. E. (2024). ..."
    #    These are ALWAYS continuations of the previous author list.
    if re.match(
        r"^(?:[A-Z]\.)(?:\s*[A-Z]\.)*\s*"
        r"\((?:(?:19|20)\d{2}[a-z]?|n\.?d\.?)\)",
        s,
    ):
        return False

    # APA year near the beginning of the reference.
    year_pattern = r"\((?:(?:19|20)\d{2}[a-z]?|n\.?d\.?)\)"

    m = re.search(year_pattern, s, re.I)
    if not m:
        return False

    before_year = s[:m.start()].strip()

    if len(before_year) < 2:
        return False

    if len(before_year) > 220:
        return False

    # Personal author: Surname, Initials.
    personal = re.match(
        r"^[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+"
        r"(?:\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)*"
        r"\s*,",
        before_year
    )

    if personal:
        return True

    # Corporate/institutional author — requires at least 2 chars
    corporate = re.match(
        r"^[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ0-9&'’()./\-\s]+[.\s]?$",
        before_year
    )

    if corporate:
        stripped = before_year.rstrip(". ").strip()
        if len(stripped) >= 2:
            return True

    return False


def _looks_like_reference_tail(text):
    """
    Return True if `text` ends like a completed APA reference entry.
    """
    t = (text or "").rstrip()
    if not t:
        return False

    if re.search(r"https?://(?:dx\.)?doi\.org/10\.\d{4,9}/\S+$", t, re.I):
        return True
    if re.search(r"https?://\S+$", t, re.I):
        return True
    if re.search(r"\b\d+\s*[–-]\s*\d+\.\s*$", t):
        return True
    if re.search(r"\b\d+\.\s*$", t):
        return True

    return False


def split_glued_reference_line(line):
    """
    Split a PDF-extracted physical line when two or more APA references
    have been collapsed onto the same line.
    """
    if not line:
        return []

    line = re.sub(r"\s+", " ", line).strip()
    if not line:
        return []

    boundary_re = re.compile(
        r"(?:"
        r"https?://(?:dx\.)?doi\.org/10\.\d{4,9}/\S+"
        r"|https?://\S+"
        r"|\.\s+\d+\s*[–-]\s*\d+\s*\."
        r"|\b\d+\s*[–-]\s*\d+\s*\."
        r")"
    )

    candidate_re = re.compile(
        r"(?=\s+"
        r"[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+"
        r"(?:\s+[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)*"
        r"\s*,"
        r".{0,180}?"
        r"\((?:(?:19|20)\d{2}[a-z]?|n\.?d\.?)\)"
        r")"
    )

    anchors = [m.end() for m in boundary_re.finditer(line)]
    if not anchors:
        return [line]

    candidates = [m.start() for m in candidate_re.finditer(line)]
    if not candidates:
        return [line]

    valid_cuts = []
    for cand in candidates:
        for anchor in anchors:
            gap = line[anchor:cand]
            if len(gap) <= 3 and gap.strip() == "":
                before_text = line[:cand].rstrip()
                candidate_text = line[cand:].strip()
                if (is_strong_reference_start(candidate_text)
                        and _looks_like_reference_tail(before_text)):
                    valid_cuts.append(cand)
                    break

    if not valid_cuts:
        return [line]

    pieces = []
    last = 0
    for pos in valid_cuts:
        before = line[last:pos].strip()
        if before:
            pieces.append(before)
        last = pos
    remainder = line[last:].strip()
    if remainder:
        pieces.append(remainder)

    return pieces or [line]


def split_references(reference_text):
    raw_lines = reference_text.splitlines()
    lines = []
    for raw_line in raw_lines:
        line = raw_line.strip()
        if not line:
            continue
        if is_page_break(line):
            lines.append(line)
            continue
        line = line.replace("\u00ad", "")
        line = re.sub(r"\s+", " ", line).strip()
        if line:
            for piece in split_glued_reference_line(line):
                if piece:
                    lines.append(piece)

    references, current = [], []
    current_has_year = False

    def save_current():
        nonlocal current, current_has_year
        if not current:
            return
        reference = re.sub(r"\s+", " ", " ".join(current)).strip()
        reference = re.sub(r"<<<PAGE_BREAK:\d+>>>", " ", reference).strip()
        if reference:
            references.append(reference)
        current = []
        current_has_year = False

    def starts_personal_author(line):
        return bool(re.match(
            r"^[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+\s*,\s*(?:[A-Z](?:\.-?[A-Z])?\.\s*)+",
            line,
        ))

    def starts_corporate_author(line):
        return bool(re.match(
            r"^[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ0-9&'’\-\s]+?\.\s*"
            r"\((?:(?:19|20)\d{2}[a-z]?|n\.d\.)\)",
            line, re.I,
        ))

    # ── NEW: A line that starts with initials + year is a continuation.
    #    This covers "E. (2024). ...", "D. W. E. (2024). ...", etc.
    _INITIALS_ONLY_RE = re.compile(
        r"^(?:[A-Z]\.)(?:\s*[A-Z]\.)*\s*"
        r"\((?:(?:19|20)\d{2}[a-z]?|n\.?d\.?)\)"
    )

    def is_initials_only_continuation(line):
        return bool(_INITIALS_ONLY_RE.match(line.strip()))

    def previous_ends_author_connector():
        """
        Return True if the previous appended line clearly indicates the
        current reference's author list is not finished.
        """
        if not current:
            return False
        previous = current[-1].strip()

        # 1. Ends with a comma, ampersand, or "and"
        if re.search(r"(?:,|&|\band)\s*$", previous):
            return True

        # 2. Ends with one or more initials like "D. W." or just "W."
        #    This catches author lists whose last initial wraps to the
        #    next line ("... & Dewi, D. W.\nE. (2024). ...").
        if re.search(r"(?:^|\s)[A-Z]\.(?:\s*[A-Z]\.)*\s*$", previous):
            if not contains_apa_date(" ".join(current)):
                return True

        # 3. Ends mid-sentence with an author-list fragment
        if previous and not re.search(r"[.!?:]$", previous):
            if re.search(r"[A-ZÀ-ÖØ-Ý][a-z]+\s*,\s*(?:[A-Z]\.\s*)+$", previous):
                return True

        return False

    def looks_reference_complete(text):
        if not contains_apa_date(text):
            return False
        tail = text.rstrip()
        if re.search(r"(?:&|,\s*&|,)\s*$", tail):
            return False
        score = 0
        if re.search(r"https?://(?:dx\.)?doi\.org/10\.\d{4,9}/\S+\.?$", text, re.I):
            score += 4
        elif re.search(r"https?://\S+\.?$", text, re.I):
            score += 3
        if re.search(
            r"\b\d+\s*\([^)]*\)\s*,\s*(?:\d+\s*[–-]\s*\d+|e\d+|article\s+\w+)\.?$",
            text, re.I,
        ):
            score += 3
        elif re.search(r"\b\d+\s*[–-]\s*\d+\.?$", text):
            score += 2
        if re.search(r"\b(?:press|publishing|publisher|university press)\.?$",
                     text, re.I):
            score += 2
        if re.search(r"[.!?]$", text):
            score += 1
        m = APA_DATE_RE.search(text)
        if m and len(text[m.end():].strip()) >= 45:
            score += 1
        return score >= 2

    for line in lines:

        line = line.strip()
        if not line:
            continue

        if is_page_break(line):
            continue

        # ── GUARD A: initials-only continuation line
        if current and is_initials_only_continuation(line):
            current.append(line)
            current_has_year = current_has_year or contains_apa_date(line)
            continue

        # ── GUARD B: previous line ends with a connector / initial
        if previous_ends_author_connector():
            current.append(line)
            current_has_year = current_has_year or contains_apa_date(line)
            continue

        # ── 1. STRONG boundary
        if current and is_strong_reference_start(line):
            save_current()
            current = [line]
            current_has_year = contains_apa_date(line)
            continue

        # First line of the entire section
        if not current:
            current = [line]
            current_has_year = contains_apa_date(line)
            continue

        # ── 2. WEAKER boundary
        if starts_personal_author(line) or starts_corporate_author(line):
            if looks_reference_complete(" ".join(current)):
                save_current()
                current = [line]
                current_has_year = contains_apa_date(line)
            else:
                current.append(line)
                current_has_year = current_has_year or contains_apa_date(line)
            continue

        # If the current reference has no year yet, keep appending
        if not current_has_year:
            current.append(line)
            current_has_year = current_has_year or contains_apa_date(line)
            continue

        # ── 3. Otherwise it is a continuation
        current.append(line)
        current_has_year = current_has_year or contains_apa_date(line)

    save_current()
    return references


# =========================================================
# BASIC APA REFERENCE PARSER
# =========================================================

def normalize_doi_or_url(value):
    if not value:
        return value

    text = str(value).strip()

    m = re.match(r'^https?://(?:dx\.)?doi\.org(https?://.+)$', text, re.I)
    if m:
        return m.group(1).rstrip(".,;)")

    text = re.sub(
        r'(?i)^https?://(?:dx\.)?doi\.org/'
        r'(?:https?://(?:dx\.)?doi\.org/)+',
        'https://doi.org/',
        text
    )

    doi_match = re.search(r'10\.\d{4,9}/[-._;()/:A-Za-z0-9]+', text, re.I)
    if doi_match:
        doi = doi_match.group(0).rstrip(".,;)")
        doi = re.sub(r'^(10\.\d{4,9})/{2,}', r'\1/', doi)
        return f"https://doi.org/{doi}"

    url_match = re.search(r'https?://\S+', text, re.I)
    if url_match:
        return url_match.group(0).rstrip(".,;)")

    return text


def parse_reference(reference):
    result = {
        "raw": reference,
        "year": None,
        "first_author": None,
        "authors": [],
        "doi": None,
        "url": None
    }

    year_match = re.search(r"\(((?:19|20)\d{2})[a-z]?\)", reference)
    if not year_match:
        year_match = re.search(r"\b((?:19|20)\d{2})\b", reference)
    if year_match:
        result["year"] = year_match.group(1)

    if year_match:
        author_block = reference[:year_match.start()].strip()
    else:
        author_block = reference[:250]

    author_matches = re.findall(
        r"(?:^|,\s*)&?\s*"
        r"([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)"
        r"\s*,\s*(?:[A-Z]\.\s*)+",
        author_block
    )

    if author_matches:
        result["authors"] = author_matches
        result["first_author"] = author_matches[0]
    else:
        corporate = author_block.rstrip(" .,")
        if corporate:
            result["authors"] = [corporate]
            result["first_author"] = corporate

    url_match = re.search(r"https?://\S+", reference, re.I)
    if url_match:
        original_url = url_match.group(0).rstrip(".,)")
        normalized_url = normalize_doi_or_url(original_url)
        if normalized_url and re.match(r"^https://doi\.org/10\.\d{4,9}/", normalized_url, re.I):
            result["doi"] = normalized_url
            result["url"] = normalized_url
        else:
            result["url"] = normalized_url
    else:
        doi_match = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", reference, re.I)
        if doi_match:
            normalized_url = normalize_doi_or_url(doi_match.group(0))
            result["doi"] = normalized_url
            result["url"] = normalized_url

    return result


def build_apa_disambiguation_map(parsed_references):
    groups = {}

    for ref in parsed_references:
        authors = ref.get("authors") or []
        year = ref.get("year")
        if len(authors) < 3 or not year:
            continue
        key = (normalize(authors[0]), str(year))
        groups.setdefault(key, []).append(authors)

    ambiguous = {}
    for key, author_lists in groups.items():
        unique = []
        for authors in author_lists:
            normalized_authors = tuple(normalize(a) for a in authors)
            if normalized_authors not in [
                tuple(normalize(a) for a in x)
                for x in unique
            ]:
                unique.append(authors)
        if len(unique) > 1:
            ambiguous[key] = unique

    return ambiguous


def get_disambiguated_author_form(authors, year, disambiguation_map):
    if not authors:
        return None
    if len(authors) < 3:
        return None

    key = (normalize(authors[0]), str(year))
    competing_lists = disambiguation_map.get(key, [])

    if len(competing_lists) <= 1:
        return f"{authors[0]} et al."

    target = [normalize(a) for a in authors]
    others = []
    for other in competing_lists:
        other_norm = [normalize(a) for a in other]
        if other_norm != target:
            others.append(other_norm)

    if not others:
        return f"{authors[0]} et al."

    for keep_count in range(2, len(authors) + 1):
        target_prefix = target[:keep_count]
        still_ambiguous = False
        for other in others:
            if other[:keep_count] == target_prefix:
                still_ambiguous = True
                break
        if not still_ambiguous:
            displayed = authors[:keep_count]
            if keep_count < len(authors):
                return ", ".join(displayed) + ", et al."
            return ", ".join(displayed)

    return ", ".join(authors)


def is_apa_disambiguated_citation(citation, disambiguation_map):
    raw = citation.get("raw", "")
    year = citation.get("year")
    first_author = citation.get("author")

    if not raw or not year or not first_author:
        return False

    key = (normalize(first_author), str(year))
    competing_lists = disambiguation_map.get(key, [])

    if len(competing_lists) <= 1:
        return False

    raw_normalized = normalize(raw)

    for authors in competing_lists:
        correct_form = get_disambiguated_author_form(authors, year, disambiguation_map)
        if not correct_form:
            continue
        expected = normalize(correct_form)
        if expected and expected in raw_normalized:
            return True

    return False


# =========================================================
# APA IN-TEXT CITATION EXTRACTION
# =========================================================

def extract_parenthetical_citations(text):
    citations = []
    year_pattern = re.compile(r"\b(?:19|20)\d{2}[a-z]?\b")
    for content in re.findall(r"\(([^()]+)\)", text):
        if not year_pattern.search(content):
            continue
        for part in content.split(";"):
            ym = year_pattern.search(part)
            if not ym:
                continue
            year = ym.group(0)[:4]
            author_part = part[:ym.start()].strip(" ,")
            et_al = bool(re.search(r"\bet\s+al\.", author_part, re.I))
            author_part_clean = re.sub(r"\bet\s+al\.", "", author_part, flags=re.I).strip()
            authors = re.findall(
                r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\b",
                author_part_clean
            )
            authors = [a for a in authors if a.lower() not in {"and", "according", "see", "cf"}]
            if not authors:
                continue
            citations.append({
                "author": authors[0],
                "authors": authors,
                "year": year,
                "type": "parenthetical",
                "et_al": et_al,
                "raw": f"({part.strip()})",
                "parenthetical_content": content.strip(),
            })
    return citations


def _inside_parentheses(text, pos):
    depth = 0
    for i in range(pos):
        ch = text[i]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
    return depth > 0


def extract_narrative_citations(text):
    citations = []
    occupied = []

    # Pattern A — "Author, Author et al. (2024)"
    for m in re.finditer(
        r"\b("
        r"[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+"
        r"(?:\s*,\s*[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)+"
        r")"
        r"\s*,?\s*et\s+al\.\s*"
        r"\(((?:19|20)\d{2})[a-z]?\)",
        text
    ):
        author_text = m.group(1)
        authors = [a.strip() for a in author_text.split(",") if a.strip()]
        if not authors:
            continue
        citations.append({
            "author": authors[0], "authors": authors,
            "year": m.group(2), "type": "narrative",
            "et_al": True, "raw": m.group(0),
        })
        occupied.append((m.start(), m.end()))

    # Pattern B — "Author et al. (2024)"
    for m in re.finditer(
        r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\s+et\s+al\.\s*"
        r"\(((?:19|20)\d{2})[a-z]?\)",
        text
    ):
        if any(m.start() >= s and m.end() <= e for s, e in occupied):
            continue
        citations.append({
            "author": m.group(1), "authors": [m.group(1)],
            "year": m.group(2), "type": "narrative",
            "et_al": True, "raw": m.group(0),
        })
        occupied.append((m.start(), m.end()))

    # Pattern C — "Author (2024)"
    for m in re.finditer(
        r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\s+"
        r"\(((?:19|20)\d{2})[a-z]?\)",
        text
    ):
        if any(m.start() >= s and m.end() <= e for s, e in occupied):
            continue
        citations.append({
            "author": m.group(1), "authors": [m.group(1)],
            "year": m.group(2), "type": "narrative",
            "et_al": False, "raw": m.group(0),
        })
        occupied.append((m.start(), m.end()))

    # Pattern D — "Author et al., 2024" (MALFORMED)
    for m in re.finditer(
        r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)"
        r"\s+et\s+al\.\s*,\s*"
        r"((?:19|20)\d{2})[a-z]?\b",
        text
    ):
        if any(m.start() >= s and m.end() <= e for s, e in occupied):
            continue
        if _inside_parentheses(text, m.start()):
            continue
        citations.append({
            "author": m.group(1), "authors": [m.group(1)],
            "year": m.group(2), "type": "narrative",
            "et_al": True, "raw": m.group(0),
            "malformed": True,
        })
        occupied.append((m.start(), m.end()))

    # Pattern E — "Author and Author, 2024"
    for m in re.finditer(
        r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)"
        r"\s+(?:and|&)\s+"
        r"([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\s*,\s*"
        r"((?:19|20)\d{2})[a-z]?\b",
        text
    ):
        if any(m.start() >= s and m.end() <= e for s, e in occupied):
            continue
        if _inside_parentheses(text, m.start()):
            continue
        citations.append({
            "author": m.group(1),
            "authors": [m.group(1), m.group(2)],
            "year": m.group(3), "type": "narrative",
            "et_al": False, "raw": m.group(0),
            "malformed": True,
        })
        occupied.append((m.start(), m.end()))

    # Pattern F — "Author, 2024"
    for m in re.finditer(
        r"(?<![,\.])\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\s*,\s*"
        r"((?:19|20)\d{2})[a-z]?\b",
        text
    ):
        if any(m.start() >= s and m.end() <= e for s, e in occupied):
            continue
        if _inside_parentheses(text, m.start()):
            continue
        if m.start() > 0 and text[m.start() - 1] == ",":
            continue
        citations.append({
            "author": m.group(1), "authors": [m.group(1)],
            "year": m.group(2), "type": "narrative",
            "et_al": False, "raw": m.group(0),
            "malformed": True,
        })
        occupied.append((m.start(), m.end()))

    return citations


def extract_all_citations(text):
    return extract_parenthetical_citations(text) + extract_narrative_citations(text)


# =========================================================
# STATISTICS / CHECKERS
# =========================================================

def calculate_citation_statistics(citations):
    narrative = sum(1 for c in citations if c.get("type") == "narrative")
    parenthetical = sum(1 for c in citations if c.get("type") == "parenthetical")
    total = narrative + parenthetical
    return {
        "total": total, "narrative": narrative, "parenthetical": parenthetical,
        "narrative_percentage": (narrative / total * 100 if total else 0),
        "parenthetical_percentage": (parenthetical / total * 100 if total else 0)
    }


def check_apa_intext_citation(citation):
    issues = []
    raw = citation.get("raw", "")
    ctype = citation.get("type")
    authors = citation.get("authors", [])
    et_al = citation.get("et_al", False)
    if ctype == "parenthetical":
        if len(authors) == 2 and not et_al and "&" not in raw:
            issues.append("Two-author parenthetical citation should use '&'.")
        if len(authors) >= 3 and not et_al:
            issues.append("APA 7 uses 'et al.' for in-text citations with three or more authors.")
    elif ctype == "narrative":
        if len(authors) == 2 and not et_al and not re.search(r"\band\b", raw, re.I):
            issues.append("Two-author narrative citation should use 'and', not '&'.")
        if len(authors) >= 3 and not et_al:
            issues.append("APA 7 uses 'et al.' for in-text citations with three or more authors.")
    return ("Pass" if not issues else "Needs review", issues)


def normalize(value):
    if not value:
        return ""
    return re.sub(r"[^a-z0-9]", "", value.lower().strip())


def match_citations_to_references(citations, parsed_references):
    results = []
    keys = {(normalize(c["author"]), c["year"]) for c in citations}
    for i, ref in enumerate(parsed_references, start=1):
        key = (normalize(ref["first_author"]), ref["year"])
        results.append({
            "Reference #": i, "Author": ref["first_author"], "Year": ref["year"],
            "Cited": key in keys, "Reference": ref["raw"]
        })
    return results


def find_missing_references(citations, parsed_references):
    ref_keys = {(normalize(r["first_author"]), r["year"]) for r in parsed_references}
    missing, seen = [], set()
    for c in citations:
        key = (normalize(c["author"]), c["year"])
        if key not in ref_keys and key not in seen:
            missing.append({"Citation": c["raw"], "Author": c["author"], "Year": c["year"]})
            seen.add(key)
    return missing


def detect_duplicates(parsed_references):
    duplicates = []
    for doi, count in Counter(normalize(r["doi"]) for r in parsed_references if r["doi"]).items():
        if count > 1:
            duplicates.append({"Type": "DOI", "Value": doi, "Count": count})
    return duplicates


def check_apa_reference(reference, style_spans=None):
    issues, warnings = [], []
    ref = reference.strip()
    source_type = detect_apa_source_type(ref)

    year_match = re.search(r"\((?:(?:19|20)\d{2}[a-z]?|n\.d\.)\)", ref, re.I)
    if not year_match:
        issues.append("Missing or incorrectly formatted APA publication date.")
        author_section = ""
    else:
        author_section = ref[:year_match.start()].strip()
        if not author_section:
            issues.append("Author or group author is missing.")
        if not re.match(r"\s*\.", ref[year_match.end():]):
            issues.append("Publication date should normally be followed by a period.")

    personal_authors = re.findall(
        r"([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\s*,\s*(?:[A-Z](?:\.-?[A-Z])?\.\s*)+",
        author_section
    )
    if len(personal_authors) >= 2 and "&" not in author_section:
        issues.append("Multiple-author reference should use '&' before the final author.")
    if re.search(r"\band\b", author_section, re.I):
        issues.append("APA reference lists use '&', not 'and', before the final author.")

    if re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", ref, re.I):
        if not re.search(r"https://doi\.org/10\.\d{4,9}/", ref, re.I):
            issues.append("DOI should use the https://doi.org/... format.")
        if re.search(r"\bdoi\s*:", ref, re.I):
            issues.append("Do not use the old 'doi:' prefix; use https://doi.org/...")

    if len([w for w in ref.split() if len(w) > 3 and w.isupper()]) >= 4:
        warnings.append("Possible incorrect capitalization; APA article/book titles normally use sentence case.")

    italic_results = {
        "journal_title_italic": None, "volume_italic": None,
        "issue_not_italic": None, "book_title_italic": None, "style_notes": []
    }
    if style_spans is not None:
        italic_results = check_reference_italics(ref, style_spans, source_type)
        if italic_results["journal_title_italic"] is False:
            issues.append("Journal title is not italic in the PDF.")
        if italic_results["volume_italic"] is False:
            issues.append("Journal volume is not italic in the PDF.")
        if italic_results["issue_not_italic"] is False:
            issues.append("Journal issue number appears italic; APA 7 normally keeps the issue number non-italic.")
        if italic_results.get("book_title_italic") is False:
            issues.append("Book title is not italic in the PDF.")
        for note in italic_results["style_notes"]:
            if "could not be located" in note or "not applied" in note:
                warnings.append(note)

    if issues:
        status = "Fail"
    elif warnings:
        status = "Review"
    else:
        status = "Pass"
    italic_results["source_type"] = source_type
    return status, issues, warnings, italic_results


def extract_reference_year(reference):
    m = re.search(r"\(((?:19|20)\d{2})(?:[a-z])?\)", reference, re.I)
    return int(m.group(1)) if m else None


def calculate_reference_recency(references, manuscript_year):
    manuscript_year = int(manuscript_year)
    start_year = manuscript_year - 9
    rows = []
    recent = older = unknown = future = 0
    year_counts = {}
    for i, ref in enumerate(references, start=1):
        year = extract_reference_year(ref)
        if year is None:
            category = "Year not detected"; unknown += 1
        elif year > manuscript_year:
            category = "Future year"; future += 1
            year_counts[year] = year_counts.get(year, 0) + 1
        elif start_year <= year <= manuscript_year:
            category = "Within last 10 years"; recent += 1
            year_counts[year] = year_counts.get(year, 0) + 1
        else:
            category = "Older than 10 years"; older += 1
            year_counts[year] = year_counts.get(year, 0) + 1
        rows.append({"Reference #": i, "Year": year if year is not None else "Not detected",
                     "Recency": category, "Reference": ref})
    total = len(references)
    return {
        "start_year": start_year, "end_year": manuscript_year, "total": total,
        "recent_count": recent, "older_count": older, "unknown_count": unknown,
        "future_count": future,
        "recent_percentage": recent / total * 100 if total else 0,
        "rows": rows,
        "year_summary": [{"Year": y, "References": c} for y, c in
                         sorted(year_counts.items(), key=lambda x: x[0], reverse=True)],
    }


# =========================================================
# CANONICAL SOURCE TYPES
# =========================================================

CANONICAL_SOURCE_TYPES = [
    "Journal Article", "Book", "Book Chapter", "Conference Proceeding",
    "Report", "Webpage / Online Document", "Other",
]

_SOURCE_TYPE_ALIASES = {
    "journal": "Journal Article", "journal article": "Journal Article",
    "journalarticle": "Journal Article", "academic journal": "Journal Article",
    "article": "Journal Article", "research article": "Journal Article",
    "journal paper": "Journal Article",
    "book": "Book", "monograph": "Book", "edited book": "Book", "textbook": "Book",
    "book chapter": "Book Chapter", "chapter": "Book Chapter",
    "book section": "Book Chapter", "chapter in book": "Book Chapter",
    "edited book chapter": "Book Chapter",
    "conference proceeding": "Conference Proceeding",
    "conference proceedings": "Conference Proceeding",
    "conference paper": "Conference Proceeding", "conference": "Conference Proceeding",
    "proceedings": "Conference Proceeding", "symposium": "Conference Proceeding",
    "conference article": "Conference Proceeding",
    "report": "Report", "technical report": "Report", "working paper": "Report",
    "government report": "Report", "research report": "Report",
    "webpage": "Webpage / Online Document", "web page": "Webpage / Online Document",
    "website": "Webpage / Online Document", "web document": "Webpage / Online Document",
    "online document": "Webpage / Online Document",
    "webpage / online document": "Webpage / Online Document",
    "web source": "Webpage / Online Document", "online source": "Webpage / Online Document",
    "other": "Other", "thesis": "Other", "dissertation": "Other",
    "newspaper article": "Other", "magazine article": "Other",
    "blog post": "Other", "unpublished": "Other",
}


def normalize_source_type(value):
    if value is None:
        return "Other"
    key = re.sub(r"\s+", " ", str(value)).strip().lower()
    if not key:
        return "Other"
    if key in _SOURCE_TYPE_ALIASES:
        return _SOURCE_TYPE_ALIASES[key]
    for alias in sorted(_SOURCE_TYPE_ALIASES, key=len, reverse=True):
        if alias in key:
            return _SOURCE_TYPE_ALIASES[alias]
    return "Other"


# =========================================================
# OPENALEX DOI VERIFICATION
# =========================================================

def _normalize_doi_for_lookup(doi):
    if not doi:
        return ""
    s = doi.strip()
    s = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", s, flags=re.I)
    s = re.sub(r"^doi\s*:\s*", "", s, flags=re.I)
    s = s.rstrip(".,;)")
    m = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", s)
    return m.group(0) if m else ""


def _fetch_openalex_metadata(doi):
    """Fetch full APA-relevant metadata from OpenAlex."""
    if not doi:
        return None
    try:
        import requests
    except Exception:
        return None

    api_key = _get_openalex_api_key()
    if not api_key:
        return None

    clean = _normalize_doi_for_lookup(doi)
    if not clean:
        return None

    url = f"https://api.openalex.org/works/doi:{clean}"
    try:
        r = requests.get(url, params={"api_key": api_key}, timeout=10)
        if r.status_code == 404:
            return {"_not_found": True}
        if r.status_code != 200:
            return None
        data = r.json()
    except Exception:
        return None

    authors = []
    for a in data.get("authorships", []) or []:
        name = (a.get("author") or {}).get("display_name")
        if name:
            authors.append(name)

    primary = data.get("primary_location") or {}
    source = primary.get("source") or {}
    biblio = data.get("biblio") or {}

    return {
        "_not_found": False,
        "title": data.get("title") or "",
        "year": data.get("publication_year"),
        "authors": authors,
        "doi": data.get("doi") or f"https://doi.org/{clean}",
        "journal": source.get("display_name"),
        "source_type": source.get("type"),
        "volume": biblio.get("volume"),
        "issue": biblio.get("issue"),
        "first_page": biblio.get("first_page"),
        "last_page": biblio.get("last_page"),
        "publisher": source.get("host_organization_name"),
        "is_oa": data.get("open_access", {}).get("is_oa"),
    }


def _verify_doi_resolution(doi):
    if not doi:
        return {"resolves": None, "reason": "No DOI supplied"}

    try:
        import requests
    except Exception:
        return {"resolves": None, "reason": "requests library missing"}

    clean = _normalize_doi_for_lookup(doi)
    if not clean:
        return {"resolves": None, "reason": "DOI could not be normalized"}

    url = f"{DOI_RESOLVER_URL}/{clean}"
    headers = {"User-Agent": "OmniCite/2.0"}

    try:
        r = requests.get(url, headers=headers, allow_redirects=True,
                         timeout=DOI_RESOLVER_TIMEOUT)
    except Exception as exc:
        return {"resolves": None, "reason": f"DOI resolver unavailable: {exc}"}

    if 200 <= r.status_code < 400:
        return {"resolves": True, "reason": "DOI resolves successfully",
                "final_url": r.url}
    if r.status_code in (404, 410):
        return {"resolves": False, "reason": f"DOI resolver returned HTTP {r.status_code}"}
    return {"resolves": None, "reason": f"Inconclusive (HTTP {r.status_code})"}


def _openalex_search_by_title(title, max_results=5):
    if not title:
        return []
    try:
        import requests
    except Exception:
        return []
    api_key = _get_openalex_api_key()
    if not api_key:
        return []
    try:
        r = requests.get(
            "https://api.openalex.org/works",
            params={"search": title, "per-page": max_results, "api_key": api_key},
            timeout=10,
        )
        if r.status_code != 200:
            return []
        return r.json().get("results", [])
    except Exception:
        return []


def _normalize_for_compare(s):
    if not s:
        return ""
    s = re.sub(r"\s+", " ", s).strip().lower()
    s = re.sub(r"[^a-z0-9 ]", "", s)
    return s


def _title_similarity(a, b):
    a_n = _normalize_for_compare(a)
    b_n = _normalize_for_compare(b)
    if not a_n or not b_n:
        return 0.0
    return difflib.SequenceMatcher(None, a_n, b_n).ratio()


def _surname_from_full_name(name):
    if not name:
        return ""
    name = name.strip()
    if "," in name:
        return name.split(",")[0].strip().lower()
    tokens = name.split()
    return tokens[-1].lower() if tokens else ""


def _extract_apa_title(reference):
    ym = re.search(r"\((?:(?:19|20)\d{2}[a-z]?|n\.d\.)\)\.\s*", reference)
    if not ym:
        return ""
    tail = reference[ym.end():].strip()
    m = re.match(r"^(.+?)\.\s", tail)
    if m:
        return m.group(1).strip()
    return tail.split(".")[0].strip()


def _compare_metadata_advanced(submitted, external):
    t = _title_similarity(
        _extract_apa_title(submitted.get("raw", "")),
        external.get("title") or "",
    )
    ref_surnames = {s.lower() for s in (submitted.get("authors") or []) if s}
    ext_surnames = {_surname_from_full_name(n) for n in (external.get("authors") or []) if n}
    ext_surnames.discard("")

    a = None
    if ref_surnames and ext_surnames:
        a = len(ref_surnames & ext_surnames) / max(1, len(ref_surnames | ext_surnames))

    sy = str(submitted.get("year") or "").strip()
    ey = str(external.get("year") or "").strip()
    year_match = (sy == ey) if sy and ey else None

    return {
        "title_similarity": t,
        "author_similarity": a,
        "year_match": year_match,
    }


def _decide_doi_status_advanced(comparison):
    t = comparison.get("title_similarity")
    a = comparison.get("author_similarity")
    y = comparison.get("year_match")

    severe_title  = t is not None and t < SEVERE_TITLE_MISMATCH
    severe_author = a is not None and a < SEVERE_AUTHOR_MISMATCH
    year_wrong    = y is False

    if severe_title and severe_author and year_wrong:
        return (
            "WITHHELD",
            "The DOI exists, but title, authors, and year strongly conflict "
            "with the DOI metadata. Possible DOI-reference substitution or "
            "fabricated citation; automatic correction withheld.",
        )

    title_good  = t is None or t >= TITLE_STRONG_MATCH
    author_good = a is None or a >= AUTHOR_STRONG_MATCH
    year_good   = y is None or y is True

    if title_good and author_good and year_good:
        return (
            "VERIFIED",
            "The DOI exists and the submitted reference is consistent "
            "with external metadata.",
        )

    title_review  = t is None or t >= TITLE_REVIEW_MATCH
    author_review = a is None or a >= AUTHOR_REVIEW_MATCH

    if title_review and author_review:
        return (
            "VERIFIED_METADATA_CORRECTED",
            "The source is identifiable, but one or more metadata fields "
            "require correction.",
        )

    return (
        "MANUAL_CHECK",
        "The DOI exists, but the submitted reference does not match "
        "external metadata strongly enough for automatic correction.",
    )


def verify_reference_against_openalex(reference, parsed):
    result = {
        "checked": False,
        "doi": parsed.get("doi"),
        "title_similarity": None,
        "author_similarity": None,
        "author_overlap": None,
        "year_match": None,
        "suspicious": False,
        "reasons": [],
        "status": "UNVERIFIED",
        "crossref_title": None,
        "crossref_authors": [],
        "source_of_truth": None,
        "doi_resolution": None,
    }

    doi = parsed.get("doi")
    if not doi:
        return result

    meta = _fetch_openalex_metadata(doi)

    if meta is None:
        result["checked"] = False
        result["status"] = "MANUAL_CHECK"
        result["reasons"].append("OpenAlex lookup failed (network).")
        return result

    if meta.get("_not_found"):
        resolution = _verify_doi_resolution(doi)
        result["doi_resolution"] = resolution

        if resolution["resolves"] is False:
            result["checked"] = True
            result["suspicious"] = True
            result["status"] = "WITHHELD"
            result["reasons"].append(
                "DOI does not resolve in OpenAlex OR through doi.org. "
                "Possible fabricated DOI."
            )
            return result

        if resolution["resolves"] is True:
            result["checked"] = True
            result["status"] = "UNVERIFIED"
            result["reasons"].append(
                "DOI resolves through doi.org but not found in OpenAlex. "
                "Cannot verify bibliographic consistency."
            )
            return result

        result["checked"] = True
        result["status"] = "MANUAL_CHECK"
        result["reasons"].append(
            "DOI verification inconclusive (resolver did not respond definitively)."
        )
        return result

    result["checked"] = True
    result["source_of_truth"] = "openalex"
    result["crossref_title"] = meta.get("title")
    result["crossref_authors"] = meta.get("authors", [])

    submitted = {
        "raw": reference,
        "authors": parsed.get("authors") or [],
        "year": parsed.get("year"),
    }
    comparison = _compare_metadata_advanced(submitted, meta)
    result["title_similarity"] = comparison["title_similarity"]
    result["author_similarity"] = comparison["author_similarity"]
    result["year_match"] = comparison["year_match"]
    result["author_overlap"] = comparison["author_similarity"]

    status, reason = _decide_doi_status_advanced(comparison)
    result["status"] = status
    result["reasons"].append(reason)

    if status == "WITHHELD":
        result["suspicious"] = True

    return result


# =========================================================
# PLACEHOLDERS FOR MISSING FIELDS
# =========================================================

PLACEHOLDER_VOL = "vol. ???"
PLACEHOLDER_NO = "no. ???"
PLACEHOLDER_PAGES = "pp. ???-???"
PLACEHOLDER_DOI = "doi: ???"


def _make_author_placeholder(count):
    if not isinstance(count, int) or count < 1:
        return "author ???"
    if count == 1:
        return "author 1"
    return ", ".join(f"author {i}" for i in range(1, count + 1))


def _guess_author_count_from_text(reference):
    year_match = re.search(r"\(((?:19|20)\d{2})[a-z]?\)", reference)
    cut = year_match.start() if year_match else min(len(reference), 250)
    block = reference[:cut].strip().rstrip(",.")
    if not block:
        return 0
    matches = re.findall(
        r"[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+\s*,\s*(?:[A-Z]\.\s*)+",
        block
    )
    return len(matches) if matches else 0


def _is_single_page_range(pages_value):
    if not pages_value:
        return False
    s = str(pages_value).strip()
    s = re.sub(r"^\s*p+p?\.\s*", "", s, flags=re.I)
    return bool(re.match(r"^\s*\d+\s*$", s))


def _insert_placeholder_after_journal(ref, placeholder):
    m = re.search(r",\s*(?:\d+(?:\s*[–-]\s*\d+)?\s*\.|https?://|doi:)", ref, re.I)
    if m:
        return ref[:m.start()] + f", {placeholder}" + ref[m.start():]
    return ref.rstrip(".").rstrip() + f", {placeholder}."


# =========================================================
# DETERMINISTIC APA IN-TEXT CORRECTION
# =========================================================

_YEAR_TOKEN_RE = re.compile(r"\b((?:19|20)\d{2}[a-z]?)\b")
_AS_CITED_IN_RE = re.compile(r"\bas\s+cited\s+in\b", re.I)


def _extract_tokens_from_citation(text):
    if not text:
        return set()
    years = set(re.findall(r"\b((?:19|20)\d{2}[a-z]?)\b", text))
    cleaned = re.sub(r"\b(?:19|20)\d{2}[a-z]?\b", " ", text)
    cleaned = re.sub(r"\bet\s+al\.?", " ", cleaned, flags=re.I)
    cleaned = re.sub(r"\b(?:and|&|see|cf|e\.g\.|i\.e\.)\b", " ", cleaned, flags=re.I)
    cleaned = re.sub(r"[\(\)\;,\.’'\-]", " ", cleaned)
    surnames = set()
    for tok in cleaned.split():
        m = re.match(r"^([A-Za-zÀ-ÖØ-öø-ÿ]+)", tok)
        if m:
            surnames.add(m.group(1).lower())
    return {(s, y) for s in surnames for y in years}


def _citation_is_hallucinated(original, corrected):
    orig_surnames = {s for s, _ in _extract_tokens_from_citation(original)}
    orig_years = {y for _, y in _extract_tokens_from_citation(original)}
    corr_surnames = {s for s, _ in _extract_tokens_from_citation(corrected)}
    corr_years = {y for _, y in _extract_tokens_from_citation(corrected)}
    return bool(corr_surnames - orig_surnames) or bool(corr_years - orig_years)


def _source_count(text):
    inner = re.sub(r"^\(|\)$", "", text or "")
    return max(1, len([p for p in inner.split(";") if p.strip()]))


def _citation_added_or_removed_sources(original, corrected):
    return _source_count(original) != _source_count(corrected)


def _strip_secondary_source_prefix(author_text):
    if not author_text:
        return author_text
    m = _AS_CITED_IN_RE.search(author_text)
    if m:
        author_text = author_text[m.end():].strip(" ,")
    author_text = re.sub(r"^(?:see|cf)\.?\s+", "", author_text, flags=re.I)
    return author_text


def _extract_surnames_from_author_text(author_text):
    if not author_text:
        return []
    author_text = _strip_secondary_source_prefix(author_text)
    if re.search(r"\s&\s|\band\b", author_text, re.I):
        parts = re.split(r"\s*&\s*|\s+\band\b\s+", author_text, flags=re.I)
        out = []
        for p in parts:
            p = re.sub(r"\bet\s+al\.?", "", p, flags=re.I).strip(" .,&")
            if p:
                out.append(p)
        return out
    if "," in author_text:
        cleaned = re.sub(r"\bet\s+al\.?", "", author_text, flags=re.I).strip(" ,")
        return [p.strip() for p in cleaned.split(",") if p.strip()]
    cleaned = re.sub(r"\bet\s+al\.?", "", author_text, flags=re.I).strip(" ,")
    return [cleaned] if cleaned else []


def revise_parenthetical_citation(content):
    parts = [p.strip() for p in content.split(";") if p.strip()]
    if not parts:
        return "(" + content.strip() + ")"
    corrected = []
    for part in parts:
        year_match = _YEAR_TOKEN_RE.search(part)
        if not year_match:
            corrected.append(part)
            continue
        year = year_match.group(1)
        author_text = part[:year_match.start()].strip(" ,")
        trailing = part[year_match.end():].strip()
        author_text = _strip_secondary_source_prefix(author_text)
        has_et_al = bool(re.search(r"\bet\s+al\.", author_text, re.I))
        surnames = _extract_surnames_from_author_text(author_text)
        if not surnames:
            corrected.append(part)
            continue
        if has_et_al or len(surnames) >= 3:
            author_final = f"{surnames[0]} et al."
        elif len(surnames) == 2:
            author_final = f"{surnames[0]} & {surnames[1]}"
        else:
            author_final = surnames[0]
        if trailing and not trailing.startswith(","):
            trailing = ", " + trailing.lstrip(", ")
        corrected.append(f"{author_final}, {year}{trailing}".rstrip())
    if len(corrected) > 1:
        def sort_key(s):
            m = re.match(r"^([A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)", s)
            return m.group(1).lower() if m else s.lower()
        corrected.sort(key=sort_key)
    return "(" + "; ".join(corrected) + ")"


def revise_narrative_citation(citation):
    raw = citation.get("raw", "")
    ctype = citation.get("type")
    authors = citation.get("authors", [])
    et_al = citation.get("et_al", False)
    if ctype != "narrative":
        return raw
    if _AS_CITED_IN_RE.search(raw):
        return revise_parenthetical_citation(re.sub(r"^\(|\)$", "", raw))

    if citation.get("malformed") and et_al and authors:
        ym = _YEAR_TOKEN_RE.search(raw)
        if ym:
            return f"{authors[0]} et al. ({ym.group(1)})"

    if len(authors) >= 3 and not et_al:
        ym = _YEAR_TOKEN_RE.search(raw)
        if ym:
            return f"{authors[0]} et al. ({ym.group(1)})"
        return raw
    if len(authors) == 2 and not et_al:
        ym = _YEAR_TOKEN_RE.search(raw)
        if ym:
            return f"{authors[0]} and {authors[1]} ({ym.group(1)})"
        return raw
    return raw


def build_local_citation_correction(citation, disambiguation_map=None):
    raw = citation.get("raw", "").strip()
    ctype = citation.get("type")
    disambiguation_map = disambiguation_map or {}
    if is_apa_disambiguated_citation(citation, disambiguation_map):
        return {"Corrected": raw, "Status": "MATCH",
                "Note": "APA author disambiguation correctly retained"}
    if ctype == "parenthetical":
        content = citation.get("parenthetical_content")
        if not content:
            m = re.match(r"^\((.*)\)$", raw)
            content = m.group(1) if m else raw
        corrected = revise_parenthetical_citation(content)
        if _source_count(raw) != _source_count(corrected):
            return {"Corrected": raw, "Status": "MATCH", "Note": ""}
        if _citation_is_hallucinated(raw, corrected):
            return {"Corrected": raw, "Status": "MATCH", "Note": ""}
    else:
        corrected = revise_narrative_citation(citation)
    if corrected.strip() == raw.strip():
        return {"Corrected": corrected, "Status": "MATCH", "Note": ""}
    notes = []
    if ctype == "parenthetical":
        if _AS_CITED_IN_RE.search(raw) and not _AS_CITED_IN_RE.search(corrected):
            notes.append("Removed secondary source ('as cited in')")
        if (re.search(r"\bet\s+al\.", corrected, re.I)
                and not re.search(r"\bet\s+al\.", raw, re.I)):
            notes.append("Collapsed 3+ authors to 'et al.'")
        raw_sources = [p.strip() for p in re.sub(r"^\(|\)$", "", raw).split(";") if p.strip()]
        corr_sources = [p.strip() for p in re.sub(r"^\(|\)$", "", corrected).split(";") if p.strip()]
        if len(raw_sources) > 1 and len(corr_sources) > 1:
            if [s.lower() for s in raw_sources] != [s.lower() for s in corr_sources]:
                if set(raw_sources) == set(corr_sources):
                    notes.append("Alphabetized multiple sources")
        if "&" in corrected and re.search(r"\band\b", raw, re.I):
            notes.append("Replaced 'and' with '&' inside parenthetical")
    else:
        if " and " in corrected and "&" in raw:
            notes.append("Replaced '&' with 'and' in narrative citation")
        if (re.search(r"\bet\s+al\.", corrected, re.I)
                and not re.search(r"\bet\s+al\.", raw, re.I)):
            notes.append("Collapsed 3+ authors to 'et al.'")
    return {"Corrected": corrected, "Status": "REVISED",
            "Note": " | ".join(notes) if notes else "APA 7 formatting adjusted"}


# =========================================================
# DETERMINISTIC APA REFERENCE CORRECTION
# =========================================================

def build_local_apa_reference_correction(reference):
    original = reference.strip()
    ref = original
    notes = []
    placeholder_flags = {
        "missing_vol": False, "missing_no": False, "missing_pp": False,
        "missing_doi": False, "missing_authors": False,
    }

    dm = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", ref, re.I)
    if dm:
        doi = dm.group(0).rstrip(".,;)")
        old = ref
        ref = re.sub(
            r"https?://(?:dx\.)?doi\.org/" + _esc(doi),
            f"https://doi.org/{doi}", ref, flags=re.I,
        )
        ref = re.sub(r"\bdoi\s*:\s*" + _esc(doi), f"https://doi.org/{doi}", ref, flags=re.I)
        if ref != old:
            notes.append("DOI normalized to https://doi.org/... format.")

    year_match = re.search(r"\(((?:19|20)\d{2})[a-z]?\)", ref)
    author_block = ref[:year_match.start()].strip() if year_match else ""

    if not author_block and year_match:
        guess = _guess_author_count_from_text(ref)
        placeholder = _make_author_placeholder(guess)
        ref = placeholder + ". " + ref[year_match.start():]
        placeholder_flags["missing_authors"] = True
        notes.append(f"Missing author list - placeholder inserted: {placeholder}")

    journal_parts = extract_journal_format_parts(ref)
    if journal_parts["is_journal"]:
        pages = journal_parts["pages"]
        if pages and _is_single_page_range(pages):
            m = re.search(r"\b" + re.escape(pages) + r"\b", ref)
            if m:
                ref = ref[:m.end()] + "-???" + ref[m.end():]
                placeholder_flags["missing_pp"] = True
                notes.append(f"Single page number - expanded to range placeholder: {pages}-???")

    if not re.search(r"10\.\d{4,9}/", ref) and not re.search(r"https?://", ref):
        ref = ref.rstrip(".").rstrip() + f". {PLACEHOLDER_DOI}"
        placeholder_flags["missing_doi"] = True
        notes.append(f"Missing DOI/URL - placeholder inserted: {PLACEHOLDER_DOI}")

    ref = ref.rstrip()
    if ref and not ref.endswith("."):
        ref = ref + "."
        notes.append("Trailing period added.")

    changed = ref != original
    return {
        "Corrected": ref,
        "Status": "REVISED" if changed else "MATCH",
        "Note": " | ".join(notes) if notes else "",
        "Placeholders": placeholder_flags,
    }


# =========================================================
# APA AUTOMATED REVIEW
# =========================================================

def _apa_status_label(status):
    status = str(status or "MANUAL_CHECK").upper().strip()
    if status in {"OK", "PASS", "MATCH", "VERIFIED"}:
        return "MATCH"
    if status in {"REVISED", "NEEDS REVIEW", "NEEDS_REVIEW",
                  "VERIFIED_METADATA_CORRECTED"}:
        return "REVISED"
    return "MANUAL CHECK"


def _safe_json_loads(text):
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
        text = re.sub(r"\s*```$", "", text)
    return json.loads(text)


def _reference_is_hallucinated(original, corrected):
    if not corrected:
        return False

    def numbers(text):
        return set(re.findall(r"\b\d+(?:\.\d+)?\b", text or ""))

    def dois(text):
        return set(re.findall(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", text or "", re.I))

    def urls(text):
        return set(re.findall(r"https?://\S+", text or "", re.I))

    if numbers(corrected) - numbers(original):
        return True
    if dois(corrected) - dois(original):
        return True
    if urls(corrected) - urls(original):
        return True
    return False


def _get_openai_client():
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        return None
    try:
        from openai import OpenAI
        return OpenAI(api_key=api_key)
    except Exception:
        return None


def review_apa_citations_with_ai(citations, parsed_references=None):
    client = _get_openai_client()
    if client is None or not citations:
        return []
    parsed_references = parsed_references or []
    disambiguation_map = build_apa_disambiguation_map(parsed_references)
    payload = [
        {
            "number": i,
            "citation": c.get("raw", ""),
            "type": c.get("type", ""),
            "apa_disambiguation_required": (
                (normalize(c.get("author")), str(c.get("year"))) in disambiguation_map
            ),
        }
        for i, c in enumerate(citations, start=1)
    ]
    reference_context = [
        {"authors": ref.get("authors") or [], "year": ref.get("year")}
        for ref in parsed_references
        if ref.get("authors") and ref.get("year")
    ]
    prompt = f"""
You are checking APA 7th edition IN-TEXT citations.

Each input item is ALREADY a separate, complete citation extracted from the PDF.
You MUST correct each item IN ISOLATION. Do not merge sources or invent authors/years.

DEFAULT APA 7 AUTHOR RULE:
- One author: use that surname.
- Two authors: parenthetical uses "&"; narrative uses "and".
- Three or more authors: "FirstSurname et al." from the first citation.

Return JSON only: {{"results": [{{"number": int, "status": "OK"|"REVISED"|"MANUAL_CHECK", "revised_citation": str, "explanation": str}}, ...]}}

REFERENCE AUTHOR/YEAR CONTEXT:
{json.dumps(reference_context, ensure_ascii=False)}

INPUT CITATIONS:
{json.dumps(payload, ensure_ascii=False)}
"""
    try:
        response = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": "You are a precise APA 7 citation editor. Return valid JSON only."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        return _safe_json_loads(response.choices[0].message.content).get("results", [])
    except Exception as exc:
        return [{"number": 0, "status": "MANUAL_CHECK", "revised_citation": "",
                 "explanation": f"OpenAI API error: {exc}"}]


def review_apa_references_with_ai(references):
    client = _get_openai_client()
    if client is None or not references:
        return []
    payload = [{"number": i, "reference": clean_text(ref)} for i, ref in enumerate(references, start=1)]
    prompt = f"""
You are checking APA 7th edition REFERENCE-LIST entries.

CRITICAL RULES:
1. Each supplied "reference" is the ORIGINAL. Correct THAT SAME reference only.
2. Do not invent missing bibliographic facts.
3. DOI should use https://doi.org/...
4. revised_reference must contain ONLY the corrected APA reference.

SOURCE TYPE - exactly one of:
  "Journal Article", "Book", "Book Chapter", "Conference Proceeding",
  "Report", "Webpage / Online Document", "Other"

Return JSON only: {{"results": [{{"number": int, "status": "OK"|"REVISED"|"MANUAL_CHECK", "revised_reference": str, "source_type": str, "italic_elements": str, "year": int|null, "missing_required_elements": [str], "explanation": str}}, ...]}}

INPUT:
{json.dumps(payload, ensure_ascii=False)}
"""
    try:
        response = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": "You are a precise APA 7 reference-list editor. Return valid JSON only."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        results = _safe_json_loads(response.choices[0].message.content).get("results", [])
        for r in results:
            r["source_type"] = normalize_source_type(r.get("source_type"))
        return results
    except Exception as exc:
        return [{"number": 0, "status": "MANUAL_CHECK", "revised_reference": "",
                 "source_type": "Other", "italic_elements": "", "year": None,
                 "missing_required_elements": [], "explanation": f"OpenAI API error: {exc}"}]


def fallback_italic_elements(source_type):
    return {
        "Journal Article": "journal title, volume",
        "Book": "book title",
        "Book Chapter": "book title",
        "Report": "report title",
        "Webpage / Online Document": "webpage title",
        "Conference Proceeding": "proceedings title",
        "Other": "",
    }.get(source_type, "")


def derive_italic_elements_from_reference(reference, source_type):
    if not reference:
        return ""

    try:
        if source_type == "Journal Article":
            parts = extract_journal_format_parts(reference)
            if parts.get("is_journal"):
                tokens = []
                if parts.get("journal"):
                    tokens.append(parts["journal"])
                if parts.get("volume"):
                    tokens.append(parts["volume"])
                return ", ".join(tokens)

        elif source_type == "Book":
            parts = extract_book_format_parts(reference)
            if parts.get("is_book") and parts.get("title"):
                return parts["title"]

        elif source_type == "Book Chapter":
            m = re.search(
                r"\bIn\s+[^()]+\(eds?\.\)\s*,?\s*(.+?)\s*\(pp?\.",
                reference, re.I,
            )
            if not m:
                m = re.search(r"\bIn\s+(.+?)\s*\(pp?\.", reference, re.I)
            if m:
                return m.group(1).strip(" ,.")

        elif source_type in (
            "Report", "Webpage / Online Document", "Conference Proceeding",
        ):
            ym = re.search(
                r"\((?:(?:19|20)\d{2}[a-z]?|n\.d\.)\)\.\s*",
                reference,
            )
            if ym:
                tail = reference[ym.end():]
                m = re.match(r"^(.+?)\.\s", tail)
                if m:
                    return m.group(1).strip()

    except Exception:
        return ""

    return ""


def build_apa_reference_comparison(references, ai_results, manuscript_year):
    by_no = {
        int(x.get("number", -1)): x
        for x in (ai_results or [])
        if str(x.get("number", "")).isdigit()
    }
    rows = []
    for i, original in enumerate(references, start=1):
        ai = by_no.get(i, {})
        original_clean = strip_markdown_markers(clean_text(original))

        ai_corrected = strip_markdown_markers(
            clean_text(ai.get("revised_reference", ""))
        ) if ai else ""
        base_text = original_clean
        if ai_corrected and ai_corrected != original_clean:
            if not _reference_is_hallucinated(original_clean, ai_corrected):
                base_text = ai_corrected

        parsed = parse_reference(base_text)
        doi_present = bool(parsed.get("doi"))

        verification = {
            "checked": False, "suspicious": False, "reasons": [],
            "crossref_title": None, "crossref_authors": [],
            "title_similarity": None, "author_similarity": None,
            "author_overlap": None, "year_match": None,
            "status": "UNVERIFIED", "source_of_truth": None,
            "doi_resolution": None,
        }
        meta = None
        if doi_present:
            verification = verify_reference_against_openalex(base_text, parsed)
            if verification.get("checked") and not verification.get("suspicious"):
                meta = _fetch_openalex_metadata(parsed["doi"])

        doi_suspicious = verification["suspicious"]
        doi_status = verification.get("status", "UNVERIFIED")

        if doi_suspicious:
            corrected = "— WITHHELD (DOI mismatch) —"
            correction_note = (
                verification["reasons"][0] if verification["reasons"] else ""
            )
            italic_elements = ""
            placeholders = {}
            source_type = normalize_source_type(ai.get("source_type")) or "Other"

        elif meta and doi_status in ("VERIFIED", "VERIFIED_METADATA_CORRECTED"):
            source_type = normalize_source_type(
                ai.get("source_type") or _source_type_from_openalex(meta)
            )

            ai_formatted = format_apa_with_ai(
                meta, original_clean, source_type
            )

            if ai_formatted.get("corrected_reference"):
                corrected       = ai_formatted["corrected_reference"]
                italic_elements = ai_formatted["italic_elements"]
                source_type     = ai_formatted["source_type"]
                correction_note = (
                    ai_formatted.get("explanation")
                    or "Formatted from OpenAlex metadata via OpenAI."
                )
                placeholders = {}

            else:
                oa_ref, oa_note = build_apa_reference_from_openalex(
                    meta, parsed, original_clean
                )
                if oa_ref:
                    corrected = oa_ref
                    correction_note = oa_note
                else:
                    local = build_local_apa_reference_correction(base_text)
                    corrected = local["Corrected"]
                    correction_note = local.get("Note", "")
                    placeholders = local.get("Placeholders", {})

                italic_elements = derive_italic_elements_from_reference(
                    corrected, source_type
                )

        else:
            local = build_local_apa_reference_correction(base_text)
            corrected = local["Corrected"]
            correction_note = local.get("Note", "")
            placeholders = local.get("Placeholders", {})
            source_type = normalize_source_type(ai.get("source_type"))
            if source_type == "Other" and not ai.get("source_type"):
                source_type = detect_apa_source_type(corrected)
            italic_elements = (
                derive_italic_elements_from_reference(corrected, source_type)
                or (ai.get("italic_elements") or "").strip()
                or fallback_italic_elements(source_type)
            )

        year = ai.get("year")
        if not isinstance(year, int):
            year = extract_reference_year(corrected)

        if doi_suspicious:
            status = "WITHHELD"
        elif doi_status in ("VERIFIED", "VERIFIED_METADATA_CORRECTED"):
            status = doi_status
        else:
            status = _apa_status_label(ai.get("status")) if ai else "NOT AI CHECKED"

        rows.append({
            "No.": i,
            "Source Type": source_type,
            "Year": year,
            "Original Reference": original_clean,
            "Corrected Version": corrected,
            "Italicized in APA": italic_elements,
            "Status": status,
            "Verification Status": doi_status,
            "Verification Reason": " | ".join(verification["reasons"]),
            "Source of Truth": verification.get("source_of_truth") or "local",
            "Missing Required Elements": ", ".join(
                str(x) for x in (ai.get("missing_required_elements") or [])
            ),
            "AI Explanation": ai.get("explanation", ""),
            "Correction Note": correction_note,
            "Placeholders": placeholders,
            "DOI Verified": verification["checked"],
            "DOI Suspicious": doi_suspicious,
            "DOI Verification Reasons": " | ".join(verification["reasons"]),
            "Title Similarity": verification.get("title_similarity"),
            "Author Similarity": verification.get("author_similarity"),
            "Author Overlap": verification.get("author_overlap"),
            "Year Match": verification.get("year_match"),
            "OpenAlex Title": verification.get("crossref_title"),
            "OpenAlex Authors": ", ".join(
                verification.get("crossref_authors", [])[:5]
            ),
            "OpenAlex Journal": meta.get("journal") if meta else None,
            "OpenAlex Volume": meta.get("volume") if meta else None,
            "OpenAlex Issue": meta.get("issue") if meta else None,
            "OpenAlex Pages": (
                f"{meta.get('first_page')}–{meta.get('last_page')}"
                if meta and meta.get("first_page") else None
            ),
        })
    return rows


def _format_apa_authors_from_openalex(authors_full, reference_authors):
    def to_surname(name):
        if not name:
            return ""
        if "," in name:
            return name.split(",")[0].strip()
        parts = name.split()
        return parts[-1] if parts else name

    if reference_authors:
        surnames = reference_authors
    else:
        surnames = [to_surname(n) for n in authors_full if n]
        surnames = [s for s in surnames if s]

    if not surnames:
        return ""
    if len(surnames) == 1:
        return surnames[0]
    if len(surnames) == 2:
        return f"{surnames[0]}, & {surnames[1]}"
    if len(surnames) <= 20:
        return ", ".join(surnames[:-1]) + f", & {surnames[-1]}"
    return ", ".join(surnames[:19]) + ", ... " + surnames[-1]


def build_apa_reference_from_openalex(meta, parsed_reference, original_reference):
    if not meta or meta.get("_not_found"):
        return None, ""

    authors_str = _format_apa_authors_from_openalex(
        meta.get("authors") or [],
        parsed_reference.get("authors") or [],
    )
    year = meta.get("year")
    title = meta.get("title") or ""
    journal = meta.get("journal") or ""
    volume = meta.get("volume") or ""
    issue = meta.get("issue") or ""
    first_page = meta.get("first_page") or ""
    last_page = meta.get("last_page") or ""
    doi = meta.get("doi") or ""

    parts = []
    if authors_str:
        parts.append(f"{authors_str} ({year})." if year else f"{authors_str}.")
    else:
        parts.append(f"({year})." if year else "")

    if title:
        parts.append(f"{title.rstrip('.')}.")

    if journal:
        loc = journal
        if volume:
            loc += f", {volume}"
        if issue:
            loc += f"({issue})"
        if first_page:
            if last_page and last_page != first_page:
                loc += f", {first_page}–{last_page}"
            else:
                loc += f", {first_page}"
        parts.append(f"{loc}.")

    if doi:
        parts.append(doi if doi.startswith("http") else f"https://doi.org/{doi}")

    reconstructed = " ".join(p.strip() for p in parts if p.strip())
    reconstructed = re.sub(r"\s+", " ", reconstructed).strip()

    note = "Corrected using OpenAlex bibliographic record."
    return reconstructed, note


def format_apa_with_ai(openalex_meta, original_reference, source_type):
    client = _get_openai_client()
    if client is None:
        return {
            "corrected_reference": "",
            "italic_elements": "",
            "source_type": source_type,
            "explanation": "OpenAI client unavailable",
        }

    translation = ""
    m = re.search(r"\[([^\]]+)\]", original_reference)
    if m:
        translation = m.group(0)

    ground_truth = {
        "title":        openalex_meta.get("title"),
        "year":         openalex_meta.get("year"),
        "authors":      openalex_meta.get("authors") or [],
        "journal":      openalex_meta.get("journal"),
        "volume":       openalex_meta.get("volume"),
        "issue":        openalex_meta.get("issue"),
        "first_page":   openalex_meta.get("first_page"),
        "last_page":    openalex_meta.get("last_page"),
        "doi":          openalex_meta.get("doi"),
        "source_type":  source_type,
        "translation_bracket_from_original": translation,
    }

    prompt = f"""
You are formatting ONE APA 7th edition REFERENCE-LIST entry.

The bibliographic facts below come from a VERIFIED OpenAlex record.
You MUST NOT change any of these fields:
  - Author surnames, initials, or ordering
  - Publication year
  - Title text
  - Journal / book / source name
  - Volume, issue, page numbers
  - DOI or URL

You MAY only:
  1. Reorder fields into correct APA 7 order.
  2. Fix punctuation (periods, commas, ampersands, en dashes).
  3. Convert the author list to APA 7 form.
  4. Apply APA 7 capitalization:
       - Article/book title → sentence case
       - Journal name → title case
  5. Preserve the translation bracket from the original if present.
  6. Decide the exact italic spans.

SOURCE TYPE (choose one):
  "Journal Article", "Book", "Book Chapter", "Conference Proceeding",
  "Report", "Webpage / Online Document", "Other"

APA 7 ITALIC RULES:
  - Journal Article  → journal name + volume number (issue NOT italic, pages NOT italic)
  - Book             → book title
  - Book Chapter     → book title
  - Report           → report title
  - Webpage          → webpage title
  - Conference       → proceedings title

CRITICAL — `italic_elements` MUST be a comma-separated list of LITERAL
substrings that appear VERBATIM inside your `corrected_reference`.
Do NOT return generic phrases like "journal title" or "volume".
For a journal article the value should look like:
  "Journal of Education Research, 5"
Note that the comma and space between journal and volume are INCLUDED in
the substring, so the volume number gets italicized together with the
journal name.

Return JSON ONLY:
{{
  "corrected_reference": "the full APA 7 reference string",
  "italic_elements":     "literal, substrings, to, italicize",
  "source_type":         "one of the source types above",
  "explanation":         "short note"
}}

GROUND TRUTH (do NOT alter these values):
{json.dumps(ground_truth, ensure_ascii=False, indent=2)}

ORIGINAL REFERENCE (for reference only — do NOT copy its errors):
{original_reference}
"""

    try:
        response = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system",
                 "content": "You format APA 7 references. Never invent "
                            "bibliographic facts. Return valid JSON only."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        data = _safe_json_loads(response.choices[0].message.content)

        corrected = (data.get("corrected_reference") or "").strip()
        italics   = (data.get("italic_elements") or "").strip()
        stype     = normalize_source_type(data.get("source_type") or source_type)
        note      = data.get("explanation") or ""

        if corrected and _reference_is_hallucinated(original_reference, corrected):
            return {
                "corrected_reference": "",
                "italic_elements": "",
                "source_type": source_type,
                "explanation": "AI output rejected (invented data).",
            }

        if italics:
            tokens = [t.strip() for t in italics.split(",") if t.strip()]
            tokens = [t for t in tokens if t in corrected]
            italics = ", ".join(tokens)

        return {
            "corrected_reference": corrected,
            "italic_elements": italics,
            "source_type": stype,
            "explanation": note,
        }

    except Exception as exc:
        return {
            "corrected_reference": "",
            "italic_elements": "",
            "source_type": source_type,
            "explanation": f"OpenAI error: {exc}",
        }


def _source_type_from_openalex(meta):
    if not meta:
        return "Other"
    t = (meta.get("source_type") or "").lower()
    if t == "journal":
        return "Journal Article"
    if t == "book":
        return "Book"
    if t in ("book series", "ebook platform"):
        return "Book Chapter"
    if t == "conference":
        return "Conference Proceeding"
    if t == "repository":
        return "Webpage / Online Document"
    return "Other"


def _citation_breaks_apa_connector_rules(original, corrected, ctype):
    if not corrected:
        return False
    corr = (corrected or "").lower()
    orig = (original or "").lower()
    corr_amp = corr.count("&")
    corr_and = len(re.findall(r"\band\b", corr))
    orig_amp = orig.count("&")
    orig_and = len(re.findall(r"\band\b", orig))
    if ctype == "parenthetical" and corr_and > orig_and:
        return True
    if ctype == "narrative" and corr_amp > orig_amp:
        return True
    return False


def build_apa_citation_comparison(citations, ai_results, parsed_references=None):
    by_no = {
        int(x.get("number", -1)): x
        for x in (ai_results or [])
        if str(x.get("number", "")).isdigit()
    }
    rows = []
    disambiguation_map = build_apa_disambiguation_map(parsed_references or [])
    for i, citation in enumerate(citations, start=1):
        local = build_local_citation_correction(citation, disambiguation_map)
        protected_disambiguation = is_apa_disambiguated_citation(citation, disambiguation_map)
        ai = by_no.get(i, {})
        original = strip_markdown_markers(clean_text(citation.get("raw", "")))
        ai_corrected = strip_markdown_markers(
            clean_text(ai.get("revised_citation", ""))
        ) if ai else ""
        local_corrected = local["Corrected"]
        hallucinated = bool(ai_corrected) and _citation_is_hallucinated(original, ai_corrected)
        source_count_changed = bool(ai_corrected) and _citation_added_or_removed_sources(original, ai_corrected)
        connector_violation = bool(ai_corrected) and _citation_breaks_apa_connector_rules(
            original, ai_corrected, citation.get("type")
        )
        ai_is_safe = (
            ai_corrected
            and ai_corrected != original
            and not protected_disambiguation
            and not hallucinated
            and not source_count_changed
            and not connector_violation
        )
        if ai_is_safe:
            corrected = ai_corrected
            status = _apa_status_label(ai.get("status")) if ai else "REVISED"
            note = ai.get("explanation", "") or local["Note"]
        else:
            corrected = local_corrected
            status = local["Status"]
            note = local["Note"]
            if hallucinated:
                note = "AI revision rejected (invented authors/years)."
            elif source_count_changed:
                note = "AI revision rejected (changed number of sources)."
            elif connector_violation:
                note = "AI revision rejected ('&' vs 'and' violation)."
        if corrected.strip() == original.strip():
            status = "MATCH"
        rows.append({
            "No.": i,
            "Type": citation.get("type", "").title(),
            "Original Citation": original,
            "Revised Citation": corrected,
            "Status": status,
            "Notes": note,
        })
    return rows


def analyze_apa_locally(uploaded_file, manuscript_year):
    uploaded_file.seek(0)
    full_text, pages, removed_running_text = extract_pdf_text(uploaded_file)
    uploaded_file.seek(0)
    style_spans = extract_pdf_style_spans(uploaded_file)
    full_text = clean_text(full_text)
    reference_text, heading, body_text = find_reference_section(full_text)
    if reference_text is None:
        raise ValueError("I could not detect a References section.")

    references = split_references(reference_text)
    parsed_references = [parse_reference(ref) for ref in references]
    apa_disambiguation_map = build_apa_disambiguation_map(parsed_references)
    citations = extract_all_citations(body_text)
    citation_stats = calculate_citation_statistics(citations)
    matching_results = match_citations_to_references(citations, parsed_references)
    missing_references = find_missing_references(citations, parsed_references)
    duplicates = detect_duplicates(parsed_references)
    recency = calculate_reference_recency(references, manuscript_year)

    local_reference_checks = []
    for i, ref in enumerate(references, start=1):
        status, issues, warnings, italic_info = check_apa_reference(ref, style_spans)
        local_reference_checks.append({
            "Reference #": i, "APA Status": status,
            "Source Type": italic_info.get("source_type", detect_apa_source_type(ref)),
            "Issues": " | ".join(issues), "Warnings": " | ".join(warnings),
        })

    return {
        "filename": uploaded_file.name, "heading": heading,
        "references": references, "parsed_references": parsed_references,
        "apa_disambiguation_map": apa_disambiguation_map,
        "citations": citations, "citation_stats": citation_stats,
        "matching_results": matching_results, "missing_references": missing_references,
        "duplicates": duplicates, "recency": recency,
        "local_reference_checks": local_reference_checks,
        "removed_running_text": removed_running_text,
        "manuscript_year": manuscript_year, "ai_complete": False,
    }


def enrich_apa_with_ai(result):
    citation_ai = review_apa_citations_with_ai(
        result["citations"], result.get("parsed_references", [])
    )
    reference_ai = review_apa_references_with_ai(result["references"])
    result["citation_comparison"] = build_apa_citation_comparison(
        result["citations"], citation_ai, result.get("parsed_references", [])
    )
    result["reference_comparison"] = build_apa_reference_comparison(
        result["references"], reference_ai, result["manuscript_year"]
    )
    result["ai_complete"] = bool(_get_openai_client())
    return result


# =========================================================
# DOCX EXPORT
# =========================================================

def _set_run_font(run, size_pt=11, bold=False, italic=False, color=None):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


RED = RGBColor(0xC0, 0x00, 0x00)


def _add_run(paragraph, text, size_pt=11, bold=False, italic=False, red=False):
    r = paragraph.add_run(text)
    _set_run_font(r, size_pt=size_pt, bold=bold, italic=italic,
                  color=RED if red else None)
    return r


def _add_red_italic_run(paragraph, text, size_pt=11):
    _add_run(paragraph, text, size_pt=size_pt, bold=True, italic=True, red=True)


_PLACEHOLDER_RE = re.compile(
    r"(author \?{3}"
    r"|author \d+(?:, author \d+)+"
    r"|vol\. \?{3}"
    r"|no\. \?{3}"
    r"|pp\. \d+-\?{3}"
    r"|pp\. \?{3}-\?{3}"
    r"|doi: \?{3})",
    re.I,
)


def _emit_with_placeholders(paragraph, text, italic=False):
    pos = 0
    for m in _PLACEHOLDER_RE.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], size_pt=11, italic=italic)
        matched = m.group(0)
        _add_run(paragraph, matched, size_pt=11, italic=True, bold=True, red=True)
        pos = m.end()
    if pos < len(text):
        _add_run(paragraph, text[pos:], size_pt=11, italic=italic)


def _add_styled_reference(paragraph, text, italic_elements):
    if isinstance(italic_elements, list):
        tokens = [t.strip() for t in italic_elements if t and t.strip()]
    else:
        tokens = [t.strip() for t in (italic_elements or "").split(",") if t.strip()]
    if not tokens:
        _emit_with_placeholders(paragraph, text, italic=False)
        return

    escaped = sorted({_esc(t) for t in tokens}, key=len, reverse=True)
    pattern = _safe_compile(
        r"(?<!\w)(" + "|".join(escaped) + r")(?!\w)",
        re.I,
    )
    if pattern is None:
        _emit_with_placeholders(paragraph, text, italic=False)
        return

    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            _emit_with_placeholders(paragraph, text[pos:m.start()], italic=False)
        _emit_with_placeholders(paragraph, m.group(0), italic=True)
        pos = m.end()
    if pos < len(text):
        _emit_with_placeholders(paragraph, text[pos:], italic=False)


def _docx_set_default_font(document, font_name="Times New Roman", size_pt=11):
    style = document.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(size_pt)


def build_correction_docx(result):
    doc = Document()
    _docx_set_default_font(doc)

    filename = result.get("filename", "manuscript.pdf")

    title = doc.add_heading(level=0)
    tr = title.add_run(filename)
    _set_run_font(tr, size_pt=18, bold=True)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    sr = subtitle.add_run("APA 7th Edition — Citation & Reference Correction Report")
    _set_run_font(sr, size_pt=11, italic=True)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = doc.add_paragraph()
    mr = meta.add_run(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    _set_run_font(mr, size_pt=9, italic=True)
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # ---------- Section 1: In-text Citation List ----------
    h1 = doc.add_heading(level=1)
    hr = h1.add_run("1. In-text Citation List")
    _set_run_font(hr, size_pt=14, bold=True)

    citation_rows = result.get("citation_comparison") or build_apa_citation_comparison(
        result.get("citations", []), []
    )

    missing_citation_keys = set()
    for m in result.get("missing_references", []) or []:
        missing_citation_keys.add((normalize(m.get("Author")), m.get("Year")))

    if not citation_rows:
        p = doc.add_paragraph()
        _set_run_font(p.add_run("No in-text citations were detected."), italic=True)
    else:
        for row in citation_rows:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)

            author = ""
            year = None
            raw = row.get("Original Citation", "")
            m = re.match(r"^\(?([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)", raw)
            if m:
                author = m.group(1)
            ym = re.search(r"\b((?:19|20)\d{2})\b", raw)
            if ym:
                year = ym.group(1)

            is_missing_from_refs = (normalize(author), year) in missing_citation_keys

            r1 = p.add_run("Original : ")
            _set_run_font(r1, size_pt=11, bold=True)
            r2 = p.add_run(row.get("Original Citation", ""))
            _set_run_font(r2, size_pt=11)
            if is_missing_from_refs:
                _add_red_italic_run(p, "  ← MISSING FROM REFERENCES")

            p2 = doc.add_paragraph()
            p2.paragraph_format.space_after = Pt(8)
            r3 = p2.add_run("Corrected: ")
            _set_run_font(r3, size_pt=11, bold=True)
            r4 = p2.add_run(row.get("Revised Citation", ""))
            _set_run_font(r4, size_pt=11, italic=False)

    doc.add_page_break()

    # ---------- Section 2: Reference List ----------
    h2 = doc.add_heading(level=1)
    hr2 = h2.add_run("2. Reference List (APA 7th Edition)")
    _set_run_font(hr2, size_pt=14, bold=True)

    reference_rows = result.get("reference_comparison") or build_apa_reference_comparison(
        result.get("references", []), [], result.get("manuscript_year")
    )

    matching = result.get("matching_results", []) or []
    uncited_ref_nos = {row.get("Reference #") for row in matching if not row.get("Cited")}

    if not reference_rows:
        p = doc.add_paragraph()
        _set_run_font(p.add_run("No references were detected."), italic=True)
    else:
        def sort_key(row):
            txt = row.get("Corrected Version", "") or row.get("Original Reference", "")
            m = re.match(r"^([A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)", txt)
            return m.group(1).lower() if m else txt.lower()

        sorted_rows = sorted(reference_rows, key=sort_key)

        for row in sorted_rows:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)

            original = row.get("Original Reference", "")
            if row.get("DOI Suspicious"):
                _add_run(p, original, size_pt=11, red=True)
            else:
                _add_run(p, original, size_pt=11)

            if row.get("No.") in uncited_ref_nos:
                _add_red_italic_run(p, "  ← NOT CITED IN TEXT")

            if row.get("DOI Suspicious"):
                withheld = doc.add_paragraph()
                withheld.paragraph_format.left_indent = Inches(0.5)
                withheld.paragraph_format.space_after = Pt(6)
                _add_run(
                    withheld,
                    "Corrected version withheld — the DOI in this reference "
                    "does not match the claimed title/authors. Manual verification "
                    "required before any correction is applied.",
                    size_pt=10, italic=True, bold=True, red=True,
                )
            else:
                p2 = doc.add_paragraph()
                p2.paragraph_format.space_after = Pt(4)
                p2.paragraph_format.left_indent = Inches(0.5)
                _add_run(p2, "Corrected: ", bold=True, size_pt=11)
                corrected = row.get("Corrected Version", "")
                italic_elements = row.get("Italicized in APA", "")
                _add_styled_reference(p2, corrected, italic_elements)

                if not row.get("DOI Suspicious"):
                    source_tag = (
                        "  [corrected from OpenAlex]"
                        if row.get("DOI Verified") and row.get("OpenAlex Journal")
                        else "  [APA tidy-up]"
                    )
                    _add_run(p2, source_tag, size_pt=9, italic=True)

                vstatus = row.get("Verification Status", "")
                if vstatus and vstatus not in ("VERIFIED",):
                    status_p = doc.add_paragraph()
                    status_p.paragraph_format.left_indent = Inches(0.5)
                    _add_run(
                        status_p,
                        f"Verification: {vstatus} — {row.get('Verification Reason', '')}",
                        size_pt=9, italic=True,
                        red=(vstatus in ("MANUAL_CHECK", "WITHHELD")),
                    )

                correction_note = row.get("Correction Note", "")
                if correction_note:
                    note_p = doc.add_paragraph()
                    note_p.paragraph_format.left_indent = Inches(0.5)
                    _add_run(note_p, f"Fix applied: {correction_note}",
                             size_pt=10, italic=True)

                placeholders = row.get("Placeholders") or {}
                active_ph = [k for k, v in placeholders.items() if v]
                if active_ph:
                    ph_p = doc.add_paragraph()
                    ph_p.paragraph_format.left_indent = Inches(0.5)
                    _add_run(ph_p,
                             "Placeholder used — fill in before submission: "
                             + ", ".join(active_ph),
                             size_pt=10, italic=True, bold=True, red=True)

            if row.get("DOI Suspicious"):
                warn_p = doc.add_paragraph()
                warn_p.paragraph_format.left_indent = Inches(0.5)
                reasons = row.get("DOI Verification Reasons", "")
                _add_run(warn_p, f"⚠ Possible fabricated reference: {reasons}",
                         size_pt=10, italic=True, bold=True, red=True)
                if row.get("OpenAlex Title"):
                    _add_run(warn_p, f'\n  OpenAlex says: "{row["OpenAlex Title"]}"',
                             size_pt=10, italic=True)
                if row.get("OpenAlex Authors"):
                    _add_run(warn_p, f"\n  Authors: {row['OpenAlex Authors']}",
                             size_pt=10, italic=True)

            missing_elems = row.get("Missing Required Elements", "")
            if missing_elems:
                note_p = doc.add_paragraph()
                note_p.paragraph_format.left_indent = Inches(0.5)
                _add_run(note_p, f"Missing required element(s): {missing_elems}",
                         size_pt=10, italic=True, red=True)

    # ---------- Section 3: Summary ----------
    doc.add_page_break()
    h3 = doc.add_heading(level=1)
    hr3 = h3.add_run("3. Summary")
    _set_run_font(hr3, size_pt=14, bold=True)

    stats = result.get("citation_stats", {})
    recency = result.get("recency", {})

    total_refs_now = len(result.get("references", []))
    doi_checked = sum(1 for r in reference_rows if r.get("DOI Verified"))
    doi_suspicious = sum(1 for r in reference_rows if r.get("DOI Suspicious"))
    doi_suspicious_pct = doi_suspicious / total_refs_now * 100 if total_refs_now else 0
    single_page_count = sum(
        1 for r in reference_rows
        if (r.get("Placeholders") or {}).get("missing_pp")
    )

    summary_rows = [
        ("Total references", str(total_refs_now), False),
        ("Total in-text citations", str(stats.get("total", 0)), False),
        ("  • Narrative", str(stats.get("narrative", 0)), False),
        ("  • Parenthetical", str(stats.get("parenthetical", 0)), False),
        ("Citations missing from references",
         str(len(result.get("missing_references", []))), False),
        ("References missing from citations",
         str(len(uncited_ref_nos)), False),
        ("DOI checked via OpenAlex", str(doi_checked), False),
        ("DOI suspicious (possible fabricated references)",
         f"{doi_suspicious} ({doi_suspicious_pct:.1f}%)",
         doi_suspicious > 0),
        ("Corrections withheld due to DOI mismatch",
         str(doi_suspicious), doi_suspicious > 0),
        ("References with incomplete page range",
         str(single_page_count), False),
        ("% references within last 10 years "
         f"({recency.get('start_year')}-{recency.get('end_year')})",
         f"{recency.get('recent_percentage', 0):.1f}%", False),
    ]

    summary_table = doc.add_table(rows=1, cols=2)
    summary_table.style = "Light Grid Accent 1"
    summary_table.autofit = True

    hdr_cells = summary_table.rows[0].cells
    for cell, text in zip(hdr_cells, ["Metric", "Value"]):
        cell.text = ""
        run = cell.paragraphs[0].add_run(text)
        _set_run_font(run, size_pt=10, bold=True)

    for label, value, is_warning in summary_rows:
        cells = summary_table.add_row().cells
        cells[0].text = ""
        run_label = cells[0].paragraphs[0].add_run(label)
        _set_run_font(run_label, size_pt=10,
                      bold=is_warning,
                      color=RED if is_warning else None)
        cells[1].text = ""
        run_val = cells[1].paragraphs[0].add_run(value)
        _set_run_font(run_val, size_pt=10,
                      bold=is_warning,
                      color=RED if is_warning else None)

    for row in summary_table.rows[1:]:
        label_run = row.cells[0].paragraphs[0].runs[0]
        if label_run.text.startswith("  •"):
            row.cells[0].paragraphs[0].paragraph_format.left_indent = Inches(0.25)

    # ---------- Section 3.1: Source Type Distribution ----------
    h31 = doc.add_heading(level=2)
    hr31 = h31.add_run("3.1 Source Type Distribution")
    _set_run_font(hr31, size_pt=12, bold=True)

    source_counts = Counter(
        (row.get("Source Type") or "Other") for row in reference_rows
    )
    source_order = CANONICAL_SOURCE_TYPES
    extra_source_types = sorted(x for x in source_counts if x not in source_order)
    ordered_sources = source_order + extra_source_types

    if total_refs_now == 0:
        p = doc.add_paragraph()
        _set_run_font(p.add_run("No references were detected."), italic=True)
    else:
        table = doc.add_table(rows=1, cols=3)
        table.style = "Light Grid Accent 1"
        table.autofit = True

        hdr = table.rows[0].cells
        for cell, text in zip(hdr, ["Source Type", "Count", "Percentage"]):
            cell.text = ""
            run = cell.paragraphs[0].add_run(text)
            _set_run_font(run, size_pt=10, bold=True)

        for source_type in ordered_sources:
            count = source_counts.get(source_type, 0)
            if count == 0 and source_type not in source_order:
                continue
            pct = count / total_refs_now * 100 if total_refs_now else 0
            cells = table.add_row().cells
            for cell, val in zip(
                cells,
                [source_type, str(count), f"{pct:.1f}%"],
            ):
                cell.text = ""
                run = cell.paragraphs[0].add_run(val)
                _set_run_font(run, size_pt=10)

        cells = table.add_row().cells
        for cell, val in zip(cells, ["Total", str(total_refs_now), "100.0%"]):
            cell.text = ""
            run = cell.paragraphs[0].add_run(val)
            _set_run_font(run, size_pt=10, bold=True)

        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_after = Pt(2)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


# =========================================================
# RENDER (called from app.py)
# =========================================================

def render():
    st.title("OmniCite Auditor - APA 7th Edition")
    st.caption("Narrative + parenthetical in-text citations ↔ APA reference list")

    uploaded_files = st.file_uploader(
        "Upload manuscript PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload up to 5 manuscripts in one batch.",
        key="apa_file_uploader",
    )

    manuscript_year = st.number_input(
        "Manuscript publication year", min_value=1900, max_value=2100,
        value=datetime.now().year, step=1,
        key="apa_manuscript_year",
    )

    if "apa_batch_results" not in st.session_state:
        st.session_state["apa_batch_results"] = {}

    if uploaded_files:
        if len(uploaded_files) > 5:
            st.error(f"You uploaded {len(uploaded_files)} manuscripts. The maximum batch size is 5 PDFs.")
            st.stop()

        with st.container(key="blue_btn_extract_apa"):
            extract_batch = st.button(
                f"Extract & Review ({len(uploaded_files)} manuscript{'s' if len(uploaded_files) != 1 else ''})",
                use_container_width=True,
                key="apa_extract_button",
            )

        if extract_batch:
            local_results = {}
            progress = st.progress(0, text="Extracting APA citations and references locally...")
            for file_index, uploaded_file in enumerate(uploaded_files, start=1):
                key = f"{file_index}::{uploaded_file.name}"
                try:
                    progress.progress(
                        int(((file_index - 1) / len(uploaded_files)) * 100),
                        text=f"Extracting {file_index}/{len(uploaded_files)}: {uploaded_file.name}"
                    )
                    local_results[key] = analyze_apa_locally(uploaded_file, int(manuscript_year))
                except Exception as exc:
                    st.error(f"{uploaded_file.name} failed:\n\n{traceback.format_exc()}")
                    local_results[key] = {
                        "filename": uploaded_file.name,
                        "error": str(exc),
                        "manuscript_year": int(manuscript_year),
                    }
            progress.empty()
            st.session_state["apa_batch_results"] = local_results

        results = st.session_state.get("apa_batch_results", {})

        if results:
            valid_results = {k: v for k, v in results.items() if not v.get("error")}
            ai_ready = bool(valid_results) and all(
                v.get("ai_complete", False) for v in valid_results.values()
            )
            preview_key = next(iter(valid_results), None)
            result = valid_results.get(preview_key) if preview_key else None

            if result is None:
                for failed in results.values():
                    if failed.get("error"):
                        st.error(f"{failed.get('filename', 'Manuscript')}: {failed['error']}")
            else:
                selector_keys = list(valid_results.keys())
                selected_key = st.selectbox(
                    "Select manuscript to review",
                    selector_keys,
                    format_func=lambda k: valid_results[k].get("filename", k),
                    key="apa_local_selected_manuscript",
                )
                result = valid_results[selected_key]

                citations = result["citations"]
                references = result["references"]
                stats = result["citation_stats"]
                matching = result["matching_results"]
                missing = result["missing_references"]

                ai_citation_rows = (
                    result.get("citation_comparison")
                    or build_apa_citation_comparison(citations, [])
                )
                narrative_rows = [
                    r for r in ai_citation_rows
                    if str(r.get("Type", "")).lower() == "narrative"
                ]
                parenthetical_rows = [
                    r for r in ai_citation_rows
                    if str(r.get("Type", "")).lower() == "parenthetical"
                ]

                st.caption(
                    f"References: {len(references)}  |  "
                    f"In-text citations: {len(citations)}"
                )

                with st.expander(
                    f"Narrative In-text Citations ({len(narrative_rows)})",
                    expanded=False,
                    key=f"narr_{selected_key}",
                ):
                    if narrative_rows:
                        st.dataframe(
                            pd.DataFrame([
                                {"Original Citation": r.get("Original Citation", "")}
                                for r in narrative_rows
                            ]),
                            use_container_width=True, hide_index=True,
                            height=min(230, max(90, 38 * (len(narrative_rows) + 1))),
                        )
                    else:
                        st.info("No narrative in-text citations were detected in this manuscript.")

                with st.expander(
                    f"Parenthetical In-text Citations ({len(parenthetical_rows)})",
                    expanded=False,
                    key=f"paren_{selected_key}",
                ):
                    if parenthetical_rows:
                        st.dataframe(
                            pd.DataFrame([
                                {"Original Citation": r.get("Original Citation", "")}
                                for r in parenthetical_rows
                            ]),
                            use_container_width=True, hide_index=True,
                            height=min(230, max(90, 38 * (len(parenthetical_rows) + 1))),
                        )
                    else:
                        st.info("No parenthetical in-text citations were detected in this manuscript.")

                citations_missing_reference = len(missing)
                references_missing_rows = [row for row in matching if not row.get("Cited")]
                references_missing_citation = len(references_missing_rows)

                if citations_missing_reference:
                    with st.expander(
                        f"Citations Missing from References ({citations_missing_reference})",
                        expanded=False,
                        key=f"missref_{selected_key}",
                    ):
                        st.dataframe(
                            pd.DataFrame(missing)[["Citation", "Author", "Year"]],
                            use_container_width=True,
                            hide_index=True,
                        )

                if references_missing_citation:
                    with st.expander(
                        f"References Missing from Citations ({references_missing_citation})",
                        expanded=False,
                        key=f"misscite_{selected_key}",
                    ):
                        st.dataframe(
                            pd.DataFrame([
                                {
                                    "Reference #": row.get("Reference #"),
                                    "Author": row.get("Author"),
                                    "Year": row.get("Year"),
                                    "Reference": row.get("Reference"),
                                }
                                for row in references_missing_rows
                            ]),
                            use_container_width=True,
                            hide_index=True,
                        )

                with st.container(key="green_btn_ai_apa"):
                    run_ai = st.button(
                        "Automated Processing Check",
                        use_container_width=True,
                        disabled=not bool(_get_openai_client()),
                        help=None if _get_openai_client() else "Set key to enable automated APA correction.",
                        key="run_automated_apa_review_below_missing_checks",
                    )

                if run_ai:
                    ai_progress = st.progress(0, text="Running automated APA review...")
                    updated = {}
                    for idx, (key, batch_result) in enumerate(valid_results.items(), start=1):
                        ai_progress.progress(
                            int(((idx - 1) / len(valid_results)) * 100),
                            text=f"APA review {idx}/{len(valid_results)}: {batch_result['filename']}",
                        )
                        try:
                            enriched = enrich_apa_with_ai(batch_result)
                        except Exception:
                            st.error(f"AI review failed for {batch_result.get('filename')}:\n\n{traceback.format_exc()}")
                            enriched = batch_result
                        updated[key] = enriched
                        st.session_state["apa_batch_results"][key] = enriched
                    ai_progress.empty()
                    st.rerun()

                if ai_ready:
                    keys = list(valid_results.keys())
                    selected_key = st.selectbox(
                        "Manuscript to review",
                        keys,
                        format_func=lambda k: valid_results[k].get("filename", k),
                        key="apa_ai_selected_manuscript",
                    )
                    result = valid_results[selected_key]

                    citations = result["citations"]
                    references = result["references"]
                    stats = result["citation_stats"]
                    matching = result["matching_results"]
                    missing = result["missing_references"]
                    recency = result["recency"]

                    citation_rows = result.get("citation_comparison") or []
                    reference_rows = result.get("reference_comparison") or []

                    total_references = len(references)
                    total_citations = len(citations)
                    cited_count = sum(1 for row in matching if row.get("Cited"))
                    citations_missing_reference = len(missing)
                    references_missing_citation = max(0, total_references - cited_count)

                    doi_checked = sum(1 for r in reference_rows if r.get("DOI Verified"))
                    doi_suspicious = sum(1 for r in reference_rows if r.get("DOI Suspicious"))
                    doi_suspicious_pct = (
                        doi_suspicious / total_references * 100 if total_references else 0
                    )

                    metric_rows = [
                        {"Metric": "Total References", "Value": total_references},
                        {"Metric": "References > 15",
                         "Value": (f"Yes ({total_references})" if total_references > 15
                                   else f"No ({total_references})")},
                        {"Metric": "Total In-text Citations", "Value": total_citations},
                        {"Metric": "Narrative Citations", "Value": stats.get("narrative", 0)},
                        {"Metric": "Parenthetical Citations", "Value": stats.get("parenthetical", 0)},
                        {"Metric": "% Last 10 Years",
                         "Value": f"{recency.get('recent_percentage', 0):.1f}%"},
                        {"Metric": "Citations Missing from References",
                         "Value": citations_missing_reference},
                        {"Metric": "References Missing from Citations",
                         "Value": references_missing_citation},
                        {"Metric": "DOI Checked (OpenAlex)", "Value": doi_checked},
                        {"Metric": "DOI Suspicious (possible fabrication)",
                         "Value": f"{doi_suspicious} ({doi_suspicious_pct:.1f}%)"},
                    ]
                    metric_df = pd.DataFrame(metric_rows)

                    source_counts = Counter(
                        (row.get("Source Type") or "Other") for row in reference_rows
                    )
                    source_order = CANONICAL_SOURCE_TYPES
                    extra_source_types = sorted(x for x in source_counts if x not in source_order)
                    source_rows = []
                    for source_type in source_order + extra_source_types:
                        count = source_counts.get(source_type, 0)
                        pct = count / total_references * 100 if total_references else 0
                        source_rows.append({
                            "Source Type": source_type,
                            "Count / Percentage": f"{count} ({pct:.1f}%)",
                        })
                    source_df = pd.DataFrame(source_rows)

                    left_col, right_col = st.columns([1, 1], gap="large")
                    with left_col:
                        st.dataframe(metric_df, use_container_width=True, hide_index=True)
                    with right_col:
                        st.dataframe(source_df, use_container_width=True, hide_index=True)

                    show_intext = st.toggle(
                        "Show In-text Citation Correction",
                        value=False,
                        key="show_intext_correction_toggle",
                    )
                    if show_intext:
                        st.markdown("#### In-text Citation Correction")
                        intext_display = pd.DataFrame([
                            {
                                "Type": row.get("Type", ""),
                                "Original Citation": row.get("Original Citation", ""),
                                "Revised Citation": row.get("Revised Citation", ""),
                                "Status": row.get("Status", "MATCH"),
                                "Notes": row.get("Notes", ""),
                            }
                            for row in citation_rows
                        ])
                        if not intext_display.empty:
                            st.dataframe(intext_display, use_container_width=True,
                                         hide_index=True, height=230)
                        else:
                            st.info("No in-text citations were available for automated review.")

                    show_reference = st.toggle(
                        "Show Reference Correction",
                        value=False,
                        key="show_reference_correction_toggle",
                    )
                    if show_reference:
                        st.markdown("#### Reference Correction")
                        reference_display = pd.DataFrame([
                            {
                                "No.": row.get("No."),
                                "Source Type": row.get("Source Type", "Other"),
                                "Year": row.get("Year"),
                                "Original Reference": row.get("Original Reference", ""),
                                "Corrected Version": (
                                    "— WITHHELD (DOI mismatch) —"
                                    if row.get("DOI Suspicious")
                                    else row.get("Corrected Version", "")
                                ),
                                "Placeholders": ", ".join(
                                    k for k, v in (row.get("Placeholders") or {}).items() if v
                                ) or "—",
                                "DOI Checked": "YES" if row.get("DOI Verified") else "NO",
                                "DOI Suspicious": "⚠️ YES" if row.get("DOI Suspicious") else "—",
                                "Verification": row.get("Verification Status", "—"),
                                "Source of Truth": row.get("Source of Truth", "local"),
                                "OpenAlex Title": (row.get("OpenAlex Title") or "")[:60],
                                "OpenAlex Journal": row.get("OpenAlex Journal") or "—",
                                "OpenAlex Vol/Issue/Pages": " / ".join(
                                    str(x) for x in [
                                        row.get("OpenAlex Volume") or "—",
                                        row.get("OpenAlex Issue") or "—",
                                        row.get("OpenAlex Pages") or "—",
                                    ]
                                ),
                                "DOI Issues": row.get("DOI Verification Reasons", ""),
                                "Status": row.get("Status", "MANUAL CHECK"),
                            }
                            for row in reference_rows
                        ])
                        if not reference_display.empty:
                            st.dataframe(reference_display, use_container_width=True,
                                         hide_index=True, height=280)
                        else:
                            st.info("No references were available for automated review.")

                    try:
                        docx_bytes = build_correction_docx(result)
                        safe_name = re.sub(r"[^\w\-]+", "_", result.get("filename", "manuscript"))
                        st.markdown('<div class="apa-green-button-marker"></div>', unsafe_allow_html=True)
                        st.download_button(
                            label="Download Diagnostic Report",
                            data=docx_bytes,
                            file_name=f"{safe_name}_diagnostic_report.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key="download_apa_correction_docx",
                        )
                    except Exception as exc:
                        st.error(f"Could not build DOCX report: {exc}")