# ieee.py
# OmniCite Auditor — IEEE Style Module
# Imported and rendered by app.py via:  import ieee; ieee.render()
#
# Full pipeline:
#   - Bottom-up reference section detection (anchored at the last [n] marker)
#   - Bottom-up reference splitting (anchored at the highest [n] marker)
#   - Staged per-file pipeline process_single_ieee_pdf (Stages A-G)
#   - Pre-classification gating for AI calls
#   - Deterministic placeholder insertion with clean formatting
#   - Wahyudin-style DOCX report (Summary table -> Citations -> References)

import re
import os
import io
import json
import difflib
import traceback
from markitdown import MarkItDown
from datetime import datetime
from collections import Counter

import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from openalex_config import get_openalex_api_key


# =========================================================
# SECRETS
# =========================================================

def _get_openalex_api_key():
    return get_openalex_api_key()


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
    uploaded_file.seek(0)
    pdf_bytes = uploaded_file.getvalue()
    stream = io.BytesIO(pdf_bytes)

    md = MarkItDown(enable_plugins=False)
    result = md.convert_stream(stream, file_extension=".pdf")
    raw_text = result.text_content or ""

    # ---- Detect approximate page count from form-feed / page separators ----
    # MarkItDown usually preserves PDF page breaks as "\f" (form feed) or
    # "\n\n---\n\n" in some versions.
    form_feed_count = raw_text.count("\f")
    approx_pages = max(1, form_feed_count + 1)

    lines = raw_text.splitlines()

    # ---- Count normalized line occurrences ----
    counter = Counter()
    for line in lines:
        stripped = line.strip()
        if not stripped or len(stripped) > 120:
            continue
        norm = re.sub(r"\s+", " ", stripped).lower()
        norm = re.sub(r"\b\d+\b", "<num>", norm)
        counter[norm] += 1

    # ---- Threshold: line must appear on >= 50% of pages ----
    # This is the correct rule for "running headers/footers" because
    # those appear once per page. 50% gives safety margin.
    min_repeats = max(2, int(approx_pages * 0.5))
    repeated = {n for n, c in counter.items() if c >= min_repeats}

    # ---- Boilerplate regex (mirror APA's BOILERPLATE_PATTERNS) ----
    BOILERPLATE_PATTERNS = [
        r"^is licensed under a .*$",
        r"^creative commons.*$",
        r"^cc[- ]by.*$",
        r"^doi:\s*10\.\S+$",
        r"^https?://creativecommons\.org.*$",
        r"^~?\s*\d+\s*~\s*\d+\(\d+\),\s*\d+[-–]\d+\s*$",
        r"^\d+\s*~\s*\d+\(\d+\),\s*\d+[-–]\d+\s*$",
        r"^p?e?ISSN\s*\d+[-–]\d+.*$",
        r"^terakreditasi.*$",
        r"^\*?email koresponden.*$",
        r"^copyright ©.*$",
        r"^received:.*accepted:.*$",
        r"^submitted:.*published:.*$",
        r"^diterima:.*disetujui:.*$",
        r"^preprint\.?\s+under review.*$",
        r"^under review.*$",
    ]
    BOILERPLATE_RE = re.compile("|".join(BOILERPLATE_PATTERNS), re.I)

    cleaned_lines = []
    removed_running_text = []
    for idx, line in enumerate(lines, start=1):
        stripped = line.strip()
        if not stripped:
            cleaned_lines.append(line)
            continue

        # Boilerplate pattern match
        if BOILERPLATE_RE.match(stripped):
            removed_running_text.append({
                "Page": 1, "Text": stripped,
                "Reason": "Boilerplate pattern",
            })
            continue

        # Repeated-line match
        norm = re.sub(r"\s+", " ", stripped).lower()
        norm = re.sub(r"\b\d+\b", "<num>", norm)
        if norm in repeated:
            removed_running_text.append({
                "Page": 1, "Text": stripped,
                "Reason": "Repeated header/footer line",
            })
            continue

        # Lone page number
        if re.fullmatch(r"\s*\d{1,4}\s*", stripped):
            removed_running_text.append({
                "Page": 1, "Text": stripped,
                "Reason": "Page number",
            })
            continue

        cleaned_lines.append(line)

    cleaned = "\n".join(cleaned_lines)
    full_text = f"<<<PAGE_BREAK:1>>>\n{cleaned}"
    pages = [{"page": 1, "text": cleaned}]

    return full_text, pages, removed_running_text


# =========================================================
# PDF FONT / ITALIC EXTRACTION
# =========================================================

def extract_pdf_style_spans(uploaded_file):
    """
    MarkItDown does not expose per-span font flags, so italic detection is
    not available. Return an empty list — the DOCX builder will fall back
    to the AI-supplied italic_elements.
    """
    uploaded_file.seek(0)
    return []


def normalize_style_text(text):
    text = (text or "").replace("\u00ad", "")
    text = text.replace("‐", "-").replace("‑", "-").replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


# =========================================================
# TEXT CLEANING
# =========================================================

def clean_text(text):
    text = text.replace("\u00ad", "")
    text = text.replace("\u2013", "-").replace("\u2014", "-")
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
# REFERENCE SECTION DETECTION — bottom-up anchor
# =========================================================

REFERENCE_HEADINGS = {
    "references", "reference", "reference list", "reference section",
    "bibliography", "works cited", "literature cited",
    "daftar pustaka", "daftar rujukan", "rujukan",
}

POST_REFERENCE_HEADINGS = {
    "acknowledgement", "acknowledgements", "acknowledgment", "acknowledgments",
    "author contribution", "author contributions", "authors contribution",
    "authors contributions", "contribution", "contributions",
    "author profile", "author profiles", "profile", "profiles",
    "biography", "biographies",
    "conflict of interest", "conflicts of interest",
    "competing interest", "competing interests",
    "declaration", "declarations",
    "funding", "funding information", "funding statement",
    "data availability", "data availability statement",
    "ethical approval", "ethics approval", "ethics statement",
    "informed consent", "consent for publication",
    "disclosure", "disclosures",
    "appendix", "appendices",
    "supplementary material", "supplementary materials",
    "notes", "author note", "author notes",
}


def normalize_heading(text):
    text = text.strip().lower()
    text = re.sub(r"^\s*(?:\d+(?:\.\d+)*)[\.\s:-]+", "", text)
    text = re.sub(r"^[\*\#•\-\s]+|[\*\#•\-\s]+$", "", text)
    text = re.sub(r"\s+", " ", text).rstrip(":").strip()
    return text


def _heading_matches_reference_vocab(line):
    stripped = line.strip()
    if not stripped:
        return False
    if "\t" in stripped or re.search(r"\s{3,}\S+\s{3,}\S+\s{3,}", stripped):
        return False
    if len(stripped) > 60:
        return False
    norm = normalize_heading(stripped)
    if not norm:
        return False
    if norm in REFERENCE_HEADINGS:
        return True
    for h in REFERENCE_HEADINGS:
        if norm == h or norm.startswith(h + " "):
            return True
    return False


def is_post_reference_heading(line):
    normalized = normalize_heading(line)
    if not normalized:
        return False
    if normalized in POST_REFERENCE_HEADINGS:
        return True
    for h in POST_REFERENCE_HEADINGS:
        if normalized.startswith(h + " "):
            return True
    return False


def find_reference_section(text):
    """
    Detect the bibliography without assuming every [n] marker survived PDF
    extraction. If a References heading exists, use it. If the heading was
    lost, anchor on the EARLIEST surviving reference marker near the end of
    the manuscript and scan backwards across bibliography-looking lines.

    The previous fallback started at the LAST surviving marker. That can
    silently discard most of the bibliography when PDF extraction removes
    the heading or early reference numbers.
    """
    lines = text.splitlines()
    marker_re = re.compile(r"^\s*\[?\s*(\d{1,3})\s*\]")

    marker_indices = []
    for i, line in enumerate(lines):
        m = marker_re.match(line)
        if not m:
            continue
        try:
            n = int(m.group(1))
        except ValueError:
            continue
        if 1 <= n <= 500 and not (1900 <= n <= 2099):
            marker_indices.append((i, n))

    # First preference: an explicit References/Bibliography heading.
    if marker_indices:
        last_marker_idx = marker_indices[-1][0]
        heading_idx = None
        for i in range(last_marker_idx, -1, -1):
            if _heading_matches_reference_vocab(lines[i]):
                heading_idx = i
                break

        if heading_idx is not None:
            end_idx = None
            saw_reference_content = False
            for i in range(heading_idx + 1, len(lines)):
                line = lines[i].strip()
                if not line:
                    continue
                if re.fullmatch(r"<<<PAGE_BREAK:\d+>>>", line):
                    continue
                if marker_re.match(line) or re.search(
                    r"\b(?:19|20)\d{2}\b|10\.\d{4,9}/|\bvol\.\s*\d+|\bpp\.\s*\d+",
                    line,
                    re.I,
                ):
                    saw_reference_content = True
                if saw_reference_content and is_post_reference_heading(line):
                    end_idx = i
                    break

            heading_found = lines[heading_idx].strip()
            body_text = "\n".join(lines[:heading_idx])
            if end_idx is not None:
                reference_text = "\n".join(lines[heading_idx + 1:end_idx])
            else:
                reference_text = "\n".join(lines[heading_idx + 1:])
            return reference_text, heading_found, body_text

    # If no marker survived at all, fall back to the normal heading search.
    if not marker_indices:
        return _find_reference_section_topdown(text)

    # ---------------------------------------------------------
    # Heading missing: DO NOT start from the last marker.
    # ---------------------------------------------------------
    # Use the earliest surviving marker. MarkItDown often drops [1]...[10]
    # while leaving [11] and later markers, so we scan backwards to recover
    # the unnumbered bibliography entries before that marker.
    first_marker_idx = marker_indices[0][0]

    def bibliography_signal(line):
        line = (line or "").strip()
        if not line:
            return False
        return bool(re.search(
            r"(?:"
            r"10\.\d{4,9}/"
            r"|\b(?:19|20)\d{2}\b"
            r"|\bvol\.\s*\d+"
            r"|\bno\.\s*\d+"
            r"|\bpp?\.\s*\d+"
            r"|\bIEEE\b"
            r"|\bProceedings\b"
            r"|\bJournal\b"
            r")",
            line,
            re.I,
        ))

    start_idx = first_marker_idx
    signals_seen = 0
    blank_run = 0

    # Scan at most 180 extracted lines backwards. This is deliberately
    # bounded so ordinary body text is not swallowed into References.
    lower_bound = max(0, first_marker_idx - 180)
    for i in range(first_marker_idx - 1, lower_bound - 1, -1):
        stripped = lines[i].strip()

        if re.fullmatch(r"<<<PAGE_BREAK:\d+>>>", stripped):
            continue

        if not stripped:
            blank_run += 1
            # Once we have already seen bibliography evidence, a large blank
            # gap is a good boundary between body and bibliography.
            if signals_seen >= 2 and blank_run >= 3:
                break
            continue

        blank_run = 0

        if _heading_matches_reference_vocab(stripped):
            start_idx = i + 1
            break

        if bibliography_signal(stripped):
            signals_seen += 1

        # Keep continuation/author lines while scanning backwards.
        start_idx = i

    # Stop at a post-reference section if one exists after the markers.
    end_idx = len(lines)
    saw_ref = False
    for i in range(start_idx, len(lines)):
        stripped = lines[i].strip()
        if marker_re.match(stripped) or bibliography_signal(stripped):
            saw_ref = True
        if saw_ref and is_post_reference_heading(stripped):
            end_idx = i
            break

    reference_text = "\n".join(lines[start_idx:end_idx])
    body_text = "\n".join(lines[:start_idx])
    return reference_text, None, body_text

def _find_reference_section_topdown(text):
    lines = text.splitlines()
    start_index = None
    for i, line in enumerate(lines):
        if _heading_matches_reference_vocab(line):
            start_index = i
            break
    if start_index is None:
        return None, None, text

    end_index = None
    for i in range(start_index + 1, len(lines)):
        if is_post_reference_heading(lines[i]):
            end_index = i
            break

    body_text = "\n".join(lines[:start_index])
    if end_index is not None:
        reference_text = "\n".join(lines[start_index + 1:end_index])
    else:
        reference_text = "\n".join(lines[start_index + 1:])
    return reference_text, lines[start_index].strip(), body_text


# =========================================================
# IEEE REFERENCE SPLITTING — robust marker + recovery logic
# =========================================================

_MARKER_TOKEN_RE = re.compile(
    r"(?:"
    r"(?:(?<=\s)|^)\[\s*(\d{1,3})\s*\]"          # [12]
    r"|"
    r"(?:(?<=\s)|^)(\d{1,3})\.\s+(?=[A-Z])"     # 12. Author...
    r")"
    r"(?=\s|[A-Z]|$)"
)


def _find_all_markers(text):
    """Return every plausible IEEE reference marker as (position, number)."""
    hits = []
    for m in _MARKER_TOKEN_RE.finditer(text or ""):
        num_str = m.group(1) or m.group(2)
        if not num_str:
            continue
        try:
            num = int(num_str)
        except (TypeError, ValueError):
            continue
        if 1900 <= num <= 2099:
            continue
        if not (1 <= num <= 500):
            continue
        hits.append((m.start(), num))
    return hits


def estimate_expected_reference_count(reference_text, full_text=None):
    """
    Estimate the expected IEEE bibliography size from the highest bracketed
    reference number found in the reference section and full manuscript.
    """
    candidates = []
    for src in (reference_text or "", full_text or ""):
        for m in re.finditer(r"\[\s*(\d{1,3})\s*\]", src):
            try:
                num = int(m.group(1))
            except (TypeError, ValueError):
                continue
            if 1 <= num <= 500:
                candidates.append(num)
    return max(candidates) if candidates else None


def _normalize_reference_text(text):
    """Normalize PDF artifacts while preserving bibliography boundaries."""
    if not text:
        return ""
    text = re.sub(r"<<<PAGE_BREAK:\d+>>>", " ", text)
    text = (
        text.replace("\u00ad", "")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
        .replace("\u00a0", " ")
    )
    # Repair cases such as "author.[12]" or "2024.[13]".
    text = re.sub(r"(?<=[A-Za-z0-9.,;)])\[(\d{1,3})\]", r" [\1]", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _split_by_year_author_boundaries(text):
    """
    Conservative recovery splitter for chunks where one or more [n] markers
    disappeared during PDF extraction. A cut is accepted only when a likely
    reference terminator is followed by a likely new author block.
    """
    text = (text or "").strip()
    if not text:
        return []

    author_start = (
        r"(?:"
        r"[A-Z](?:[-.]?[A-Z])?\."
        r"(?:\s*[A-Z](?:[-.]?[A-Z])?\.)*"
        r"\s+[A-Z][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+"
        r"|"
        r"[A-Z][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+,\s*[A-Z]\."
        r")"
    )

    boundary_re = re.compile(
        r"("
        r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+\.?"
        r"|"
        r"(?:19|20)\d{2}\."
        r")"
        r"\s+(?=" + author_start + r")",
        re.I,
    )

    cuts = [m.end() for m in boundary_re.finditer(text)]
    if not cuts:
        return [text]

    refs = []
    start = 0
    for cut in cuts:
        chunk = text[start:cut].strip()
        if chunk:
            refs.append(chunk)
        start = cut
    tail = text[start:].strip()
    if tail:
        refs.append(tail)
    return refs


def _split_by_year_author_boundaries_aggressive(text):
    """Looser last-resort splitter used only when extraction remains short."""
    text = (text or "").strip()
    if not text:
        return []

    author_start = (
        r"(?:"
        r"[A-Z](?:[-.]?[A-Z])?\."
        r"(?:\s*[A-Z](?:[-.]?[A-Z])?\.)*"
        r"\s+[A-Z][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+"
        r"|"
        r"[A-Z][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+,\s*[A-Z]\."
        r")"
    )

    boundary_re = re.compile(
        r"("
        r"(?:19|20)\d{2}\.?"
        r"|10\.\d{4,9}/[-._;()/:A-Za-z0-9]+\.?"
        r"|https?://\S+?\."
        r"|pp?\.\s*\d+(?:\s*[-–]\s*\d+)?\."
        r")\s+(?=" + author_start + r")",
        re.I,
    )

    cuts = [m.end() for m in boundary_re.finditer(text)]
    if not cuts:
        return [text]

    parts = []
    last = 0
    for cut in cuts:
        chunk = text[last:cut].strip()
        if chunk:
            parts.append(chunk)
        last = cut
    tail = text[last:].strip()
    if tail:
        parts.append(tail)

    merged = []
    for part in parts:
        if len(part) < 40 and merged:
            merged[-1] = merged[-1] + " " + part
        else:
            merged.append(part)
    return merged if len(merged) > 1 else [text]


def _split_chunk_using_author_boundaries(chunk, aggressive=False):
    if aggressive:
        return _split_by_year_author_boundaries_aggressive(chunk)
    return _split_by_year_author_boundaries(chunk)


def _uninterpretable_reference_placeholder(number):
    return f"[UNINTERPRETABLE_REFERENCE_{int(number)}]"


def _is_uninterpretable_reference(ref):
    return bool(re.fullmatch(r"\[UNINTERPRETABLE_REFERENCE_\d+\]", (ref or "").strip()))


def split_references(reference_text, expected_count=None):
    """Split an IEEE bibliography while preserving its original numbering.

    The key rule is positional integrity: when the manuscript indicates that
    reference [n] exists, slot n must never disappear merely because its text is
    malformed, authorless, or difficult to parse.  If a numbered entry cannot
    be recovered, a deterministic placeholder is inserted for that number.
    This keeps citation/reference numbering stable all the way to the DOCX.
    """
    if not reference_text:
        return []

    text = _normalize_reference_text(reference_text)
    if not text:
        return []

    markers = sorted(set(_find_all_markers(text)), key=lambda x: x[0])

    # Without surviving markers we cannot reliably assign original numbers.
    # Fall back to boundary splitting, then pad to expected_count if known.
    if not markers:
        refs = [re.sub(r"\s+", " ", r).strip()
                for r in _split_by_year_author_boundaries_aggressive(text)
                if r.strip()]
        if expected_count:
            refs = refs[:expected_count]
            while len(refs) < expected_count:
                refs.append(_uninterpretable_reference_placeholder(len(refs) + 1))
        return refs

    highest_marker = max(n for _, n in markers)
    target_count = int(expected_count or highest_marker)
    target_count = max(target_count, highest_marker)
    slots = {n: None for n in range(1, target_count + 1)}

    # Text before the first surviving marker usually represents early entries
    # whose [1], [2], ... markers were lost by PDF extraction.
    first_pos, first_num = markers[0]
    prefix = text[:first_pos].strip()
    if prefix and first_num > 1:
        pieces = _split_by_year_author_boundaries_aggressive(prefix)
        pieces = [re.sub(r"\s+", " ", x).strip() for x in pieces if x.strip()]
        # Align recovered prefix entries to the slots immediately preceding the
        # first surviving marker. If fewer pieces are recovered, the remaining
        # slots deliberately stay empty and become Cannot Interpret placeholders.
        start_num = max(1, first_num - len(pieces))
        for num, piece in zip(range(start_num, first_num), pieces[-(first_num-start_num):]):
            slots[num] = piece

    # Every surviving marker owns its slot. If markers are missing between n
    # and the next surviving marker, attempt to split the chunk into n..next-1.
    for i, (pos, marker_num) in enumerate(markers):
        next_pos = markers[i + 1][0] if i + 1 < len(markers) else len(text)
        next_num = markers[i + 1][1] if i + 1 < len(markers) else target_count + 1
        chunk = text[pos:next_pos].strip()
        chunk = re.sub(r"^\s*(?:\[\s*\d{1,3}\s*\]|\d{1,3}\.)\s*", "", chunk).strip()
        if not chunk:
            continue

        gap = max(1, next_num - marker_num)
        pieces = [chunk]
        if gap > 1:
            recovered = _split_chunk_using_author_boundaries(chunk, aggressive=False)
            if len(recovered) < min(gap, 2):
                recovered = _split_chunk_using_author_boundaries(chunk, aggressive=True)
            if len(recovered) > 1:
                pieces = recovered

        for offset, piece in enumerate(pieces[:gap]):
            num = marker_num + offset
            if 1 <= num <= target_count:
                cleaned = re.sub(r"\s+", " ", piece).strip()
                if cleaned:
                    slots[num] = cleaned

    # Never collapse numbering. Missing/unreadable slots remain represented.
    return [slots[n] if slots[n] and len(slots[n]) >= 5
            else _uninterpretable_reference_placeholder(n)
            for n in range(1, target_count + 1)]


def recover_glued_references(references, expected_count):
    """Additional safety pass for unusually long glued bibliography chunks."""
    if not references:
        return references

    recovered = []
    for ref in references:
        if len(ref) <= 300:
            recovered.append(ref)
            continue
        parts = _split_by_year_author_boundaries(ref)
        recovered.extend(parts if len(parts) > 1 else [ref])

    if expected_count and len(recovered) < expected_count * 0.90:
        aggressive = []
        for ref in recovered:
            parts = _split_by_year_author_boundaries_aggressive(ref)
            aggressive.extend(parts if len(parts) > 1 else [ref])
        if len(aggressive) > len(recovered):
            recovered = aggressive

    return recovered

# =========================================================
# IEEE REFERENCE PARSER
# =========================================================

_IEEE_YEAR_RE = re.compile(r"\b((?:19|20)\d{2})\b")
_IEEE_DOI_RE = re.compile(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", re.I)
_IEEE_URL_RE = re.compile(r"https?://\S+", re.I)
_ET_AL_RE = re.compile(r"\bet\s+al\.?", re.I)


def parse_ieee_authors(author_block):
    if not author_block:
        return []
    block = author_block.strip().rstrip(".,").strip()
    block = re.sub(r"^\s*and\s+", "", block, flags=re.I)
    block = re.sub(r",\s*and\s+", ", ", block, flags=re.I)
    parts = [p.strip() for p in block.split(",") if p.strip()]
    surnames = []
    for p in parts:
        p = re.sub(r"^(?:and|&)\s+", "", p, flags=re.I).strip()
        if not p:
            continue
        tokens = [t for t in re.split(r"\s+", p) if t]
        if not tokens:
            continue
        surnames.append(tokens[-1].rstrip(".,;"))
    return surnames


# =========================================================
# AUTHOR NAME NORMALIZATION
# =========================================================

_AUTHOR_PARTICLES = {
    "van", "von", "de", "del", "della", "der", "den", "dos", "da",
    "di", "la", "le", "los", "las", "bin", "binti", "bte", "al",
    "el", "ter", "ten", "op", "in", "'t",
}


def _split_ieee_author_entries(author_block):
    if not author_block:
        return [], []

    block = author_block.strip().rstrip(".,").strip()
    pattern = re.compile(
        r"\s*(,\s*and\s+|,\s*&\s+|\s+and\s+|\s+&\s+|,\s*|$)",
        flags=re.I,
    )

    entries = []
    separators = []
    pos = 0
    for m in pattern.finditer(block):
        chunk = block[pos:m.start()].strip()
        sep = m.group(1).strip()
        if chunk:
            chunk = re.sub(r"^(?:and|&)\s+", "", chunk, flags=re.I).strip()
            entries.append(chunk)
            separators.append(sep)
        pos = m.end()
        if pos >= len(block):
            break

    if len(separators) < len(entries):
        separators.append("")
    return entries, separators


def _is_ieee_initial_form(name):
    if not name:
        return False
    name = name.strip().rstrip(".,")
    m = re.match(r"^([A-Z](?:\.\s*[A-Z])*\.?)\s+(\S.*)$", name)
    return bool(m)


def _is_full_name_form(name):
    if not name:
        return False
    name = name.strip().rstrip(".,")
    tokens = name.split()
    if len(tokens) < 2:
        return False
    first = tokens[0]
    if len(first) <= 1:
        return False
    if not re.match(r"^[A-Z][a-zA-Z\-']+$", first):
        return False
    return True


def _to_ieee_initial_form(full_name):
    if not full_name:
        return full_name
    raw = full_name.strip().rstrip(",.").strip()
    if not raw:
        return raw
    tokens = raw.split()

    surname_start = len(tokens) - 1
    while surname_start - 1 >= 1 and tokens[surname_start - 1].lower() in _AUTHOR_PARTICLES:
        surname_start -= 1

    given = tokens[:surname_start]
    surname = tokens[surname_start:]

    if not given:
        return " ".join(surname)

    initials = []
    for g in given:
        m = re.match(r"^([A-Za-zÀ-ÖØ-öø-ÿ])", g)
        if m:
            initials.append(m.group(1).upper() + ".")
        elif g:
            initials.append(g[0].upper() + ".")

    return " ".join(initials + surname)


def normalize_ieee_authors(author_block):
    entries, _ = _split_ieee_author_entries(author_block)
    if not entries:
        return author_block, False, []

    notes = []
    new_entries = []

    for idx, entry in enumerate(entries):
        cleaned = re.sub(r"^(?:and|&)\s+", "", entry, flags=re.I).strip()

        if _is_ieee_initial_form(cleaned):
            new_entries.append(cleaned)
            continue

        if _is_full_name_form(cleaned):
            converted = _to_ieee_initial_form(cleaned)
            if converted != cleaned:
                notes.append(
                    f"Author {idx + 1} converted from '{cleaned}' to '{converted}'"
                )
            new_entries.append(converted)
            continue

        new_entries.append(cleaned)

    if len(new_entries) == 1:
        rebuilt = new_entries[0]
    elif len(new_entries) == 2:
        rebuilt = f"{new_entries[0]} and {new_entries[1]}"
    else:
        rebuilt = ", ".join(new_entries[:-1]) + ", and " + new_entries[-1]

    changed = rebuilt.strip().rstrip(".,") != author_block.strip().rstrip(".,")
    return rebuilt, changed, notes


# =========================================================
# IEEE REFERENCE PARSER (main)
# =========================================================

def parse_ieee_reference(reference):
    result = {
        "raw": reference,
        "year": None,
        "first_author": None,
        "authors": [],
        "doi": None,
        "url": None,
        "title": None,
        "venue": None,
        "volume": None,
        "issue": None,
        "pages": None,
        "source_type": "Other",
    }

    ym = _IEEE_YEAR_RE.search(reference)
    if ym:
        result["year"] = ym.group(1)

    dm = _IEEE_DOI_RE.search(reference)
    if dm:
        result["doi"] = "10." + dm.group(0).split("10.", 1)[-1].rstrip(".,;)")
    um = _IEEE_URL_RE.search(reference)
    if um:
        result["url"] = um.group(0).rstrip(".,;)")

    title_match = re.search(r"[\"“](.+?)[\"”]", reference)
    if title_match:
        result["title"] = title_match.group(1).strip().rstrip(",")
        after_title = reference[title_match.end():].strip()
    else:
        after_title = reference

    venue_match = re.search(
        r",\s*(?P<venue>[A-Z][^,]*?),\s*vol\.\s*(?P<vol>\d+)",
        after_title,
    )
    if venue_match:
        result["venue"] = venue_match.group("venue").strip().strip(",")
        result["volume"] = venue_match.group("vol")

    vol_match = re.search(r"\bvol\.\s*(\d+)", reference)
    if vol_match and not result["volume"]:
        result["volume"] = vol_match.group(1)

    no_match = re.search(r"\bno\.\s*([\w\-]+)", reference)
    if no_match:
        result["issue"] = no_match.group(1)

    pp_match = re.search(r"\bpp\.\s*([\w\-]+(?:\s*[-–—]\s*[\w\-]+)?)", reference)
    if pp_match:
        result["pages"] = pp_match.group(1).strip()

    cut = len(reference)
    if title_match:
        cut = min(cut, title_match.start())
    if ym:
        cut = min(cut, ym.start())
    author_block = reference[:cut].strip().rstrip(",.")
    authors = parse_ieee_authors(author_block)
    result["authors"] = authors
    if authors:
        result["first_author"] = authors[0]

    low = reference.lower()

    # Conference/proceedings evidence takes precedence over DOI.
    conference_pattern = (
        r"\b(?:proc\.?|proceedings|conference|symposium|workshop|"
        r"congress|congres|colloquium|convention|seminar|procedia|meeting|"
        r"international\s+conference)\b"
    )
    journal_pattern = (
        r"\b(?:journal|jurnal|review|quarterly|bulletin|transactions|letters|magazine)\b"
    )
    book_pattern = r"\b(?:book|monograph|handbook|textbook)\b"
    report_pattern = (
        r"\b(?:tech\.?\s*rep\.?|technical\s+report|report|working\s+paper|"
        r"white\s+paper|standard)\b"
    )
    chapter_pattern = r"\b(?:ed\.|eds\.|chapter)\b"

    if re.search(conference_pattern, low):
        result["source_type"] = "Conference Paper"
    elif re.search(chapter_pattern, low):
        result["source_type"] = "Book Chapter"
    elif re.search(book_pattern, low):
        result["source_type"] = "Book"
    elif re.search(report_pattern, low):
        result["source_type"] = "Technical Report"
    elif re.search(journal_pattern, low) or (result["volume"] and result["pages"] and result["venue"]):
        result["source_type"] = "Journal Article"
    elif re.search(r"https?://", low) and not result["doi"]:
        result["source_type"] = "Web Page"
    else:
        # DOI alone does NOT establish a journal article.
        result["source_type"] = "Other"

    return result


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
    """Fetch rich OpenAlex metadata for a DOI. None means unavailable; _not_found means DOI did not resolve."""
    if not doi:
        return None
    try:
        import requests
    except Exception:
        return None

    clean = _normalize_doi_for_lookup(doi)
    if not clean:
        return None

    api_key = _get_openalex_api_key()
    url = f"https://api.openalex.org/works/doi:{clean}"
    params = {"api_key": api_key} if api_key else {}
    try:
        r = requests.get(url, params=params, timeout=10)
        if r.status_code == 404:
            return {"_not_found": True, "doi": clean}
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
        "doi": clean,
        "title": data.get("title") or data.get("display_name") or "",
        "authors": authors,
        "year": data.get("publication_year"),
        "work_type": data.get("type") or "",
        "crossref_type": data.get("type_crossref") or "",
        "venue": source.get("display_name") or "",
        "source_type": source.get("type") or "",
        "volume": biblio.get("volume"),
        "issue": biblio.get("issue"),
        "first_page": biblio.get("first_page"),
        "last_page": biblio.get("last_page"),
        "is_oa": (data.get("open_access") or {}).get("is_oa"),
        "raw": data,
    }


def _openalex_ieee_source_type(meta, original_reference=""):
    """Map OpenAlex metadata to OmniCite source types; conference evidence wins."""
    if not meta or meta.get("_not_found"):
        return "Other"
    wt = str(meta.get("work_type") or "").lower()
    ct = str(meta.get("crossref_type") or "").lower()
    st = str(meta.get("source_type") or "").lower()
    venue = str(meta.get("venue") or "")
    hay = " ".join([wt, ct, st, venue, original_reference or ""]).lower()

    if any(x in hay for x in (
        "proceedings", "conference", "symposium", "workshop", "congress",
        "colloquium", "convention", "seminar", "procedia", "meeting"
    )):
        return "Conference Paper"
    if "book-chapter" in hay or "book chapter" in hay or "chapter" in wt:
        return "Book Chapter"
    if any(x in hay for x in ("book", "monograph")) and "article" not in wt:
        return "Book"
    if any(x in hay for x in ("report", "report-series")):
        return "Technical Report"
    if any(x in hay for x in ("journal", "article")):
        return "Journal Article"
    return "Other"


def _openalex_author_to_ieee(name):
    """Convert an OpenAlex display name to IEEE initials-first form."""
    if not name:
        return ""
    name = re.sub(r"\s+", " ", str(name)).strip()
    if not name:
        return ""
    if "," in name:
        surname, given = [x.strip() for x in name.split(",", 1)]
        name = (given + " " + surname).strip()
    return _to_ieee_initial_form(name)


def _join_ieee_authors(names):
    authors = [_openalex_author_to_ieee(n) for n in (names or [])]
    authors = [a for a in authors if a]
    if not authors:
        return ""
    if len(authors) == 1:
        return authors[0]
    if len(authors) == 2:
        return f"{authors[0]} and {authors[1]}"
    return ", ".join(authors[:-1]) + ", and " + authors[-1]


def build_ieee_reference_from_openalex(original_reference, meta):
    """Construct a corrected IEEE reference from verified OpenAlex DOI metadata.

    OpenAlex supplies bibliographic facts; Python supplies IEEE punctuation and
    formatting. Missing OpenAlex fields are filled only from the parsed original
    reference when available. Nothing is invented.
    """
    if not meta or meta.get("_not_found"):
        return None

    original = strip_markdown_markers(clean_text(original_reference or ""))
    op = parse_ieee_reference(original)
    source_type = _openalex_ieee_source_type(meta, original)
    if source_type == "Other":
        source_type = normalize_source_type(op.get("source_type"))

    authors = _join_ieee_authors(meta.get("authors"))
    title = (meta.get("title") or op.get("title") or "").strip().rstrip(".")
    venue = (meta.get("venue") or op.get("venue") or "").strip().rstrip(".,")
    year = meta.get("year") or op.get("year")
    volume = meta.get("volume") or op.get("volume")
    issue = meta.get("issue") or op.get("issue")
    first_page = meta.get("first_page")
    last_page = meta.get("last_page")
    pages = None
    # OpenAlex biblio occasionally returns malformed page fields.  Accept only
    # page-like values and never concatenate a publication year onto pages.
    def _page_value(v):
        if v is None:
            return None
        x = str(v).strip()
        return x if re.fullmatch(r"[A-Za-z]?\d+(?:[-–][A-Za-z]?\d+)?", x) else None

    fp = _page_value(first_page)
    lp = _page_value(last_page)
    if fp and lp:
        # If first_page already contains a range, do not append last_page.
        if re.search(r"[-–]", fp):
            pages = fp
        elif fp == lp:
            pages = fp
        elif str(lp) == str(year):
            pages = fp
        else:
            pages = f"{fp}-{lp}"
    elif fp:
        pages = fp
    else:
        pages = op.get("pages")
    doi = meta.get("doi") or op.get("doi")

    parts = []
    if authors:
        parts.append(authors)

    if source_type in {"Journal Article", "Conference Paper"} and title:
        parts.append(f'"{title}"')
    elif title:
        parts.append(title)

    if source_type == "Conference Paper":
        if venue:
            parts.append(f"in {venue}")
        if pages:
            parts.append(f"pp. {pages}")
        if year:
            parts.append(str(year))
    elif source_type == "Journal Article":
        if venue:
            parts.append(venue)
        if volume:
            parts.append(f"vol. {volume}")
        if issue:
            parts.append(f"no. {issue}")
        if pages:
            parts.append(f"pp. {pages}")
        if year:
            parts.append(str(year))
    else:
        if venue:
            parts.append(venue)
        if volume:
            parts.append(f"vol. {volume}")
        if pages:
            parts.append(f"pp. {pages}")
        if year:
            parts.append(str(year))

    corrected = ", ".join(p for p in parts if p)
    if doi:
        corrected += (", " if corrected else "") + f"doi: {doi}"
    corrected = re.sub(r"\s+", " ", corrected).strip().rstrip(".") + "."
    italic_tokens = compute_ieee_italic_tokens(source_type, corrected, parse_ieee_reference(corrected))
    if venue and source_type in {"Journal Article", "Conference Paper"} and venue in corrected and venue not in italic_tokens:
        italic_tokens.insert(0, venue)

    return {
        "Corrected": corrected,
        "Source Type": source_type,
        "Year": int(year) if str(year).isdigit() else None,
        "Italic Tokens": italic_tokens,
        "Metadata Source": "OpenAlex",
    }

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


def verify_reference_against_openalex(reference, parsed):
    result = {
        "checked": False,
        "doi": parsed.get("doi"),
        "title_similarity": None,
        "author_overlap": None,
        "suspicious": False,
        "reasons": [],
        "crossref_title": None,
        "crossref_authors": [],
        "openalex_metadata": None,
    }

    doi = parsed.get("doi")
    if not doi:
        return result

    meta = _fetch_openalex_metadata(doi)
    if meta is None:
        return result

    if meta.get("_not_found"):
        result["checked"] = True
        result["suspicious"] = True
        result["reasons"].append("DOI does not resolve in OpenAlex (possible fake DOI).")
        return result

    result["checked"] = True
    result["openalex_metadata"] = meta
    result["crossref_title"] = meta.get("title")
    result["crossref_authors"] = meta.get("authors", [])

    ref_title = parsed.get("title") or ""
    oa_title = meta.get("title") or ""
    if ref_title and oa_title:
        sim = _title_similarity(ref_title, oa_title)
        result["title_similarity"] = sim
        if sim < 0.60:
            result["suspicious"] = True
            result["reasons"].append(
                f"DOI resolves to a different title (similarity {sim:.0%})."
            )

    ref_surnames = {s.lower() for s in (parsed.get("authors") or []) if s}
    oa_surnames = {_surname_from_full_name(n) for n in meta.get("authors", []) if n}
    oa_surnames.discard("")
    if oa_surnames:
        overlap = len(ref_surnames & oa_surnames) / max(1, len(oa_surnames))
        result["author_overlap"] = overlap
        if overlap < 0.50:
            result["suspicious"] = True
            result["reasons"].append(
                f"DOI author list does not match the manuscript "
                f"(only {overlap:.0%} of DOI surnames found)."
            )

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
    title_match = re.search(r"[\"“].+?[\"”]", reference)
    year_match = _IEEE_YEAR_RE.search(reference)
    cut = len(reference)
    if title_match:
        cut = min(cut, title_match.start())
    if year_match:
        cut = min(cut, year_match.start())
    block = reference[:cut].strip().rstrip(",.")
    if not block:
        return 0
    entries, _ = _split_ieee_author_entries(block)
    return len(entries)


# =========================================================
# IEEE IN-TEXT CITATION EXTRACTION
# =========================================================

def collapse_citation_cluster(numbers):
    nums = sorted({int(n) for n in numbers})
    if not nums:
        return ""
    parts = []
    i = 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        run_len = j - i + 1
        if run_len >= 3:
            parts.append(f"[{nums[i]}]-[{nums[j]}]")
        elif run_len == 2:
            parts.append(f"[{nums[i]}]")
            parts.append(f"[{nums[j]}]")
        else:
            parts.append(f"[{nums[i]}]")
        i = j + 1
    return ", ".join(parts)


def _page_for_offset(full_text, offset):
    if offset is None or offset < 0:
        return None
    prefix = full_text[:offset]
    matches = list(re.finditer(r"<<<PAGE_BREAK:(\d+)>>>", prefix))
    if not matches:
        return None
    return int(matches[-1].group(1))


def extract_ieee_citations(text):
    citations = []
    seen_positions = set()
    clusters = []

    cluster_re = re.compile(
        r"\[\s*(\d+)\s*\]"
        r"(?:\s*[–—-]\s*\[\s*(\d+)\s*\])?"
        r"(?:\s*[,;]\s*\[\s*(\d+)\s*\]"
        r"(?:\s*[–—-]\s*\[\s*(\d+)\s*\])?)*"
    )
    single_re = re.compile(r"\[\s*(\d+)\s*\]")
    range_re = re.compile(r"\[\s*(\d+)\s*\]\s*[–—-]\s*\[\s*(\d+)\s*\]")

    for m in cluster_re.finditer(text):
        chunk = m.group(0)
        numbers = set()

        for rng in range_re.finditer(chunk):
            a, b = int(rng.group(1)), int(rng.group(2))
            if a <= b and b - a <= 50:
                numbers.update(range(a, b + 1))

        covered = set()
        for rng in range_re.finditer(chunk):
            a, b = int(rng.group(1)), int(rng.group(2))
            covered.update(range(a, b + 1))

        for sm in single_re.finditer(chunk):
            n = int(sm.group(1))
            if n not in covered:
                numbers.add(n)

        if not numbers:
            continue

        for n in sorted(numbers):
            key = (m.start(), n)
            if key in seen_positions:
                continue
            citations.append({
                "number": n,
                "raw": f"[{n}]",
                "type": "bracketed",
                "context": text[max(0, m.start() - 40): m.end() + 40],
                "offset": m.start(),
            })
            seen_positions.add(key)

        clusters.append({
            "numbers": sorted(numbers),
            "raw": chunk.strip(),
            "canonical": collapse_citation_cluster(numbers),
            "start": m.start(),
            "end": m.end(),
        })

    return citations, clusters


# =========================================================
# IEEE CITATION / CLUSTER QUALITY CHECKS
# =========================================================

def _cluster_out_of_order(numbers):
    return numbers != sorted(numbers)


def _cluster_needs_range_collapse(cluster):
    nums = sorted({int(n) for n in cluster["numbers"]})
    if len(nums) < 3:
        return False
    raw_norm = cluster["raw"].replace(" ", "").replace("–", "-").replace("—", "-")
    for i in range(len(nums) - 2):
        if nums[i + 1] == nums[i] + 1 and nums[i + 2] == nums[i] + 2:
            already = f"[{nums[i]}]-[{nums[i+2]}]" in raw_norm
            if not already:
                return True
    return False


def evaluate_cluster(cluster):
    reasons = []
    if _cluster_out_of_order(cluster["numbers"]):
        reasons.append("Citation numbers are not in ascending order.")
    if _cluster_needs_range_collapse(cluster):
        reasons.append("Consecutive numbers [n], [n+1], [n+2] must use the range form [n]-[n+2].")
    if cluster["raw"].strip() != cluster["canonical"].strip() and not reasons:
        reasons.append("Citation cluster does not match IEEE canonical form.")
    return (bool(reasons), reasons)


def enforce_ieee_citation_rules(cluster):
    return collapse_citation_cluster(cluster["numbers"])


# =========================================================
# REFERENCE STRUCTURE VALIDATION
# =========================================================

_SINGLE_PAGE_RE = re.compile(r"^\s*(\d+)\s*$")


def _is_single_page_range(pages_value):
    if not pages_value:
        return False
    s = str(pages_value).strip()
    s = re.sub(r"^\s*p+p?\.\s*", "", s, flags=re.I)
    return bool(_SINGLE_PAGE_RE.match(s))


def ieee_structure_errors(reference):
    errors = []
    ref = reference.strip()
    parsed = parse_ieee_reference(ref)

    if not parsed["authors"]:
        errors.append("No authors detected.")

    if _ET_AL_RE.search(ref):
        errors.append("'et al.' is not allowed in IEEE references; list all authors.")

    title_match = re.search(r"[\"“].+?[\"”]", ref)
    year_match = _IEEE_YEAR_RE.search(ref)
    cut = len(ref)
    if title_match:
        cut = min(cut, title_match.start())
    if year_match:
        cut = min(cut, year_match.start())
    author_block = ref[:cut].strip().rstrip(",.")

    if author_block:
        entries, _ = _split_ieee_author_entries(author_block)
        bad_full = []
        bad_comma = []
        for entry in entries:
            cleaned = re.sub(r"^(?:and|&)\s+", "", entry, flags=re.I).strip()
            if _is_full_name_form(cleaned) and not _is_ieee_initial_form(cleaned):
                bad_full.append(cleaned)
            if re.match(r"^[A-Z][A-Za-z'\-]+\s*,\s*[A-Z]\.", cleaned):
                bad_comma.append(cleaned)

        if bad_full:
            errors.append(
                "Full-name author(s) must be initials-first "
                f"(e.g. 'R. Lima'): {', '.join(bad_full)}"
            )
        if bad_comma:
            errors.append(
                "Authors must use 'Initials Surname', not 'Surname, Initials': "
                f"{', '.join(bad_comma)}"
            )

    has_quoted_title = bool(re.search(r"[\"“].+?[\"”]", ref))
    if parsed["source_type"] in {"Journal Article", "Conference Paper"} and not has_quoted_title:
        errors.append("Article/conference title must be in double quotes.")

    if parsed["doi"] and "https://doi.org/" in ref:
        errors.append("Use 'doi: 10.xxxx/xxxxx' instead of 'https://doi.org/...'.")

    if parsed["source_type"] == "Journal Article":
        if not parsed["volume"]:
            errors.append("Missing 'vol.' for journal article.")
        if not parsed["pages"]:
            errors.append("Missing page range ('pp. Z-W').")
        elif _is_single_page_range(parsed["pages"]):
            errors.append(
                f"Single page number '{parsed['pages']}' — IEEE requires a page range ('pp. Z-W')."
            )

    if ref and not re.search(r"[.\]]\s*$", ref):
        errors.append("Reference must end with a period.")

    return errors


def ieee_missing_elements(reference):
    parsed = parse_ieee_reference(reference)
    missing = []
    if parsed["source_type"] == "Journal Article":
        if not parsed["volume"]:
            missing.append("vol.")
        if not parsed["pages"]:
            missing.append("pp.")
    if not parsed["doi"] and not parsed["url"]:
        missing.append("doi / url")
    return missing


# =========================================================
# STATISTICS / MATCHING
# =========================================================

def calculate_citation_statistics(citations, clusters=None):
    total = len(citations)
    unique = len({c["number"] for c in citations})
    crowded = 0
    if clusters:
        for cl in clusters:
            needs, _ = evaluate_cluster(cl)
            if needs:
                crowded += 1
    return {
        "total": total,
        "unique": unique,
        "clusters": len(clusters) if clusters else 0,
        "crowded_clusters": crowded,
    }


def match_citations_to_references(citations, parsed_references):
    results = []
    cited_numbers = {c["number"] for c in citations}
    for i, ref in enumerate(parsed_references, start=1):
        results.append({
            "Reference #": i,
            "First Author": ref.get("first_author"),
            "Year": ref.get("year"),
            "Title": ref.get("title"),
            "Cited": i in cited_numbers,
            "Reference": ref.get("raw"),
        })
    return results


def find_orphan_citations(citations, parsed_references):
    total = len(parsed_references)
    missing = []
    seen = set()
    for c in citations:
        n = c["number"]
        if n < 1 or n > total:
            if n in seen:
                continue
            missing.append({
                "Citation": c["raw"],
                "Reference #": n,
                "Problem": f"[{n}] is cited in the text but there is no reference entry #{n}.",
            })
            seen.add(n)
    return missing


def detect_uncited_references(citations, parsed_references):
    cited = {c["number"] for c in citations}
    uncited = []
    for i in range(1, len(parsed_references) + 1):
        if i not in cited:
            uncited.append({
                "Reference #": i,
                "Reference": parsed_references[i - 1].get("raw"),
                "Problem": f"Reference [{i}] appears in the reference list but is never cited in the text.",
            })
    return uncited


def detect_duplicates(parsed_references):
    duplicates = []
    doi_counter = Counter(
        r["doi"].lower() for r in parsed_references if r.get("doi")
    )
    for doi, count in doi_counter.items():
        if count > 1:
            duplicates.append({"Type": "DOI", "Value": doi, "Count": count})
    return duplicates


def calculate_reference_recency(references, manuscript_year):
    manuscript_year = int(manuscript_year)
    start_year = manuscript_year - 9
    rows = []
    recent = older = unknown = future = 0
    for i, ref in enumerate(references, start=1):
        year = None
        m = _IEEE_YEAR_RE.search(ref)
        if m:
            year = int(m.group(1))
        if year is None:
            category = "Year not detected"; unknown += 1
        elif year > manuscript_year:
            category = "Future year"; future += 1
        elif start_year <= year <= manuscript_year:
            category = "Within last 10 years"; recent += 1
        else:
            category = "Older than 10 years"; older += 1
        rows.append({"Reference #": i, "Year": year if year is not None else "Not detected",
                     "Recency": category, "Reference": ref})
    total = len(references)
    return {
        "start_year": start_year, "end_year": manuscript_year, "total": total,
        "recent_count": recent, "older_count": older, "unknown_count": unknown,
        "future_count": future,
        "recent_percentage": recent / total * 100 if total else 0,
        "rows": rows,
    }


# =========================================================
# SOURCE TYPE NORMALIZATION
# =========================================================

CANONICAL_SOURCE_TYPES = [
    "Journal Article",
    "Conference Paper",
    "Book",
    "Book Chapter",
    "Technical Report",
    "Web Page",
    "Other",
]

_SOURCE_TYPE_ALIASES = {
    "journal article": "Journal Article",
    "journal": "Journal Article",
    "article": "Journal Article",
    "research article": "Journal Article",
    "conference paper": "Conference Paper",
    "conference": "Conference Paper",
    "proceedings": "Conference Paper",
    "symposium": "Conference Paper",
    "workshop": "Conference Paper",
    "book": "Book",
    "monograph": "Book",
    "handbook": "Book",
    "book chapter": "Book Chapter",
    "chapter": "Book Chapter",
    "technical report": "Technical Report",
    "report": "Technical Report",
    "tech report": "Technical Report",
    "web page": "Web Page",
    "website": "Web Page",
    "web": "Web Page",
    "online": "Web Page",
    "other": "Other",
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
# DETERMINISTIC IEEE REFERENCE CORRECTION
# =========================================================

def build_local_ieee_reference_correction(reference):
    original = reference.strip()
    ref = original
    notes = []
    placeholder_flags = {
        "missing_vol": False,
        "missing_no": False,
        "missing_pp": False,
        "missing_doi": False,
        "missing_authors": False,
    }

    title_match = re.search(r"[\"“].+?[\"”]", ref)
    year_match = _IEEE_YEAR_RE.search(ref)

    cut = len(ref)
    if title_match:
        cut = min(cut, title_match.start())
    if year_match:
        cut = min(cut, year_match.start())

    author_block = ref[:cut].strip().rstrip(",.")
    rest = ref[cut:]

    # ---- Authors ----
    if author_block:
        author_block = _ET_AL_RE.sub("", author_block).strip(" ,")
        new_authors, changed, author_notes = normalize_ieee_authors(author_block)
        if changed:
            notes.extend(author_notes)
            ref = new_authors + ", " + rest.lstrip()
            ref = re.sub(r"\s{2,}", " ", ref)
            ref = re.sub(r",\s*,", ",", ref)
    else:
        guess = _guess_author_count_from_text(original)
        placeholder = _make_author_placeholder(guess)
        ref = placeholder + (", " + rest.lstrip() if rest else "")
        placeholder_flags["missing_authors"] = True
        notes.append(f"Missing author list — placeholder inserted: {placeholder}")

    if _ET_AL_RE.search(ref):
        ref = _ET_AL_RE.sub("", ref)
        ref = re.sub(r"\s*,\s*,", ",", ref)
        ref = re.sub(r"\s{2,}", " ", ref).strip()
        notes.append("Removed 'et al.' — IEEE requires all authors to be listed.")

    # ---- DOI normalization ----
    dm = _IEEE_DOI_RE.search(ref)
    if dm:
        doi = dm.group(0).rstrip(".,;")
        old = ref
        ref = re.sub(
            r"https?://(?:dx\.)?doi\.org/" + _esc(doi),
            f"doi: {doi}",
            ref,
            flags=re.I,
        )
        ref = re.sub(r"\bdoi\s*:\s*https?://\S+", f"doi: {doi}", ref, flags=re.I)
        if ref != old:
            notes.append("DOI converted to 'doi: ...' form.")

    # ---- Journal-article placeholders (rebuilt cleanly) ----
    parsed_now = parse_ieee_reference(ref)
    original_parsed = parse_ieee_reference(original)

    if parsed_now["source_type"] == "Journal Article":
        missing_vol = not original_parsed["volume"]
        missing_no = not original_parsed["issue"]
        missing_pp = not original_parsed["pages"]

        ref = re.sub(r",?\s*vol\.\s*\?{3}", "", ref)
        ref = re.sub(r",?\s*no\.\s*\?{3}", "", ref)
        ref = re.sub(r",?\s*pp\.\s*\?{3}-\?{3}", "", ref)
        ref = re.sub(r",?\s*pp\.\s*(\d+)-\?{3}", r" pp. \1", ref)
        ref = re.sub(r"\s{2,}", " ", ref)
        ref = re.sub(r",\s*,", ",", ref)
        ref = ref.strip().rstrip(", ")

        parsed_now = parse_ieee_reference(ref)

        year_tail_match = re.search(
            r"(,\s*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s*)?((?:19|20)\d{2})",
            ref,
        )
        if year_tail_match:
            tail_start = year_tail_match.start()
            head = ref[:tail_start].rstrip(", ")
            tail = ref[tail_start:]

            head = re.sub(r",?\s*vol\.\s*[\w\-?]+", "", head)
            head = re.sub(r",?\s*no\.\s*[\w\-?]+", "", head)
            head = re.sub(r",?\s*pp?\.\s*[\w\-–—?]+(?:\s*[-–—]\s*[\w\-–—?]+)?", "", head)
            head = re.sub(r"\s{2,}", " ", head).rstrip(", ")

            pieces = []

            vol = parsed_now["volume"]
            if vol:
                pieces.append(f"vol. {vol}")
            elif missing_vol:
                pieces.append("vol. ???")
                placeholder_flags["missing_vol"] = True
                notes.append("Missing volume — placeholder inserted: vol. ???")

            no = parsed_now["issue"]
            if no:
                pieces.append(f"no. {no}")
            elif missing_no:
                pieces.append("no. ???")
                placeholder_flags["missing_no"] = True
                notes.append("Missing issue — placeholder inserted: no. ???")

            pages = parsed_now["pages"]
            if pages:
                if _is_single_page_range(pages):
                    pages = f"{pages}-???"
                    placeholder_flags["missing_pp"] = True
                    notes.append(
                        f"Single page number — expanded to range placeholder: "
                        f"pp. {pages}"
                    )
                pieces.append(f"pp. {pages}")
            elif missing_pp:
                pieces.append("pp. ???-???")
                placeholder_flags["missing_pp"] = True
                notes.append("Missing pages — placeholder inserted: pp. ???-???")

            ref = head + (", " + ", ".join(pieces) if pieces else "") + tail
        else:
            if missing_vol and "vol." not in ref.lower():
                ref = ref.rstrip(".").rstrip(", ") + ", vol. ???"
                placeholder_flags["missing_vol"] = True
                notes.append("Missing volume — placeholder inserted: vol. ???")
            if missing_no and "no." not in ref.lower():
                ref = ref.rstrip(".").rstrip(", ") + ", no. ???"
                placeholder_flags["missing_no"] = True
                notes.append("Missing issue — placeholder inserted: no. ???")
            if missing_pp and "pp." not in ref.lower():
                ref = ref.rstrip(".").rstrip(", ") + ", pp. ???-???"
                placeholder_flags["missing_pp"] = True
                notes.append("Missing pages — placeholder inserted: pp. ???-???")

    # ---- DOI/URL placeholder ----
    parsed_now = parse_ieee_reference(ref)
    if not parsed_now["doi"] and not parsed_now["url"]:
        if "doi:" not in ref.lower() and "http" not in ref.lower():
            ref = ref.rstrip(".").rstrip(", ") + f", {PLACEHOLDER_DOI}."
            placeholder_flags["missing_doi"] = True
            notes.append(f"Missing DOI/URL — placeholder inserted: {PLACEHOLDER_DOI}")

    # ---- Final normalization pass ----
    ref = re.sub(r",\s*,", ",", ref)
    ref = re.sub(r"\s+,", ",", ref)
    ref = re.sub(r",\s*", ", ", ref)
    ref = re.sub(r"\s{2,}", " ", ref)
    ref = re.sub(r",\s*\.", ".", ref)
    ref = re.sub(r"\s+\.", ".", ref)
    ref = re.sub(r"\.\.+$", ".", ref)

    ref = re.sub(
        r"(\b(?:vol|no)\.\s*\?{3})(?!\s*[,.])(\s+)(?!\d)",
        r"\1, \2",
        ref,
    )
    ref = re.sub(
        r"(\bpp\.\s*(?:\?{3}|\d+-\?{3}|\?{3}-\?{3}))(?!\s*[,.])(\s+)(?!\d)",
        r"\1, \2",
        ref,
    )

    ref = ref.strip().rstrip(",")
    if ref and not ref.endswith("."):
        ref += "."
        notes.append("Trailing period added.")

    changed = ref != original
    return {
        "Corrected": ref,
        "Status": "REVISED" if changed else "MATCH",
        "Note": " | ".join(notes) if notes else "",
        "Placeholders": placeholder_flags,
    }


# =========================================================
# OPENAI HELPERS
# =========================================================

def _get_openai_client():
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        return None
    try:
        from openai import OpenAI
        return OpenAI(api_key=api_key)
    except Exception:
        return None


def _safe_json_loads(text):
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
        text = re.sub(r"\s*```$", "", text)
    return json.loads(text)


# =========================================================
# OPENAI — IEEE REFERENCE REVIEW
# =========================================================

def review_ieee_references_with_ai(references):
    client = _get_openai_client()
    if client is None or not references:
        return []

    payload = [{"number": i, "reference": clean_text(r)}
               for i, r in enumerate(references, start=1)]

    prompt = f"""
You are checking IEEE-style REFERENCE-LIST entries.

IEEE STYLE RULES:
- Author names: "A. B. Smith" (initials first). Full given names are wrong.
- All authors must be listed; "et al." is NOT allowed.
- Article titles: double quotes, Title Case.
- Journal / conference names: italic (list them in "italic_elements").
- SOURCE TYPE PRECEDENCE IS STRICT:
  a) Conference Paper if the entry or venue contains Proc./Proceedings/Conference/Symposium/Workshop/Congress/Colloquium/Convention/Seminar/Procedia/Meeting.
  b) Journal Article only with clear journal evidence such as Journal/Transactions/Letters/Review, or journal-like volume+pages metadata.
  c) Otherwise classify Book, Book Chapter, Technical Report, Web Page, or Other as applicable.
- NEVER classify an entry as Journal Article merely because a DOI is present. Conference and proceedings papers also have DOIs.
- Volume: "vol. X". Issue: "no. Y". Pages: "pp. Z-W".
- DOI: "doi: 10.xxxx/xxxxx" — never https://doi.org/.
- Must end with a period.

Return JSON only: {{"results": [{{"number": int, "status": "OK"|"REVISED"|"MANUAL_CHECK", "revised_reference": str, "source_type": str, "italic_elements": str, "year": int|null, "missing_required_elements": [str], "explanation": str}}, ...]}}

INPUT:
{json.dumps(payload, ensure_ascii=False)}
"""
    try:
        response = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": "You are a precise IEEE reference editor. Return valid JSON only."},
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
        return [{
            "number": 0, "status": "MANUAL_CHECK", "revised_reference": "",
            "source_type": "Other", "italic_elements": "", "year": None,
            "missing_required_elements": [],
            "explanation": f"OpenAI API error: {exc}",
        }]


def _reference_is_hallucinated(original, corrected):
    if not corrected:
        return False
    numbers = lambda t: set(re.findall(r"\b\d+(?:\.\d+)?\b", t or ""))
    dois = lambda t: set(re.findall(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", t or "", re.I))
    urls = lambda t: set(re.findall(r"https?://\S+", t or "", re.I))
    if numbers(corrected) - numbers(original): return True
    if dois(corrected) - dois(original): return True
    if urls(corrected) - urls(original): return True
    return False


# =========================================================
# OPENAI — IEEE CITATION CLUSTER REVIEW
# =========================================================

def review_ieee_citations_with_ai(clusters, reference_rows):
    client = _get_openai_client()
    if client is None or not clusters:
        return []

    payload = [
        {"number": i, "raw": c["raw"], "canonical": c["canonical"],
         "numbers": c["numbers"]}
        for i, c in enumerate(clusters, start=1)
    ]

    reference_context = [
        {"no": row["No."], "year": row.get("Year"),
         "authors": parse_ieee_reference(row.get("Original Reference", "")).get("authors", []),
         "reference": row.get("Original Reference", "")}
        for row in reference_rows
    ]

    prompt = f"""
You are checking IEEE-style IN-TEXT citation clusters.

Each item is a bracketed cluster already extracted. Correct each IN ISOLATION.

IEEE RULES:
- Citations must be ascending: [1], [2], [3] — never [3], [1], [2].
- Three or more consecutive numbers MUST use range form: [1]-[3].
- Two consecutive numbers stay separate: [1], [2].
- Non-consecutive stay comma-separated: [1], [3], [5].
- Do NOT invent numbers. Do NOT merge separate clusters.

Return JSON only:
{{"results": [
  {{"number": int, "status": "OK"|"REVISED"|"MANUAL_CHECK",
    "revised_citation": str, "explanation": str}}
]}}

REFERENCE LIST CONTEXT:
{json.dumps(reference_context, ensure_ascii=False)}

INPUT CLUSTERS:
{json.dumps(payload, ensure_ascii=False)}
"""
    try:
        response = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": "You are a precise IEEE citation editor. JSON only."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        return _safe_json_loads(response.choices[0].message.content).get("results", [])
    except Exception as exc:
        return [{"number": 0, "status": "MANUAL_CHECK", "revised_citation": "",
                 "explanation": f"OpenAI API error: {exc}"}]


# =========================================================
# PRE-CLASSIFICATION (cost control)
# =========================================================

def preclassify_ieee_references(references, verification_rows, local_source_types):
    review_payload = []
    preclassified = {}

    for i, (ref, v, stype) in enumerate(
        zip(references, verification_rows, local_source_types), start=1
    ):
        parsed = parse_ieee_reference(ref)

        if _is_uninterpretable_reference(ref):
            preclassified[i] = {
                "number": i, "status": "MANUAL_CHECK", "revised_reference": "",
                "source_type": "Other", "italic_elements": "", "year": None,
                "missing_required_elements": [],
                "explanation": f"Cannot interpret reference [{i}] from the extracted PDF. The reference number has been preserved.",
            }
            continue

        if stype == "Journal Article":
            if not parsed.get("doi"):
                preclassified[i] = {
                    "number": i, "status": "WITHHELD", "revised_reference": "",
                    "source_type": stype, "italic_elements": "", "year": parsed.get("year"),
                    "missing_required_elements": [],
                    "explanation": "DOI not provided — automated verification/correction withheld.",
                }
                continue
            if v.get("suspicious") or not v.get("checked"):
                reason = " | ".join(v.get("reasons", [])) or "DOI metadata could not be independently verified."
                preclassified[i] = {
                    "number": i, "status": "MANUAL_CHECK", "revised_reference": "",
                    "source_type": stype, "italic_elements": "", "year": parsed.get("year"),
                    "missing_required_elements": [],
                    "explanation": reason,
                }
                continue
            review_payload.append({"number": i, "reference": ref, "openalex": {
                "title": v.get("crossref_title"),
                "authors": v.get("crossref_authors", []),
                "year": parsed.get("year"),
            }})
            continue

        if stype == "Conference Paper" and v.get("checked") and not v.get("suspicious"):
            review_payload.append({"number": i, "reference": ref, "openalex": {
                "title": v.get("crossref_title"),
                "authors": v.get("crossref_authors", []),
                "year": parsed.get("year"),
            }})
            continue

        preclassified[i] = {
            "number": i, "status": "MANUAL_CHECK", "revised_reference": "",
            "source_type": stype, "italic_elements": "", "year": parsed.get("year"),
            "missing_required_elements": [],
            "explanation": "No independently verified metadata workflow is configured for this source type; manual verification required.",
        }

    return review_payload, preclassified


# =========================================================
# STATUS / FALLBACK LABELS
# =========================================================

def _ieee_status_label(status):
    status = str(status or "MANUAL_CHECK").upper().strip()
    if status in {"OK", "PASS", "MATCH"}:
        return "MATCH"
    if status in {"REVISED", "NEEDS REVIEW", "NEEDS_REVIEW"}:
        return "REVISED"
    if status == "WITHHELD":
        return "WITHHELD"
    return "MANUAL CHECK"


def _fallback_ieee_italic_elements(source_type, parsed):
    if source_type in {"Journal Article", "Conference Paper"} and parsed.get("venue"):
        return parsed["venue"]
    if source_type in {"Book", "Book Chapter", "Technical Report"} and parsed.get("title"):
        return parsed["title"]
    return ""

def compute_ieee_italic_tokens(source_type, corrected_reference, parsed=None):
    """
    Deterministically derive the list of substrings that must be italic
    in an IEEE corrected reference, based on source type.

    Returns a list of exact substrings (as they appear in
    corrected_reference) that should be italicized.
    """
    if not corrected_reference:
        return []

    if parsed is None:
        parsed = parse_ieee_reference(corrected_reference)

    tokens = []
    ref = corrected_reference

    def _add(token):
        token = (token or "").strip()
        if not token:
            return
        # Only add if the token literally appears in the reference
        if token in ref and token not in tokens:
            tokens.append(token)

    # ---------------- JOURNAL ARTICLE ----------------
    # Italic: journal name (venue). Volume/issue/pages are NOT italic.
    if source_type == "Journal Article":
        venue = parsed.get("venue")
        if venue:
            _add(venue)
            return tokens

        # Fallback: try to infer the journal name from the reference shape.
        # Journal name = the comma-delimited phrase right after the closing
        # quote and before "vol." / a volume number.
        m = re.search(
            r'"\s*,\s*([^,]+?)\s*,\s*(?:vol\.|no\.|\d+\s*,)',
            ref,
        )
        if m:
            _add(m.group(1).strip())
        return tokens

    # ---------------- CONFERENCE PAPER ----------------
    # Italic: conference or proceedings name.
    if source_type == "Conference Paper":
        venue = parsed.get("venue")
        if venue:
            _add(venue)
            return tokens

        # Fallback: capture the proceedings/conference name after "in".
        # This also handles forms such as:
        #   in 2024 Asia Pacific Conference on Innovation in Technology APCIT 2024, 2024.
        # where the venue does not begin with the word "Conference".
        m = re.search(
            r'\bin\s+(.+?)(?=,\s*(?:pp?\.|(?:19|20)\d{2}\b|doi\s*:))',
            ref,
            re.I,
        )
        if m:
            candidate = m.group(1).strip().rstrip(',.')
            if re.search(r'\b(?:proc\.?|proceedings|conference|symposium|workshop)\b', candidate, re.I):
                _add(candidate)
        return tokens

    # ---------------- BOOK ----------------
    # Italic: the book title. In IEEE, book titles are usually NOT in
    # quotes, so the title is the phrase after the author block and
    # before a publisher or edition marker.
    if source_type == "Book":
        title = parsed.get("title")
        if title:
            _add(title)
            return tokens

        m = re.search(r"\.\s*([^.]+?)\s*,\s*[A-Z][a-zA-Z]+", ref)
        if m:
            _add(m.group(1).strip())
        return tokens

    # ---------------- BOOK CHAPTER ----------------
    # Italic: the containing book title, not the chapter title.
    if source_type == "Book Chapter":
        # IEEE convention: chapter title in quotes, book title italic
        m = re.search(
            r'"[^"]+"\s*,\s*in\s+([^,]+?)\s*,',
            ref,
        )
        if m:
            _add(m.group(1).strip())
            return tokens
        return tokens

    # ---------------- TECHNICAL REPORT ----------------
    # Italic: report title.
    if source_type == "Technical Report":
        title = parsed.get("title")
        if title:
            _add(title)
            return tokens
        return tokens

    # ---------------- WEB PAGE ----------------
    # Italic: page/document title.
    if source_type == "Web Page":
        title = parsed.get("title")
        if title:
            _add(title)
            return tokens
        return tokens

    # ---------------- OTHER / UNKNOWN ----------------
    return tokens

# =========================================================
# STAGED PIPELINE
# =========================================================

def extract_ieee_pdf(uploaded_file, batch, manuscript_year):
    """
    Stage 1: read PDF, pull text + style spans, clean, find reference
    section, split references, extract citations. No DOI/OpenAlex calls,
    no AI calls.
    """
    uploaded_file.seek(0)
    full_text, pages, removed_running_text = extract_pdf_text(uploaded_file)
    uploaded_file.seek(0)
    style_spans = extract_pdf_style_spans(uploaded_file)
    full_text = clean_text(full_text)

    reference_text, heading, body_text = find_reference_section(full_text)
    if reference_text is None:
        raise ValueError("I could not detect a References section.")

    # ---- Estimate expected bibliography size BEFORE splitting ----
    expected_count = estimate_expected_reference_count(reference_text, full_text)

    # ---- Split with expected count available for marker-gap recovery ----
    references = split_references(reference_text, expected_count=expected_count)

    # ---- Attempt recovery if we came up short ----
    if expected_count and len(references) < expected_count * 0.9:
        recovered = recover_glued_references(references, expected_count)
        if len(recovered) > len(references):
            references = recovered

    extraction_quality = {
        "expected_count": expected_count,
        "actual_count": len(references),
        "complete": (
            expected_count is None
            or len(references) >= int(expected_count * 0.9)
        ),
    }

    parsed_refs = [parse_ieee_reference(r) for r in references]
    local_source_types = [p.get("source_type", "Other") for p in parsed_refs]

    citations, clusters = extract_ieee_citations(body_text)
    for c in citations:
        c["page"] = _page_for_offset(full_text, c.get("offset"))

    batch.update({
        "filename": uploaded_file.name,
        "reference_text": reference_text,
        "heading": heading,
        "references": references,
        "parsed_references": parsed_refs,
        "local_source_types": local_source_types,
        "citations": citations,
        "citation_clusters": clusters,
        "removed_running_text": removed_running_text,
        "style_spans": style_spans,
        "manuscript_year": int(manuscript_year),
        "expected_reference_count": expected_count,
        "reference_extraction_quality": extraction_quality,
        "ai_done": False,
    })
    return batch


def process_ieee_references_and_citations(batch, client, manuscript_year):
    """
    Stage 2: verify DOIs via OpenAlex, run AI reference + citation review,
    build all final rows and statistics. Assumes batch was already
    populated by extract_ieee_pdf().
    """
    references = batch["references"]
    parsed_refs = batch["parsed_references"]
    local_source_types = batch["local_source_types"]
    citations = batch["citations"]
    clusters = batch["citation_clusters"]

    # ---- DOI verification via OpenAlex ----
    verification_rows = []
    for ref, parsed in zip(references, parsed_refs):
        if _is_uninterpretable_reference(ref):
            v = {
                "checked": False, "suspicious": False,
                "reasons": ["Reference text could not be interpreted from PDF extraction."],
                "openalex_metadata": None, "crossref_title": None,
                "crossref_authors": [], "title_similarity": None, "author_overlap": None,
            }
        else:
            v = verify_reference_against_openalex(ref, parsed)
        verification_rows.append(v)

    # ---- Pre-classification + AI reference review ----
    review_payload, preclassified = preclassify_ieee_references(
        references, verification_rows, local_source_types
    )

    ref_ai = review_ieee_references_with_ai(
        [x["reference"] for x in review_payload]
    ) if review_payload else []

    by_no = {
        int(x.get("number", -1)): x for x in (ref_ai or [])
        if str(x.get("number", "")).isdigit()
    }
    by_no.update(preclassified)

    rows = []
    for i, original in enumerate(references, start=1):
        ai = by_no.get(i, {})
        original_clean = strip_markdown_markers(clean_text(original))
        structure_errors = ieee_structure_errors(original_clean)

        v = verification_rows[i - 1]
        uninterpretable = _is_uninterpretable_reference(original_clean)
        oa_built = None
        if (not uninterpretable) and v.get("checked") and not v.get("suspicious") and v.get("openalex_metadata"):
            oa_built = build_ieee_reference_from_openalex(
                original_clean, v.get("openalex_metadata")
            )

        if uninterpretable:
            corrected = ""
            source_type = "Other"
            source_type_origin = "Extraction-Placeholder"
            parsed = parse_ieee_reference("")
            deterministic_tokens = []
            italic_elements = ""
            local = {"Corrected": "", "Note": f"Cannot interpret reference [{i}] from the extracted PDF; numbering preserved."}
            ai = {
                "status": "MANUAL_CHECK",
                "explanation": f"Cannot interpret reference [{i}] from the extracted PDF. Please check the original manuscript manually.",
                "missing_required_elements": [],
            }
        elif oa_built:
            # DOI resolved and matched: OpenAlex bibliographic metadata is authoritative.
            corrected = oa_built["Corrected"]
            source_type = oa_built["Source Type"]
            source_type_origin = "OpenAlex"
            parsed = parse_ieee_reference(corrected)
            deterministic_tokens = oa_built.get("Italic Tokens", [])
            italic_elements = ", ".join(deterministic_tokens)
            local = {"Corrected": corrected, "Note": "Constructed from verified OpenAlex DOI metadata."}
        else:
            # No verified OpenAlex metadata: NEVER reconstruct bibliographic facts
            # from AI guesses. Preserve the extracted reference exactly (apart from
            # whitespace cleanup) and flag it for manual review when necessary.
            corrected = original_clean
            parsed = parse_ieee_reference(corrected)
            local = {"Corrected": corrected, "Note": "Original preserved because no verified OpenAlex DOI metadata was available."}

            local_source_type = normalize_source_type(parsed.get("source_type"))
            ai_source_type = normalize_source_type(ai.get("source_type"))

            # Strong conference/proceedings evidence in the actual reference wins
            # over an AI journal guess. DOI presence is intentionally ignored here.
            strong_conference_signal = bool(re.search(
                r"\b(?:proc\.?|proceedings|conference|symposium|workshop|congress|"
                r"congres|colloquium|convention|seminar|procedia|meeting)\b",
                corrected.lower(),
            ))
            if strong_conference_signal and local_source_type != "Book":
                source_type = "Conference Paper"
                source_type_origin = "Local-Strong"
            elif local_source_type != "Other":
                source_type = local_source_type
                source_type_origin = "Local"
            else:
                source_type = ai_source_type
                source_type_origin = "AI"

            deterministic_tokens = compute_ieee_italic_tokens(source_type, corrected, parsed)
            italic_elements = ", ".join(deterministic_tokens)

        year = oa_built.get("Year") if oa_built else ai.get("year")
        if not isinstance(year, int):
            year = int(parsed["year"]) if parsed["year"] else None

        missing = ai.get("missing_required_elements", []) or []
        if isinstance(missing, str):
            missing = [missing]
        missing = list({*missing, *ieee_missing_elements(corrected)})

        status = _ieee_status_label(ai.get("status")) if ai else "NOT AI CHECKED"
        if ai.get("status") == "WITHHELD":
            corrected_display = "— WITHHELD —"
        elif v.get("suspicious"):
            corrected_display = "— WITHHELD (DOI mismatch) —"
        else:
            corrected_display = corrected

        rows.append({
            "No.": i,
            "Source Type": source_type,
            "Source Type Origin": source_type_origin,
            "Year": year,
            "Original Reference": (
                f"Cannot interpret reference [{i}] from the extracted PDF."
                if uninterpretable else original_clean
            ),
            "Corrected Version": corrected_display,
            "Italicized in IEEE": italic_elements,
            "Italic Tokens": deterministic_tokens,
            "Status": status,
            "Missing Required Elements": ", ".join(str(x) for x in missing),
            "AI Explanation": ai.get("explanation", ""),
            "Original Structure Errors": " | ".join(structure_errors),
            "Original Has Structure Error": bool(structure_errors),
            "Correction Note": local.get("Note", ""),
            "Metadata Source": "OpenAlex" if oa_built else ("AI/Local" if ai else "Local"),
            "DOI Verified": v["checked"],
            "DOI Suspicious": v["suspicious"],
            "DOI Verification Reasons": " | ".join(v["reasons"]),
            "Title Similarity": v.get("title_similarity"),
            "Author Overlap": v.get("author_overlap"),
            "OpenAlex Title": v.get("crossref_title"),
            "OpenAlex Authors": ", ".join(v.get("crossref_authors", [])[:5]),
            "Placeholders": local.get("Placeholders", {}),
        })

    # ---- AI citation cluster review ----
    cit_ai = review_ieee_citations_with_ai(clusters, rows) if clusters else []
    cit_by_no = {
        int(x.get("number", -1)): x for x in (cit_ai or [])
        if str(x.get("number", "")).isdigit()
    }

    cluster_rows = []
    for i, cl in enumerate(clusters, start=1):
        ai = cit_by_no.get(i, {})
        revised = enforce_ieee_citation_rules(cl)
        needs_fix, reasons = evaluate_cluster(cl)
        cluster_rows.append({
            "No.": i,
            "Page": next((c.get("page") for c in citations if c["number"] in cl["numbers"]), None),
            "Original Form": cl["raw"],
            "Corrected Form": revised,
            "Numbers Cited": ", ".join(str(n) for n in cl["numbers"]),
            "Status": "REVISED" if needs_fix else "MATCH",
            "Reason": " | ".join(reasons) or ai.get("explanation", ""),
        })

    # ---- Statistics / matching ----
    citation_stats = calculate_citation_statistics(citations, clusters)
    matching = match_citations_to_references(citations, parsed_refs)
    orphan = find_orphan_citations(citations, parsed_refs)
    uncited = detect_uncited_references(citations, parsed_refs)
    duplicates = detect_duplicates(parsed_refs)
    recency = calculate_reference_recency(references, manuscript_year)

    batch.update({
        "citation_stats": citation_stats,
        "matching_results": matching,
        "orphan_citations": orphan,
        "uncited_references": uncited,
        "duplicates": duplicates,
        "recency": recency,
        "reference_comparison": rows,
        "cluster_rows": cluster_rows,
        "verification_rows": verification_rows,
        "ai_done": True,
        "ai_complete": True,
    })
    return batch


# =========================================================
# DOCX EXPORT — Wahyudin-style layout
# =========================================================

RED = RGBColor(0xC0, 0x00, 0x00)


def _set_run_font(run, size_pt=11, bold=False, italic=False, color=None):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def _add_run(paragraph, text, size_pt=11, bold=False, italic=False, red=False):
    r = paragraph.add_run(text)
    _set_run_font(
        r, size_pt=size_pt, bold=bold, italic=italic,
        color=RED if red else None,
    )
    return r


def _docx_set_default_font(document, font_name="Times New Roman", size_pt=11):
    """Set the Normal style with zero default paragraph spacing so our
    explicit Pt() values are the ONLY source of vertical rhythm."""
    style = document.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(size_pt)
    pf = style.paragraph_format
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.line_spacing = 1.15


def _add_divider(doc):
    """Thin grey horizontal rule with controlled spacing. Used between
    citation/reference entries to visually separate them without adding
    excessive whitespace."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.0
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "BFBFBF")
    pBdr.append(bottom)
    pPr.append(pBdr)


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

_PAGE_PARTIAL_RE = re.compile(r"pp\.\s*(\d+)-\?{3}", re.I)


def _emit_with_placeholders(paragraph, text, italic=False, size_pt=11):
    pos = 0
    for m in _PLACEHOLDER_RE.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], size_pt=size_pt, italic=italic)

        matched = m.group(0)
        page_partial = _PAGE_PARTIAL_RE.match(matched)
        if page_partial:
            known = page_partial.group(1)
            prefix = matched[: matched.index(known)]
            _add_run(paragraph, prefix, size_pt=size_pt, italic=italic)
            _add_run(paragraph, f"{known}-???", size_pt=size_pt,
                     italic=True, bold=True, red=True)
        else:
            _add_run(paragraph, matched, size_pt=size_pt,
                     italic=True, bold=True, red=True)

        pos = m.end()
    if pos < len(text):
        _add_run(paragraph, text[pos:], size_pt=size_pt, italic=italic)


def _highlight_missing_tokens_in_corrected(paragraph, text, missing_tokens, size_pt=11):
    """
    Render `text` such that:
      - Placeholder tokens (vol. ???, no. ???, pp. ???-???, doi: ???, pp. N-???)
        are red bold italic.
      - Tokens from `missing_tokens` (e.g. "pp.", "vol.", "no.", "doi") are also
        red bold italic wherever they appear as isolated IEEE markers.
      - Everything else is normal weight.
    """
    parts = []
    parts.append(
        r"(?P<ph>"
        r"author \?{3}"
        r"|author \d+(?:, author \d+)+"
        r"|vol\. \?{3}"
        r"|no\. \?{3}"
        r"|pp\. \d+-\?{3}"
        r"|pp\. \?{3}-\?{3}"
        r"|doi: \?{3}"
        r")"
    )
    if missing_tokens:
        escaped = sorted({re.escape(t) for t in missing_tokens if t}, key=len, reverse=True)
        if escaped:
            parts.append(r"(?P<miss>\b(?:" + "|".join(escaped) + r")\b)")

    if not parts:
        _add_run(paragraph, text, size_pt=size_pt)
        return

    pattern = re.compile("|".join(parts), re.I)
    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], size_pt=size_pt)

        matched = m.group(0)
        if m.lastgroup == "ph":
            page_partial = _PAGE_PARTIAL_RE.match(matched)
            if page_partial:
                known = page_partial.group(1)
                prefix = matched[: matched.index(known)]
                _add_run(paragraph, prefix, size_pt=size_pt, italic=True, bold=True, red=True)
                _add_run(paragraph, f"{known}-???", size_pt=size_pt,
                         italic=True, bold=True, red=True)
            else:
                _add_run(paragraph, matched, size_pt=size_pt,
                         italic=True, bold=True, red=True)
        else:
            _add_run(paragraph, matched, size_pt=size_pt,
                     italic=True, bold=True, red=True)

        pos = m.end()
    if pos < len(text):
        _add_run(paragraph, text[pos:], size_pt=size_pt)

def _highlight_missing_tokens_in_corrected_italic(paragraph, text, missing_tokens, size_pt=11):
    """Same as _highlight_missing_tokens_in_corrected, but the default
    weight is italic."""
    parts = []
    parts.append(
        r"(?P<ph>"
        r"author \?{3}"
        r"|author \d+(?:, author \d+)+"
        r"|vol\. \?{3}"
        r"|no\. \?{3}"
        r"|pp\. \d+-\?{3}"
        r"|pp\. \?{3}-\?{3}"
        r"|doi: \?{3}"
        r")"
    )
    if missing_tokens:
        escaped = sorted({re.escape(t) for t in missing_tokens if t}, key=len, reverse=True)
        if escaped:
            parts.append(r"(?P<miss>\b(?:" + "|".join(escaped) + r")\b)")

    pattern = re.compile("|".join(parts), re.I)
    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], size_pt=size_pt, italic=True)
        matched = m.group(0)
        if m.lastgroup == "ph":
            page_partial = _PAGE_PARTIAL_RE.match(matched)
            if page_partial:
                known = page_partial.group(1)
                prefix = matched[: matched.index(known)]
                _add_run(paragraph, prefix, size_pt=size_pt, italic=True, bold=True, red=True)
                _add_run(paragraph, f"{known}-???", size_pt=size_pt,
                         italic=True, bold=True, red=True)
            else:
                _add_run(paragraph, matched, size_pt=size_pt,
                         italic=True, bold=True, red=True)
        else:
            _add_run(paragraph, matched, size_pt=size_pt,
                     italic=True, bold=True, red=True)
        pos = m.end()
    if pos < len(text):
        _add_run(paragraph, text[pos:], size_pt=size_pt, italic=True)

def _add_ieee_reference_with_italics(paragraph, text, italic_elements, missing_tokens=None, size_pt=11):
    """Emit corrected reference; italicize venue/title tokens; skip DOI spans;
    highlight placeholders AND missing marker tokens in red bold italic.
    Italicizes only the FIRST occurrence of each token."""
    doi_spans = []
    for m in re.finditer(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", text):
        doi_spans.append((m.start(), m.end()))

    def _in_doi(pos):
        return any(s <= pos < e for s, e in doi_spans)

    if missing_tokens is None:
        missing_tokens = []

    if isinstance(italic_elements, (list, tuple)):
        tokens = [str(t).strip() for t in italic_elements if str(t).strip()]
    else:
        # Backward compatibility for older rows. New rows should pass Italic Tokens
        # as a list so venue names containing commas are never split accidentally.
        tokens = [str(italic_elements).strip()] if str(italic_elements or "").strip() else []
    if not tokens:
        _highlight_missing_tokens_in_corrected(paragraph, text, missing_tokens, size_pt=size_pt)
        return

    # Build a list of (start, end) spans, one per token, first occurrence only
    spans = []
    for tok in tokens:
        if not tok:
            continue
        pat = re.compile(re.escape(tok), re.I)
        m = pat.search(text)
        if m:
            spans.append((m.start(), m.end()))

    # Sort and drop overlapping spans (keep earliest start)
    spans.sort()
    merged = []
    for s, e in spans:
        if merged and s < merged[-1][1]:
            # Overlap — merge into the existing span
            merged[-1] = (merged[-1][0], max(e, merged[-1][1]))
        else:
            merged.append((s, e))

    pos = 0
    for s, e in merged:
        if _in_doi(s):
            # Span overlaps a DOI — emit as normal (non-italic), then advance
            if s > pos:
                _highlight_missing_tokens_in_corrected(
                    paragraph, text[pos:s], missing_tokens, size_pt=size_pt
                )
            _highlight_missing_tokens_in_corrected(
                paragraph, text[s:e], missing_tokens, size_pt=size_pt
            )
            pos = e
            continue

        if s > pos:
            _highlight_missing_tokens_in_corrected(
                paragraph, text[pos:s], missing_tokens, size_pt=size_pt
            )
        # Italic span — italicize the whole span, but still apply
        # placeholder highlighting inside it if present.
        _highlight_missing_tokens_in_corrected_italic(
            paragraph, text[s:e], missing_tokens, size_pt=size_pt
        )
        pos = e

    if pos < len(text):
        _highlight_missing_tokens_in_corrected(
            paragraph, text[pos:], missing_tokens, size_pt=size_pt
        )

def _ref_is_withheld(row):
    cv = (row.get("Corrected Version") or "").strip()
    return cv.startswith("— WITHHELD")


def _ref_is_suspicious(row):
    return bool(row.get("DOI Suspicious"))


def _ref_is_manual(row):
    return str(row.get("Status", "")).upper().strip() == "MANUAL CHECK"


def _ref_has_placeholder(row):
    ph = row.get("Placeholders") or {}
    return any(bool(v) for v in ph.values())


def build_ieee_docx(result):
    doc = Document()
    _docx_set_default_font(doc)

    filename = result.get("filename", "manuscript.pdf")
    manuscript_year = result.get("manuscript_year", datetime.now().year)

    reference_rows = result.get("reference_comparison") or []
    cluster_rows = result.get("cluster_rows") or []
    citations = result.get("citations", [])
    references = result.get("references", [])
    stats = result.get("citation_stats", {})
    recency = result.get("recency", {})
    orphan = result.get("orphan_citations", []) or []
    uncited = result.get("uncited_references", []) or []
    matching = result.get("matching_results", []) or []

    total_refs = len(references)
    total_cits = stats.get("total", 0)

    doi_checked = sum(1 for r in reference_rows if r.get("DOI Verified"))
    doi_suspicious = sum(1 for r in reference_rows if r.get("DOI Suspicious"))
    doi_suspicious_pct = (doi_suspicious / total_refs * 100) if total_refs else 0

    withheld_count = sum(1 for r in reference_rows if _ref_is_withheld(r))
    manual_count = sum(1 for r in reference_rows if _ref_is_manual(r))
    placeholder_count = sum(1 for r in reference_rows if _ref_has_placeholder(r))

    window_start = recency.get("start_year", manuscript_year - 9)
    window_end = recency.get("end_year", manuscript_year)

    # ---- Title block ----
    title = doc.add_heading(level=0)
    tr = title.add_run(filename)
    _set_run_font(tr, size_pt=18, bold=True)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    sr = subtitle.add_run("IEEE Reference Style — Citation & Reference Diagnostic Report")
    _set_run_font(sr, size_pt=11, italic=True)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = doc.add_paragraph()
    mr = meta.add_run(
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Engine: OmniCite-IEEE-v1"
    )
    _set_run_font(mr, size_pt=9, italic=True)
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # ========================================================
    # 1. SUMMARY
    # ========================================================
    h1 = doc.add_heading(level=1)
    hr = h1.add_run("1. Summary")
    _set_run_font(hr, size_pt=14, bold=True)

    window_start = recency.get("start_year", manuscript_year - 9)
    window_end = recency.get("end_year", manuscript_year)
    recent = recency.get("recent_count", 0)
    recent_pct = recency.get("recent_percentage", 0)

    # IEEE never reconstructs references from OpenAlex
    doi_reconstructed = 0

    summary_rows = [
        ("Total References", str(total_refs), False),
        ("References > 15",
         f"Yes ({total_refs})" if total_refs > 15 else f"No ({total_refs})",
         total_refs > 15),
        ("Total Unique In-text Citations",
         str(stats.get("unique", 0)), False),
        ("Collapse Corrected",
         str(stats.get("crowded_clusters", 0)),
         stats.get("crowded_clusters", 0) > 0),
        (f"% Last 10 Years ({window_start}–{window_end})",
         f"{recent_pct:.1f}% ({recent}/{total_refs})"
         if total_refs else "0.0%", False),
        ("Citations Missing from References",
         str(len(orphan)), len(orphan) > 0),
        ("References Missing from Citations",
         str(len(uncited)), len(uncited) > 0),
        ("DOI Checked", str(doi_checked), False),
        ("References Reconstructed", str(doi_reconstructed), False),
        ("DOI Suspicious (possible fabrication)",
         f"{doi_suspicious} "
         f"({doi_suspicious / total_refs * 100:.1f}%)"
         if total_refs else "0",
         doi_suspicious > 0),
    ]

    summary_table = doc.add_table(rows=1, cols=2)
    summary_table.style = "Light Grid Accent 1"
    hdr = summary_table.rows[0].cells
    for cell, text in zip(hdr, ["Metric", "Value"]):
        cell.text = ""
        _set_run_font(cell.paragraphs[0].add_run(text), size_pt=10, bold=True)

    for label, value, warn in summary_rows:
        cells = summary_table.add_row().cells
        cells[0].text = ""
        cells[1].text = ""
        _set_run_font(
            cells[0].paragraphs[0].add_run(label),
            size_pt=10, bold=warn, color=RED if warn else None,
        )
        _set_run_font(
            cells[1].paragraphs[0].add_run(value),
            size_pt=10, bold=warn, color=RED if warn else None,
        )

    doc.add_paragraph()

    # ========================================================
    # 1.1 SOURCE TYPE DISTRIBUTION
    # ========================================================
    h11 = doc.add_heading(level=2)
    hr11 = h11.add_run("1.1 Source Type Distribution")
    _set_run_font(hr11, size_pt=12, bold=True)

    src_counts = Counter((r.get("Source Type") or "Other") for r in reference_rows)
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Light Grid Accent 1"
    for cell, text in zip(tbl.rows[0].cells,
                          ["Source Type", "Count", "Percentage"]):
        cell.text = ""
        _set_run_font(cell.paragraphs[0].add_run(text), size_pt=10, bold=True)

    for stype in CANONICAL_SOURCE_TYPES:
        cnt = src_counts.get(stype, 0)
        pct = (cnt / total_refs * 100) if total_refs else 0
        cells = tbl.add_row().cells
        for cell, val in zip(cells, [stype, str(cnt), f"{pct:.1f}%"]):
            cell.text = ""
            _set_run_font(cell.paragraphs[0].add_run(val), size_pt=10)

    cells = tbl.add_row().cells
    for cell, val in zip(cells, ["Total", str(total_refs), "100.0%"]):
        cell.text = ""
        _set_run_font(cell.paragraphs[0].add_run(val), size_pt=10, bold=True)

    doc.add_page_break()

    # ========================================================
    # 2. IN-TEXT CITATION CORRECTION
    # ========================================================
    h2 = doc.add_heading(level=1)
    hr2 = h2.add_run("2. In-text Citation Correction")
    _set_run_font(hr2, size_pt=14, bold=True)

    caption = doc.add_paragraph()
    cr = caption.add_run(
        "Format: original citation — corrected citation — note"
    )
    _set_run_font(cr, size_pt=9, italic=True)

    ref_numbers = set(range(1, total_refs + 1))

    if not cluster_rows:
        p = doc.add_paragraph()
        _set_run_font(
            p.add_run("No bracketed in-text citations were detected."),
            italic=True,
        )
    else:
        for i, row in enumerate(cluster_rows, start=1):
            numbers = [int(n) for n in re.findall(r"\d+", row.get("Numbers Cited", ""))]
            has_orphan = any(n not in ref_numbers for n in numbers)
            is_revised = row.get("Status") == "REVISED"

            # Compact header — no extra spacers; the divider below does the separating.
            head = doc.add_paragraph()
            head.paragraph_format.space_before = Pt(4)
            head.paragraph_format.space_after = Pt(1)

            hrun = head.add_run(f"{i}. ")
            _set_run_font(hrun, size_pt=11, bold=True)

            type_run = head.add_run("IN-TEXT CITATIONS")
            _set_run_font(type_run, size_pt=10, italic=True)

            if has_orphan:
                _add_run(head, "  [NOT IN REFERENCES]",
                         size_pt=10, bold=True, italic=True, red=True)

            p_orig = doc.add_paragraph()
            p_orig.paragraph_format.left_indent = Inches(0.25)
            p_orig.paragraph_format.space_after = Pt(1)
            _add_run(p_orig, "Original:  ", bold=True, size_pt=11)
            if is_revised or has_orphan:
                _add_run(p_orig, row.get("Original Form", ""),
                         size_pt=11, red=True)
            else:
                _add_run(p_orig, row.get("Original Form", ""), size_pt=11)

            p_corr = doc.add_paragraph()
            p_corr.paragraph_format.left_indent = Inches(0.25)
            p_corr.paragraph_format.space_after = Pt(1)
            _add_run(p_corr, "Corrected: ", bold=True, size_pt=11)
            _add_run(p_corr, row.get("Corrected Form", ""), size_pt=11)

            note = (row.get("Reason") or "").strip()
            if is_revised and note:
                note_text = note
            elif is_revised:
                note_text = "Citation cluster rewritten to IEEE canonical form."
            elif has_orphan:
                note_text = "Citation number has no matching reference entry; manual check needed."
            else:
                note_text = "No changes needed; citation is correct."

            np = doc.add_paragraph()
            np.paragraph_format.left_indent = Inches(0.25)
            np.paragraph_format.space_after = Pt(2)
            np.paragraph_format.line_spacing = 1.0
            _add_run(np, f"Note: {note_text}", size_pt=10, italic=True,
                     red=(is_revised or has_orphan))

            _add_divider(doc)

    doc.add_page_break()

    # ========================================================
    # 3. REFERENCE LIST (IEEE STYLE)
    # ========================================================
    h3 = doc.add_heading(level=1)
    hr3 = h3.add_run("3. Reference List (IEEE Style)")
    _set_run_font(hr3, size_pt=14, bold=True)

    caption = doc.add_paragraph()
    cr2 = caption.add_run(
        "Format: original reference — corrected reference — comment — status"
    )
    _set_run_font(cr2, size_pt=9, italic=True)

    cited_numbers = {row.get("Reference #") for row in matching if row.get("Cited")}

    if not reference_rows:
        p = doc.add_paragraph()
        _set_run_font(p.add_run("No references were detected."), italic=True)
    else:
        for i, row in enumerate(reference_rows, start=1):
            ref_no = row.get("No.", i)
            not_cited = ref_no not in cited_numbers
            withheld = _ref_is_withheld(row)
            suspicious = _ref_is_suspicious(row)
            manual = _ref_is_manual(row)
            has_placeholder = _ref_has_placeholder(row)

            head = doc.add_paragraph()
            head.paragraph_format.space_before = Pt(6)
            head.paragraph_format.space_after = Pt(2)

            hrun = head.add_run(f"{ref_no}.")
            _set_run_font(hrun, size_pt=11, bold=True)

            if not_cited:
                _add_run(head, "  [NOT CITED IN TEXT]",
                         size_pt=10, bold=True, italic=True, red=True)

            p_type = doc.add_paragraph()
            p_type.paragraph_format.left_indent = Inches(0.25)
            p_type.paragraph_format.space_after = Pt(2)
            _add_run(p_type, "Source Type: ", bold=True, size_pt=11)
            _add_run(p_type, row.get("Source Type", "Other"), size_pt=11)

            p_orig = doc.add_paragraph()
            p_orig.paragraph_format.left_indent = Inches(0.25)
            p_orig.paragraph_format.space_after = Pt(2)
            p_orig.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            _add_run(p_orig, "Original: ", bold=True, size_pt=11)
            original = row.get("Original Reference", "")
            original_has_error = bool(row.get("Original Has Structure Error"))
            _add_run(p_orig, original, size_pt=11,
                     red=(withheld or suspicious or original_has_error))

            if withheld or suspicious:
                p_w = doc.add_paragraph()
                p_w.paragraph_format.left_indent = Inches(0.25)
                p_w.paragraph_format.space_after = Pt(2)
                _add_run(p_w, "Corrected: ", bold=True, size_pt=11)
                withheld_label = (
                    "— WITHHELD (DOI mismatch) —" if suspicious else "— WITHHELD —"
                )
                _add_run(p_w, withheld_label, size_pt=11, bold=True, red=True)
            else:
                corrected = row.get("Corrected Version", "").strip()
                if corrected:
                    p_corr = doc.add_paragraph()
                    p_corr.paragraph_format.left_indent = Inches(0.25)
                    p_corr.paragraph_format.space_after = Pt(2)
                    p_corr.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    _add_run(p_corr, "Corrected: ", bold=True, size_pt=11)

                    italic_elements = row.get("Italic Tokens") or row.get("Italicized in IEEE", "")

                    missing = row.get("Missing Required Elements", "")
                    missing_tokens = [m.strip() for m in missing.split(",") if m.strip()]
                    norm_missing = set()
                    for tok in missing_tokens:
                        t = tok.lower().strip().rstrip(".")
                        if t in {"vol", "volume"}:
                            norm_missing.add("vol.")
                        elif t in {"no", "issue", "number"}:
                            norm_missing.add("no.")
                        elif t in {"pp", "pages", "page"}:
                            norm_missing.add("pp.")
                        elif t in {"doi", "url", "doi / url"}:
                            norm_missing.add("doi")
                            norm_missing.add("url")

                    _add_ieee_reference_with_italics(
                        p_corr, corrected, italic_elements,
                        missing_tokens=sorted(norm_missing), size_pt=11
                    )

            p_comment = doc.add_paragraph()
            p_comment.paragraph_format.left_indent = Inches(0.25)
            p_comment.paragraph_format.space_after = Pt(2)
            _add_run(p_comment, "Comment: ", bold=True, size_pt=11)

            comment_parts = []
            if suspicious:
                reasons = row.get("DOI Verification Reasons", "").strip()
                comment_parts.append(
                    f"DOI mismatch — possible fabricated reference. {reasons}"
                )
                oa_title = row.get("OpenAlex Title")
                if oa_title:
                    comment_parts.append(f'OpenAlex title: "{oa_title}"')
            elif withheld:
                comment_parts.append(
                    "DOI not provided — automated verification/correction withheld."
                )
            elif manual:
                ai_exp = (row.get("AI Explanation") or "").strip()
                comment_parts.append(
                    ai_exp or "Manual verification required for this source type."
                )
            elif has_placeholder:
                fixes = (row.get("Correction Note") or "").strip()
                comment_parts.append(
                    fixes or "Placeholder fields inserted — fill in before submission."
                )
            else:
                fixes = (row.get("Correction Note") or "").strip()
                if fixes:
                    comment_parts.append(fixes)
                else:
                    comment_parts.append("Reference is consistent with IEEE style.")

            comment_text = " | ".join(comment_parts)
            _add_run(p_comment, comment_text, size_pt=10, italic=True,
                     red=(withheld or suspicious or manual))

            if has_placeholder and not (withheld or suspicious):
                ph = row.get("Placeholders") or {}
                active = [k for k, v in ph.items() if v]
                p_ph = doc.add_paragraph()
                p_ph.paragraph_format.left_indent = Inches(0.25)
                p_ph.paragraph_format.space_after = Pt(2)
                _add_run(
                    p_ph,
                    f"Placeholders to fill: {', '.join(active)}",
                    size_pt=10, italic=True, bold=True, red=True,
                )

            p_status = doc.add_paragraph()
            p_status.paragraph_format.left_indent = Inches(0.25)
            p_status.paragraph_format.space_after = Pt(4)
            _add_run(p_status, "Status: ", bold=True, size_pt=11)

            status = str(row.get("Status", "MANUAL CHECK")).upper().strip()
            if status not in {"MATCH", "REVISED", "WITHHELD", "MANUAL CHECK"}:
                status = "MANUAL CHECK"

            status_red = status in {"WITHHELD", "MANUAL CHECK"}
            _add_run(p_status, status, size_pt=11, bold=True, red=status_red)

            _add_divider(doc)

    # ---- Save ----
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


# =========================================================
# RENDER — called from app.py
# =========================================================

def render():
    st.title("OmniCite Auditor - IEEE Style")

    # ---- ALL button styling injected BEFORE any widget renders ----
    st.markdown(
        """
        <style>
        /* Blue Extract & Review */
        div[class*="st-key-ieee_extract_btn"] button {
            background-color: #2563eb !important;
            color: #ffffff !important;
            border: 1px solid #1d4ed8 !important;
            font-weight: 600 !important;
            transition: background-color 0.15s ease;
        }
        div[class*="st-key-ieee_extract_btn"] button:hover {
            background-color: #1d4ed8 !important;
            color: #ffffff !important;
            border-color: #1e40af !important;
        }
        div[class*="st-key-ieee_extract_btn"] button:focus {
            box-shadow: 0 0 0 0.2rem rgba(37, 99, 235, 0.4) !important;
        }
        /* Green Download */
        div[class*="st-key-ieee_download_btn"] button {
            background-color: #16a34a !important;
            color: #ffffff !important;
            border: 1px solid #15803d !important;
            font-weight: 600 !important;
            transition: background-color 0.15s ease;
        }
        div[class*="st-key-ieee_download_btn"] button:hover {
            background-color: #15803d !important;
            color: #ffffff !important;
            border-color: #166534 !important;
        }
        div[class*="st-key-ieee_download_btn"] button:focus {
            box-shadow: 0 0 0 0.2rem rgba(22, 163, 74, 0.4) !important;
        }
        /* Red Start Fresh */
        div[class*="st-key-ieee_reset_btn"] button {
            background-color: #dc2626 !important;
            color: #ffffff !important;
            border: 1px solid #b91c1c !important;
            font-weight: 600 !important;
            transition: background-color 0.15s ease;
        }
        div[class*="st-key-ieee_reset_btn"] button:hover {
            background-color: #b91c1c !important;
            color: #ffffff !important;
            border-color: #991b1b !important;
        }
        div[class*="st-key-ieee_reset_btn"] button:focus {
            box-shadow: 0 0 0 0.2rem rgba(220, 38, 38, 0.4) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    openai_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not openai_key:
        st.warning("OPENAI_API_KEY not set — AI correction will be skipped. Local checks still run.")

    client = None
    if openai_key:
        try:
            from openai import OpenAI
            client = OpenAI(api_key=openai_key)
        except Exception:
            client = None

    if "ieee_uploader_version" not in st.session_state:
        st.session_state["ieee_uploader_version"] = 0

    uploaded_files = st.file_uploader(
        "Upload manuscript PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload up to 5 manuscripts in one batch.",
        key=f"ieee_uploader_{st.session_state['ieee_uploader_version']}",
    )

    if "ieee_batches" not in st.session_state:
        st.session_state["ieee_batches"] = {}

    if not uploaded_files:
        return

    if len(uploaded_files) > 5:
        st.error(f"You uploaded {len(uploaded_files)} manuscripts. Maximum batch size is 5.")
        st.stop()

    current_year = datetime.now().year
    default_year = st.session_state.get("ieee_manuscript_year", current_year)
    manuscript_year = st.number_input(
        "Manuscript publication year",
        min_value=1900, max_value=current_year + 5,
        value=int(default_year), step=1,
        help="Used for the % of references within the last 10 years [year-9, year].",
        key="ieee_manuscript_year_input",
    )
    st.session_state["ieee_manuscript_year"] = int(manuscript_year)

    file_keys = [(f"{i}::{uf.name}", uf) for i, uf in enumerate(uploaded_files)]

    active_keys = {k for k, _ in file_keys}
    for k in list(st.session_state["ieee_batches"].keys()):
        if k not in active_keys:
            del st.session_state["ieee_batches"][k]

    with st.container(key="ieee_extract_btn"):
        extract_clicked = st.button(
            "Extract & Review",
            type="primary",
            use_container_width=True,
            key="ieee_btn_run_all_pdfs",
        )

    if extract_clicked:
        overall = st.progress(0, text="Starting...")
        n = len(file_keys)
        step = 100 / max(n, 1)

        for i, (key, uf) in enumerate(file_keys, start=1):
            base_pct = int((i - 1) * step)

            # ---- Stage 1: extract PDF ----
            overall.progress(
                base_pct + int(step * 0.10),
                text=f"[{i}/{n}] Extracting {uf.name}...",
            )
            batch = {
                "filename": uf.name,
                "manuscript_year": int(manuscript_year),
                "ai_done": False,
            }
            try:
                uf.seek(0)
                extract_ieee_pdf(uf, batch, int(manuscript_year))
            except Exception as exc:
                st.error(f"{uf.name} extraction failed: {exc}")
                batch["error"] = str(exc)
                st.session_state["ieee_batches"][key] = batch
                continue

            # ---- Stage 2: extract references & verify DOIs ----
            overall.progress(
                base_pct + int(step * 0.35),
                text=f"[{i}/{n}] Extracting references & verifying DOIs...",
            )
            try:
                process_ieee_references_and_citations(
                    batch, client, int(manuscript_year)
                )
            except Exception as exc:
                st.error(f"{uf.name} IEEE check failed: {exc}")
                batch["error"] = str(exc)

            st.session_state["ieee_batches"][key] = batch

            # ---- Stage 3: done ----
            overall.progress(
                base_pct + int(step * 1.00),
                text=f"[{i}/{n}] {uf.name} done.",
            )

        overall.progress(100, text="All manuscripts processed.")
        overall.empty()

    any_done = any(b.get("ai_done") for b in st.session_state["ieee_batches"].values())
    if not any_done:
        return

    selector_options = [k for k, _ in file_keys]

    def _fmt(k):
        b = st.session_state["ieee_batches"].get(k, {})
        suffix = "" if b.get("ai_done") else "  (not yet processed)"
        return b.get("filename", k) + suffix

    selected_key = st.selectbox(
        "Select manuscript to review",
        selector_options,
        format_func=_fmt,
        key="ieee_selected_key",
    )

    batch = st.session_state["ieee_batches"].get(selected_key)
    if not batch or not batch.get("ai_done"):
        st.info("This manuscript has not been processed yet. Click 'Extract & Review' above.")
        return

    batch["manuscript_year"] = int(st.session_state.get("ieee_manuscript_year", current_year))

    # ---- Debug: show the reference section extracted from the PDF ----
    with st.expander("View reference section", expanded=False):
        numbered_slice = "\n\n".join(
            f"[{i}] " + (
                "Cannot interpret reference from extracted PDF."
                if _is_uninterpretable_reference(ref) else ref
            )
            for i, ref in enumerate(batch.get("references", []), start=1)
        )
        st.text_area(
            "Reference slice",
            numbered_slice,
            height=300,
            key=f"dbg_ref_{selected_key}",
        )
        with st.expander("Raw extracted reference text", expanded=False):
            st.text_area(
                "Raw reference section",
                batch.get("reference_text", ""),
                height=220,
                key=f"dbg_raw_ref_{selected_key}",
            )

    # ---- Extraction quality warning ----
    quality = batch.get("reference_extraction_quality") or {}
    if not quality.get("complete", True):
        expected = quality.get("expected_count")
        actual = quality.get("actual_count")
        st.warning(
            f"⚠️ **Reference extraction may be incomplete.** "
            f"The manuscript appears to contain **{expected}** references "
            f"(based on the highest `[n]` marker found), but only "
            f"**{actual}** were successfully extracted. "
            f"Please verify the reference list manually."
        )

    reference_rows = batch.get("reference_comparison", [])
    total_refs_now = len(batch.get("references", []))

    # ========================================================
    # BUTTON STYLING — green download, red reset
    # ========================================================
    st.markdown(
        """
        <style>
        div[class*="st-key-ieee_download_btn"] button {
            background-color: #16a34a !important;
            color: #ffffff !important;
            border: 1px solid #15803d !important;
            font-weight: 600 !important;
            transition: background-color 0.15s ease;
        }
        div[class*="st-key-ieee_download_btn"] button:hover {
            background-color: #15803d !important;
            color: #ffffff !important;
            border-color: #166534 !important;
        }
        div[class*="st-key-ieee_download_btn"] button:focus {
            box-shadow: 0 0 0 0.2rem rgba(22, 163, 74, 0.4) !important;
        }
        div[class*="st-key-ieee_reset_btn"] button {
            background-color: #dc2626 !important;
            color: #ffffff !important;
            border: 1px solid #b91c1c !important;
            font-weight: 600 !important;
            transition: background-color 0.15s ease;
        }
        div[class*="st-key-ieee_reset_btn"] button:hover {
            background-color: #b91c1c !important;
            color: #ffffff !important;
            border-color: #991b1b !important;
        }
        div[class*="st-key-ieee_reset_btn"] button:focus {
            box-shadow: 0 0 0 0.2rem rgba(220, 38, 38, 0.4) !important;
        }
        div[class*="st-key-ieee_extract_btn"] button {
            background-color: #2563eb !important;
            color: #ffffff !important;
            border: 1px solid #1d4ed8 !important;
            font-weight: 600 !important;
            transition: background-color 0.15s ease;
        }
        div[class*="st-key-ieee_extract_btn"] button:hover {
            background-color: #1d4ed8 !important;
            color: #ffffff !important;
            border-color: #1e40af !important;
        }
        div[class*="st-key-ieee_extract_btn"] button:focus {
            box-shadow: 0 0 0 0.2rem rgba(37, 99, 235, 0.4) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ========================================================
    # METRICS TABLE (compact, matches APA structure)
    # ========================================================
    stats = batch["citation_stats"]
    recency = batch["recency"]
    orphan = batch["orphan_citations"]
    uncited = batch["uncited_references"]

    doi_checked = sum(1 for r in reference_rows if r.get("DOI Verified"))
    doi_suspicious = sum(1 for r in reference_rows if r.get("DOI Suspicious"))
    doi_suspicious_pct = (
        doi_suspicious / total_refs_now * 100 if total_refs_now else 0
    )
    withheld_count = sum(
        1 for r in reference_rows if _ref_is_withheld(r)
    )
    manual_count = sum(
        1 for r in reference_rows if _ref_is_manual(r)
    )
    placeholder_count = sum(
        1 for r in reference_rows if _ref_has_placeholder(r)
    )

    window_start = recency.get("start_year", int(manuscript_year) - 9)
    window_end = recency.get("end_year", int(manuscript_year))
    recent = recency.get("recent_count", 0)
    recent_pct = recency.get("recent_percentage", 0)

    # IEEE never reconstructs references from OpenAlex
    doi_reconstructed = 0

    metrics = [
        ("Total References", total_refs_now),
        ("References > 15",
         f"Yes ({total_refs_now})" if total_refs_now > 15
         else f"No ({total_refs_now})"),
        ("Total Unique In-text Citations", stats.get("unique", 0)),
        ("Collapse Corrected", stats.get("crowded_clusters", 0)),
        (f"% Last 10 Years ({window_start}–{window_end})",
         f"{recent_pct:.1f}% ({recent}/{total_refs_now})"
         if total_refs_now else "0.0%"),
        ("Citations Missing from References", len(orphan)),
        ("References Missing from Citations", len(uncited)),
        ("DOI Checked", doi_checked),
        ("References Reconstructed", doi_reconstructed),
        ("DOI Suspicious (possible fabrication)",
         f"{doi_suspicious} "
         f"({doi_suspicious / total_refs_now * 100:.1f}%)"
         if total_refs_now else "0"),
    ]

    metric_df = pd.DataFrame(metrics, columns=["Metric", "Value"])

    source_counts = Counter(
        (row.get("Source Type") or "Other") for row in reference_rows
    )
    source_rows = []
    for source_type in CANONICAL_SOURCE_TYPES:
        count = source_counts.get(source_type, 0)
        pct = count / total_refs_now * 100 if total_refs_now else 0
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

    # ========================================================
    # DOWNLOAD (green)
    # ========================================================
    try:
        docx_bytes = build_ieee_docx(batch)
        safe_name = re.sub(r"[^\w\-]+", "_", batch.get("filename", "manuscript"))
        with st.container(key="ieee_download_btn"):
            st.download_button(
                label="📄 Download Diagnostic Report (.docx)",
                data=docx_bytes,
                file_name=f"{safe_name}_IEEE_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"download_ieee_report_docx_{selected_key}",
            )
    except Exception as exc:
        st.error(f"Could not build DOCX report: {exc}")

    # ========================================================
    # START FRESH (red)
    # ========================================================
    with st.container(key="ieee_reset_btn"):
        if st.button(
            "🔄 Start Fresh — Clear All Uploads & Results",
            use_container_width=True,
            key="ieee_reset_btn_start_fresh",
        ):
            for k in list(st.session_state.keys()):
                if k == "ieee_batches" or k.startswith("ieee_batches"):
                    del st.session_state[k]
                if k.startswith("dbg_ref_"):
                    del st.session_state[k]
                if k == "ieee_selected_key":
                    del st.session_state[k]
                if k == "ieee_show_extra":
                    del st.session_state[k]

            st.session_state["ieee_batches"] = {}
            st.session_state["ieee_uploader_version"] = (
                st.session_state.get("ieee_uploader_version", 0) + 1
            )
            st.session_state["ieee_manuscript_year"] = datetime.now().year
            st.rerun()