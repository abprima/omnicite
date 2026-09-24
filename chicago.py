# chicago.py
# OmniCite Auditor — Chicago Notes & Bibliography Module (APA-flow adaptation)
# Called from app.py via:  import chicago; chicago.render()

import os
import io
import re
import json
import difflib
from collections import Counter
from datetime import datetime
from pathlib import Path
from io import BytesIO
from typing import Optional

import streamlit as st
import pandas as pd
import requests
from markitdown import MarkItDown
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pydantic import BaseModel

from openalex_config import get_openalex_api_key


# ============================================================
# HEADINGS
# ============================================================

NOTE_HEADINGS = {
    "notes", "note", "footnotes", "footnote",
    "endnotes", "endnote", "catatan", "catatan kaki", "catatan akhir",
}

REFERENCE_HEADINGS = {
    "bibliography", "bibliographies",
    "references", "reference", "reference list",
    "daftar pustaka", "daftar rujukan", "rujukan",
    "bibliografi", "referensi", "sumber",
}

POST_REFERENCE_HEADINGS = {
    "acknowledgement", "acknowledgements", "acknowledgment", "acknowledgments",
    "author contribution", "author contributions", "contribution", "contributions",
    "conflict of interest", "conflicts of interest", "declaration", "declarations",
    "funding", "funding information", "funding statement",
    "data availability", "data availability statement",
    "ethical approval", "ethics approval", "ethics statement",
    "informed consent", "consent for publication",
    "appendix", "appendices", "supplementary material", "supplementary materials",
    "author note", "author notes",
    "ucapan terima kasih", "ucapan terimakasih",
    "kontribusi penulis", "pernyataan kontribusi penulis",
    "profil penulis", "biografi penulis",
    "konflik kepentingan", "pernyataan konflik kepentingan",
    "pendanaan", "pernyataan pendanaan", "ketersediaan data",
    "persetujuan etik", "lampiran", "catatan penulis",
}


def _normalize_heading(line: str) -> str:
    cleaned = line.strip()
    cleaned = re.sub(r"^[#>*\-\s]+", "", cleaned)
    cleaned = cleaned.strip("*_` ")
    cleaned = re.sub(r"\s+", " ", cleaned).lower()
    cleaned = cleaned.rstrip(":. ")
    return cleaned


def _is_plausible_reference_heading(line: str) -> bool:
    stripped = line.strip()
    if not stripped or len(stripped.split()) > 4:
        return False
    if stripped.endswith((".", "?", "!", ",", ";")):
        return False
    return _normalize_heading(stripped) in REFERENCE_HEADINGS


def _is_plausible_note_heading(line: str) -> bool:
    stripped = line.strip()
    if not stripped or len(stripped.split()) > 4:
        return False
    if stripped.endswith((".", "?", "!", ",", ";")):
        return False
    return _normalize_heading(stripped) in NOTE_HEADINGS


# ============================================================
# RUNNING HEADER / FOOTER STRIPPER
# ============================================================

REPEAT_THRESHOLD = 3
MAX_BOILERPLATE_LEN = 120

BOILERPLATE_PATTERNS = [
    r"^is licensed under a .*$",
    r"^creative commons.*$",
    r"^cc[- ]by.*$",
    r"^doi:\s*10\.\S+$",
    r"^~?\s*\d+\s*~\s*\d+\(\d+\),\s*\d+[-\u2013]\d+\s*$",
    r"^\d+\s*~\s*\d+\(\d+\),\s*\d+[-\u2013]\d+\s*$",
    r"^p?e?ISSN\s*\d+[-\u2013]\d+.*$",
    r"^terakreditasi.*$",
    r"^\*?email koresponden.*$",
    r"^copyright \u00a9.*$",
    r"^received:.*accepted:.*$",
    r"^submitted:.*published:.*$",
    r"^diterima:.*disetujui:.*$",
]

BOILERPLATE_RE = re.compile(
    "|".join(f"(?:{p})" for p in BOILERPLATE_PATTERNS),
    re.IGNORECASE,
)


def strip_running_headers_footers(text: str) -> str:
    lines = text.splitlines()
    counter = Counter()
    normalized = []

    for raw in lines:
        stripped = raw.strip()
        if not stripped or len(stripped) > MAX_BOILERPLATE_LEN:
            normalized.append(None)
            continue
        norm = re.sub(r"\s+", " ", stripped).lower()
        normalized.append(norm)
        counter[norm] += 1

    repeated = {n for n, c in counter.items() if c >= REPEAT_THRESHOLD}

    kept = []
    for raw, norm in zip(lines, normalized):
        stripped = raw.strip()
        if not stripped:
            kept.append(raw)
            continue
        if BOILERPLATE_RE.match(stripped):
            continue
        if norm is not None and norm in repeated:
            continue
        kept.append(raw)

    return "\n".join(kept)


# ============================================================
# CLEANING
# ============================================================

def clean_text(text):
    if not text:
        return ""
    text = text.replace("\u00ad", "")
    text = text.replace("\u2010", "-").replace("\u2011", "-")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# ============================================================
# SLICERS
# ============================================================

# A reference-start signature used for both the quality gate and the
# glued-line splitter: "Surname, A." or "Surname, Given" preceded by a
# sentence-ending period and whitespace.
_REF_START_INLINE_RE = re.compile(
    r"(?<=\.)\s+(?=[A-Z\u00c0-\u00d6\u00d8-\u00dd]"
    r"[A-Za-z\u00c0-\u00ff'\u2019\-]+,\s+[A-Z])"
)


def _count_reference_signatures(text: str) -> int:
    """
    Count how many *reference-shaped* fragments appear in `text`.

    A fragment is reference-shaped if it either:
      - contains a parenthetical year like (2022) or (2019a), OR
      - starts with a DOI/URL, OR
      - starts with a 'Surname, A.' author pattern.

    This is used both for the section quality gate and for the
    glued-line splitter so that a single physical line containing
    multiple references is counted as multiple references.
    """
    if not text:
        return 0

    # Split on reference-start boundaries first.
    pieces = re.split(_REF_START_INLINE_RE, text)
    if len(pieces) <= 1:
        pieces = [text]

    count = 0
    for piece in pieces:
        p = piece.strip()
        if not p or len(p) < 15:
            continue
        if re.match(r"^\s*(Table|Figure|Fig\.|Tab\.)\s+\d+", p, re.I):
            continue
        if re.search(r"\((?:19|20)\d{2}[a-z]?\)", p):
            count += 1
            continue
        if re.match(r"^(https?://|10\.\d{4,9}/)", p):
            count += 1
            continue
        if _looks_like_reference_start(p):
            count += 1
    return count


def _quality_of_reference_block(block: str, cap: int = 300) -> int:
    """Count reference-shaped entries in the first `cap` lines."""
    count = 0
    for line in block.splitlines()[:cap]:
        count += _count_reference_signatures(line)
    return count


def slice_reference_section(text: str):
    lines = text.splitlines()
    candidates = [i for i, l in enumerate(lines) if _is_plausible_reference_heading(l)]
    if not candidates:
        return "", False, False

    for ref_index in reversed(candidates):
        post_index = None
        non_empty_seen = 0
        for j in range(ref_index + 1, len(lines)):
            stripped = lines[j].strip()
            if not stripped:
                continue
            non_empty_seen += 1
            if non_empty_seen >= 3 and _normalize_heading(stripped) in POST_REFERENCE_HEADINGS:
                post_index = j
                break

        end = post_index if post_index is not None else len(lines)
        block = "\n".join(lines[ref_index:end]).strip()
        if _quality_of_reference_block(block, cap=300) >= 5:
            return block, True, (post_index is not None)

    return "", False, False


def slice_body_section(text: str):
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if _is_plausible_reference_heading(line):
            return "\n".join(lines[:i]).strip(), True
    return text, False


def slice_note_section(text: str):
    lines = text.splitlines()
    note_start = None
    for i, line in enumerate(lines):
        if _is_plausible_note_heading(line):
            note_start = i
            break
    if note_start is None:
        return "", False

    for j in range(note_start + 1, len(lines)):
        if _is_plausible_reference_heading(lines[j]):
            return "\n".join(lines[note_start + 1:j]).strip(), True

    return "\n".join(lines[note_start + 1:]).strip(), True


# ============================================================
# CHICAGO FOOTNOTE EXTRACTION (Markdown-based)
# ============================================================

_NOTE_START_RE = re.compile(r"^\s*(\d{1,3})[.)]\s+(.*)$")


def extract_chicago_footnotes(text: str):
    note_block, found_block = slice_note_section(text)
    source_text = note_block if found_block else text

    footnotes = []
    current = None
    expected = None

    for raw_line in source_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        m = _NOTE_START_RE.match(line)
        if m:
            number = int(m.group(1))
            body = clean_text(m.group(2))

            if expected is None and number == 1:
                pass
            elif expected is not None and number == expected:
                pass
            elif found_block and number == 1 and not footnotes:
                pass
            else:
                if current is not None:
                    current["text"] += " " + line
                    current["line_count"] += 1
                continue

            if current is not None:
                footnotes.append(current)
            current = {
                "number": number,
                "page": None,
                "end_page": None,
                "y0": 0.0,
                "text": body,
                "line_count": 1,
                "cross_page": False,
                "layout_warning": False,
                "layout_warning_message": "",
            }
            expected = number + 1
            continue

        if current is not None:
            current["text"] += " " + line
            current["line_count"] += 1

    if current is not None:
        footnotes.append(current)

    footnotes = [
        n for n in footnotes
        if clean_text(n["text"]) and len(clean_text(n["text"])) >= 5
    ]

    return {
        "footnotes": footnotes,
        "count": len(footnotes),
        "bibliography_page": None,
        "debug": [],
        "page_candidates": [],
    }


def attach_page_numbers_to_footnotes(footnotes, page_candidates=None):
    out = []
    for n in footnotes:
        item = dict(n)
        item.setdefault("pages", [])
        out.append(item)
    return out


def footnote_has_indonesian_author_conjunction(text):
    return bool(re.search(r"\s+dan\s+", clean_text(text), flags=re.I))


# ============================================================
# BIBLIOGRAPHY REFERENCE SPLITTING  (glued-line aware)
# ============================================================

_URL_OR_DOI_RE = re.compile(
    r"^(?:https?://|www\.|10\.\d{4,9}/|doi\s*:)", re.I
)

# Personal-author start: "Surname, A." or "Surname, Given"
_AUTHOR_START_RE = re.compile(
    r"^[A-Z\u00c0-\u00d6\u00d8-\u00dd]"
    r"[A-Za-z\u00c0-\u00ff'\u2019\-]+"
    r",\s+[A-Z]"
)

# Corporate-author start: "Some Organization Name."
_CORPORATE_START_RE = re.compile(
    r"^[A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z]+){1,6}\."
)

# Title-in-quotes start
_QUOTED_START_RE = re.compile(r"^[\u201c\"]")

# Numbered reference start: "1. "
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
    # Corporate start — but reject if it's just a journal name wrap
    m = _CORPORATE_START_RE.match(s)
    if m:
        # Reject if the "corporate name" is a common journal keyword
        # (WSEAS Transactions on ..., Journal of ..., Review of ...)
        first_two = " ".join(m.group(0).split()[:2]).rstrip(".")
        if re.match(
            r"^(?:WSEAS|Journal|Jurnal|Review|International|Proceedings|"
            r"Transactions|Bulletin|Studies|Research)\b",
            first_two, re.I,
        ):
            return False
        return True
    return False


def _split_glued_line(text: str) -> list[str]:
    """
    Split one physical line that may contain multiple glued references.

    Uses the author-start boundary `.<space><Surname>, <Initial>` plus
    the corporate-author boundary `.<space>Some Name.` while never
    splitting inside a URL or DOI.
    """
    if not text:
        return []

    # Fast path: single-reference line
    if not re.search(r"\.\s+[A-Z\u00c0-\u00d6\u00d8-\u00dd]", text):
        return [text.strip()]

    # Candidate boundaries: a sentence-ending period, whitespace, then
    # a capital letter. We then verify the fragment after the boundary
    # looks like a reference start.
    candidates = []
    for m in re.finditer(r"(?<=\.)\s+(?=[A-Z\u00c0-\u00d6\u00d8-\u00dd])", text):
        candidates.append(m.start())

    if not candidates:
        return [text.strip()]

    boundaries = [0]
    for pos in candidates:
        after = text[pos:].lstrip()
        # Do not split if the char just before the period is part of a URL
        before = text[:pos]
        if re.search(r"https?://[^\s]*$", before):
            continue
        # Do not split if the fragment after boundary is a URL continuation
        if _URL_OR_DOI_RE.match(after):
            continue
        if _looks_like_reference_start(after):
            boundaries.append(pos)
    boundaries.append(len(text))

    pieces = []
    for i in range(len(boundaries) - 1):
        start = boundaries[i]
        end = boundaries[i + 1]
        piece = text[start:end].strip()
        if piece:
            pieces.append(piece)

    # Post-merge: if a piece does not actually begin a reference, glue
    # it back to the previous one.
    merged = []
    for piece in pieces:
        if not merged:
            merged.append(piece)
            continue
        if _looks_like_reference_start(piece):
            merged.append(piece)
        else:
            merged[-1] = merged[-1] + " " + piece
    return merged


def split_references_from_lines(lines):
    """
    Split bibliography lines into individual references.

    Handles:
      A. Two-column layout where lines carry a `column` key.
      B. One-column layout where multiple references are glued on the
         same physical line.
    """
    if not lines:
        return []

    # ---- Normalise input ----
    normalised = []
    for item in lines:
        if isinstance(item, dict):
            normalised.append({
                "text": item.get("text", ""),
                "x0": float(item.get("x0", 0.0)),
                "y0": float(item.get("y0", 0.0)),
                "column": item.get("column", "LEFT"),
            })
        else:
            text = str(item).strip()
            if text:
                normalised.append({
                    "text": text, "x0": 0.0, "y0": 0.0, "column": "LEFT",
                })

    if not normalised:
        return []

    # ---- Pre-split glued lines ----
    expanded = []
    for record in normalised:
        for piece in _split_glued_line(record["text"]):
            new = dict(record)
            new["text"] = piece
            expanded.append(new)

    # ---- Column-aware hanging-indent clustering ----
    references = []
    current = []

    column_x_values = {"LEFT": [], "RIGHT": []}
    for line in expanded:
        column_x_values[line["column"]].append(round(line["x0"], 1))

    def cluster_x_positions(values, tolerance=2.5):
        if not values:
            return []
        values = sorted(values)
        clusters = []
        for value in values:
            matched = False
            for cluster in clusters:
                center = sum(cluster) / len(cluster)
                if abs(value - center) <= tolerance:
                    cluster.append(value)
                    matched = True
                    break
            if not matched:
                clusters.append([value])
        return sorted(
            [{"x": sum(c) / len(c), "count": len(c)} for c in clusters],
            key=lambda c: c["x"],
        )

    start_margin = {}
    for column in ("LEFT", "RIGHT"):
        clusters = cluster_x_positions(column_x_values[column])
        if not clusters:
            start_margin[column] = None
            continue
        meaningful = [c for c in clusters if c["count"] >= 2]
        start_margin[column] = (
            min(c["x"] for c in meaningful) if meaningful else clusters[0]["x"]
        )

    margin_tolerance = 4.0

    for line in expanded:
        text = line["text"].strip()
        if not text:
            continue

        column = line["column"]
        base_x = start_margin.get(column)

        # Single-column degenerate case: rely on the reference-start rule.
        if base_x is None:
            if not current:
                current = [text]
            elif _looks_like_reference_start(text) and re.search(r"[.?!]\s*$", current[-1]):
                references.append(" ".join(current))
                current = [text]
            else:
                current.append(text)
            continue

        at_base_margin = abs(line["x0"] - base_x) <= margin_tolerance

        # URL-start rule: only a continuation if the previous reference
        # has NOT already ended with a sentence terminator.
        if _URL_OR_DOI_RE.match(text):
            if current and re.search(r"[.?!]\s*$", current[-1]):
                at_base_margin = True
            else:
                at_base_margin = False

        if re.fullmatch(r"\d+\.?", text):
            at_base_margin = False

        if at_base_margin:
            if current:
                references.append(" ".join(current))
            current = [text]
        else:
            if current:
                current.append(text)
            else:
                current = [text]

    if current:
        references.append(" ".join(current))

    cleaned = []
    for ref in references:
        ref = clean_text(ref)
        if ref:
            cleaned.append(ref)
    return cleaned


# ============================================================
# PLACEHOLDERS
# ============================================================

PLACEHOLDER_DOI = "doi: ???"


def build_local_chicago_reference_correction(reference):
    original = clean_text(reference)
    ref = original
    notes = []
    flags = {"missing_pp": False, "missing_doi": False, "missing_authors": False}

    if not re.match(r"^[A-Z\u00c0-\u00d6\u00d8-\u00dd]", ref):
        ref = "author ???. " + ref
        flags["missing_authors"] = True
        notes.append("Missing author list — placeholder inserted.")

    m = re.search(r"\bpp?\.\s*(\d+)\b(?!\s*[-\u2013\u2014]\s*\d)", ref)
    if m:
        ref = ref[:m.end()] + "-???" + ref[m.end():]
        flags["missing_pp"] = True
        notes.append(f"Single page number — expanded to range placeholder: {m.group(1)}-???")

    if not re.search(r"10\.\d{4,9}/", ref) and not re.search(r"https?://", ref):
        ref = ref.rstrip(".").rstrip() + f". {PLACEHOLDER_DOI}"
        flags["missing_doi"] = True
        notes.append(f"Missing DOI/URL — placeholder inserted: {PLACEHOLDER_DOI}")

    ref = ref.rstrip()
    if ref and not ref.endswith("."):
        ref = ref + "."

    return {
        "Corrected": ref,
        "Status": "REVISED" if ref != original else "MATCH",
        "Note": " | ".join(notes) if notes else "",
        "Placeholders": flags,
    }


# ============================================================
# DOI NORMALIZATION
# ============================================================

def normalize_doi_from_text(text):
    text = clean_text(text)
    compact = re.sub(r"\s+", "", text)
    previous = None
    while previous != compact:
        previous = compact
        compact = re.sub(r"(?i)(?:https?://)?(?:dx\.)?doi\.org/", "", compact, count=1)
        compact = re.sub(r"(?i)^doi:", "", compact, count=1)

    match = re.search(r"(10\.\d{4,9}/[-._;()/:A-Z0-9]+)", compact, flags=re.I)
    if not match:
        return ""
    return match.group(1).rstrip(".,;)").lower()


def canonicalize_doi_url_in_citation(text):
    text = clean_text(text)
    doi = normalize_doi_from_text(text)
    if not doi:
        return text
    doi_pattern = re.compile(
        r"(?ix)((?:(?:https?://)?(?:dx\.)?doi\.org/)+|doi\s*:\s*)?"
        r"10\.\d{4,9}\s*/\s*[-._;()/:A-Z0-9\s]+"
    )
    matches = list(doi_pattern.finditer(text))
    if not matches:
        return text
    target = matches[-1]
    before = text[:target.start()].rstrip(" ,")
    after = text[target.end():].strip()
    canonical = f"https://doi.org/{doi}"
    terminal = "." if after.endswith(".") else ""
    return f"{before}, {canonical}{terminal}" if before else f"{canonical}{terminal}"


def normalize_malformed_locator_text(text):
    value = clean_text(text or "")
    if not value:
        return ""
    previous = None
    while value != previous:
        previous = value
        value = re.sub(r"https?://(?:dx\.)?doi\.org/\s*(?=https?://)", "", value, count=1, flags=re.I)
    value = re.sub(r"https?://(?:dx\.)?doi\.org/{2,}(?=10\.)", "https://doi.org/", value, flags=re.I)
    value = re.sub(r"https?://(?:dx\.)?doi\.org/(?=10\.)", "https://doi.org/", value, flags=re.I)
    value = re.sub(r"(?<![\w/])/{2,}(10\.\d{4,9}/)", r"\1", value, flags=re.I)
    value = re.sub(
        r"\s*(?:\.\s*,|\.\s*\.+|,\s*\.|,\s*,+)\s*(?=https?://|www\.|10\.\d{4,9}/)",
        ". ", value, flags=re.I,
    )
    value = re.sub(
        r"(?<!doi\.org/)(?<![\w/])(10\.\d{4,9}/[^\s,;]+)",
        lambda m: "https://doi.org/" + m.group(1), value, flags=re.I,
    )
    value = re.sub(r"(https?://\S+?)[,;]+(?=\s|$)", r"\1", value, flags=re.I)
    loc = re.search(r"(https?://\S+|www\.\S+)\s*$", value, flags=re.I)
    if loc:
        prefix = value[:loc.start()]
        locator = re.sub(r"[.,;:]+$", "", loc.group(1).strip())
        value = prefix.rstrip() + " " + locator + "."
    value = re.sub(r"\.{2,}\s*$", ".", value)
    value = re.sub(r"\.,\s*$", ".", value)
    value = re.sub(r",\.\s*$", ".", value)
    return clean_text(value)


def get_duplicate_doi_reference_numbers(references):
    doi_to_numbers = {}
    for number, reference in enumerate(references, start=1):
        doi = normalize_doi_from_text(reference)
        if doi:
            doi_to_numbers.setdefault(doi, []).append(number)
    duplicates = set()
    for numbers in doi_to_numbers.values():
        if len(numbers) > 1:
            duplicates.update(numbers)
    return duplicates


# ============================================================
# OPENALEX DOI VERIFICATION
# ============================================================

def _get_openalex_api_key():
    return get_openalex_api_key()


def _normalize_doi_for_lookup(doi):
    if not doi:
        return ""
    s = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", doi.strip(), flags=re.I)
    s = re.sub(r"^doi\s*:\s*", "", s, flags=re.I)
    s = s.rstrip(".,;)")
    m = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", s)
    return m.group(0) if m else ""


def _fetch_openalex_metadata(doi):
    if not doi:
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

    title = data.get("title") or ""
    year = data.get("publication_year")
    authors = []
    for a in data.get("authorships", []) or []:
        name = (a.get("author") or {}).get("display_name")
        if name:
            authors.append(name)
    return {"title": title, "authors": authors, "year": year}


def _normalize_for_compare(s):
    s = re.sub(r"\s+", " ", s or "").strip().lower()
    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"[^a-z0-9 ]", "", s)
    return re.sub(r"\s+", " ", s).strip()


def _title_similarity(a, b):
    a_n, b_n = _normalize_for_compare(a), _normalize_for_compare(b)
    if not a_n or not b_n:
        return 0.0
    return difflib.SequenceMatcher(None, a_n, b_n).ratio()


def _normalize_title_for_compare(s):
    if not s:
        return ""
    s = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", s.strip())
    s = s.lower()
    s = re.sub(r"\s*-\s*", "-", s).replace("-", " ")
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _is_same_work_from_reference(reference, openalex_title, openalex_authors):
    if not openalex_title and not openalex_authors:
        return False
    ref_norm = _normalize_title_for_compare(reference)
    oa_norm = _normalize_title_for_compare(openalex_title or "")
    title_ok = False
    if ref_norm and oa_norm:
        title_ok = (oa_norm in ref_norm) or (ref_norm in oa_norm)
    if not title_ok:
        rt, ot = set(ref_norm.split()), set(oa_norm.split())
        if rt and ot:
            shared = len(rt & ot)
            title_ok = shared / max(1, min(len(rt), len(ot))) >= 0.80
    if not title_ok:
        return False
    if not openalex_authors:
        return True
    ref_tokens = set(ref_norm.split())
    for full_name in openalex_authors:
        surname = _surname_from_full_name(full_name)
        if surname and surname in ref_tokens:
            return True
    return False


def _extract_chicago_title(reference):
    m = re.search(r"[\"\u201c](.+?)[\"\u201d]", reference)
    if m:
        return m.group(1).strip().rstrip(",")
    parts = re.split(r"\.\s+", reference, maxsplit=3)
    if len(parts) >= 3:
        return parts[2].strip()
    return ""


def _surname_from_full_name(name):
    if not name:
        return ""
    name = name.strip()
    if "," in name:
        return name.split(",")[0].strip().lower()
    tokens = name.split()
    return tokens[-1].lower() if tokens else ""


def verify_reference_against_openalex(reference, parsed_doi, parsed_authors):
    result = {
        "checked": False, "doi": parsed_doi, "title_similarity": None,
        "author_overlap": None, "suspicious": False, "reasons": [],
        "crossref_title": None, "crossref_authors": [],
    }
    if not parsed_doi:
        return result
    meta = _fetch_openalex_metadata(parsed_doi)
    if meta is None:
        return result
    if meta.get("_not_found"):
        result["checked"] = True
        result["suspicious"] = True
        result["reasons"].append("DOI does not resolve in OpenAlex (possible fake DOI).")
        return result

    result["checked"] = True
    result["crossref_title"] = meta.get("title")
    result["crossref_authors"] = meta.get("authors", [])

    ref_title = _extract_chicago_title(reference)
    oa_title = meta.get("title") or ""
    if ref_title and oa_title:
        sim = _title_similarity(ref_title, oa_title)
        result["title_similarity"] = sim
        if sim < 0.55:
            result["suspicious"] = True
            result["reasons"].append(f"DOI resolves to a different title (similarity {sim:.0%}).")

    ref_surnames = {s.lower() for s in (parsed_authors or []) if s}
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
        if _is_same_work_from_reference(reference, meta.get("title") or "", meta.get("authors") or []):
            result["suspicious"] = False
            result["reasons"] = []
            result["rescued"] = True
            result["rescue_note"] = "DOI is genuine — title and author confirmed in the manuscript reference (fuzzy match)."

    return result


def extract_authors_from_chicago_reference(reference):
    text = clean_text(reference)
    text = re.sub(r"^\s*\d+[\.\)]?\s*", "", text)
    m = re.match(r"^([A-Z\u00c0-\u00d6\u00d8-\u00dd][^.,]+)", text)
    if not m:
        return []
    block = m.group(1)
    parts = re.split(r"\s+(?:and|dan|&)\s+", block, flags=re.I)
    surnames = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if "," in p:
            surnames.append(p.split(",")[0].strip())
        else:
            tokens = p.split()
            if tokens:
                surnames.append(tokens[-1])
    return surnames


def bibliography_has_truncated_author_list(text):
    text = clean_text(text)
    patterns = [
        r"\bet\s+al\.?\b", r"\band\s+others\b", r"\bwith\s+others\b",
        r"\bdkk\.?\b", r"\bdan\s+lain(?:nya)?\b", r"\bcs\.?\b",
    ]
    return any(re.search(p, text, flags=re.I) for p in patterns)


# ============================================================
# FOOTNOTE <-> BIBLIOGRAPHY MATCHING
# ============================================================

STOPWORDS = {
    "the", "and", "of", "in", "on", "for", "to", "a", "an", "with",
    "dan", "yang", "di", "dalam", "pada", "untuk", "dari", "oleh",
    "no", "vol", "volume", "issue", "https", "http", "www", "doi",
}


def footnote_bibliography_score(footnote_text, reference_text):
    def identity_tokens(text):
        text = clean_text(text).lower()
        text = re.sub(r"https?://\S+", " ", text)
        text = re.sub(r"\bdoi\s*:\s*\S+", " ", text)
        text = re.sub(r"[^\w\s'-]", " ", text, flags=re.UNICODE)
        tokens = re.findall(r"\b[\w'-]{2,}\b", text, flags=re.UNICODE)
        stop = {
            "and", "the", "of", "in", "on", "for", "to", "from", "with",
            "ed", "eds", "vol", "volume", "no", "number", "pp", "page",
            "pages", "accessed", "https", "http", "www", "doi",
        }
        return [t for t in tokens if t not in stop and not re.fullmatch(r"\d{4}", t)]

    ft = identity_tokens(footnote_text)
    bt = identity_tokens(reference_text)
    if not ft or not bt:
        return 0.0

    fs, bs = set(ft), set(bt)
    shared = fs & bs
    jaccard = len(shared) / max(1, len(fs | bs))
    coverage = len(shared) / max(1, min(len(fs), len(bs)))
    early = len(set(ft[:12]) & set(bt[:12])) / max(1, min(len(set(ft[:12])), len(set(bt[:12]))))
    fy = set(re.findall(r"\b(?:19|20)\d{2}\b", footnote_text))
    by = set(re.findall(r"\b(?:19|20)\d{2}\b", reference_text))
    yb = 0.08 if (fy and by and fy & by) else 0.0
    return min(1.0, 0.40 * jaccard + 0.35 * coverage + 0.25 * early + yb)


def match_footnotes_to_bibliography(footnotes, references, threshold=0.30):
    rows = []
    for note in footnotes:
        best_index, best_score = None, 0.0
        for idx, reference in enumerate(references):
            score = footnote_bibliography_score(note["text"], reference)
            if score > best_score:
                best_score, best_index = score, idx
        matched = best_index is not None and best_score >= threshold
        rows.append({
            "Footnote": int(note["number"]),
            "Matched": matched,
            "Best bibliography no.": (best_index + 1) if matched else None,
            "Score": round(best_score, 3),
            "Footnote text": clean_text(note["text"]),
            "Matched bibliography": references[best_index] if matched else "",
        })
    return rows


def get_matched_bibliography_numbers(match_rows):
    matched = set()
    for row in (match_rows or []):
        if not row.get("Matched", False):
            continue
        value = row.get("Best bibliography no.")
        if value is None:
            continue
        try:
            matched.add(int(value))
        except (TypeError, ValueError):
            pass
    return matched


# ============================================================
# SOURCE TYPE + STATS
# ============================================================

def normalize_source_type(source_type):
    value = clean_text(str(source_type or "")).lower()
    mapping = [
        (("journal",), "Journal Article"),
        (("book chapter", "chapter in"), "Book Chapter"),
        (("book", "monograph"), "Book"),
        (("conference", "proceeding"), "Conference"),
        (("thesis",), "Thesis"),
        (("dissertation",), "Dissertation"),
        (("newspaper", "news article", "online news"), "News / Newspaper"),
        (("magazine",), "Magazine"),
        (("government", "regulation", "legislation", "policy", "court", "legal"),
         "Government / Legal"),
        (("report", "working paper", "research report"), "Report"),
        (("website", "web page", "webpage"), "Website"),
        (("dataset",), "Dataset"),
    ]
    for keys, label in mapping:
        if any(k in value for k in keys):
            return label
    return "Other"


def extract_reference_year(reference):
    matches = list(re.finditer(r"\b(?:18|19|20)\d{2}[a-z]?\b", reference, flags=re.I))
    if not matches:
        return None
    for match in matches:
        m = re.match(r"(\d{4})", match.group(0))
        if m:
            return int(m.group(1))
    return None


def heuristic_source_type(reference):
    t = clean_text(reference).lower()
    if any(re.search(p, t, re.I) for p in [
        r"\bjournal\b", r"\bjurnal\b", r"\breview\b", r"\bquarterly\b",
        r"\bvol\.?\s*\d+", r"\bvolume\s*\d+",
        r"\bno\.?\s*\d+\s*\(", r"\d+\s*,\s*no\.?\s*\d+",
    ]):
        return "Journal Article"
    if re.search(r"\bdoi\b|doi\.org/|^10\.\d{4,9}/", t, re.I):
        return "Journal Article"
    if re.search(r"\b(thesis|skripsi)\b", t, re.I):
        return "Thesis"
    if re.search(r"\bdissertation\b", t, re.I):
        return "Dissertation"
    if re.search(r"\b(conference|proceedings|symposium)\b", t, re.I):
        return "Conference"
    if re.search(
        r"\b(undang-undang|peraturan|regulation|government regulation|"
        r"putusan|mahkamah|court|kementerian|ministry|presidential|"
        r"keputusan|surat edaran)\b", t, re.I
    ):
        return "Government / Legal"
    if re.search(
        r"\b(report|laporan|policy brief|working paper|white paper|"
        r"annual report|research report)\b", t, re.I
    ):
        return "Report"
    if re.search(
        r"\b(news|newspaper|times|post|tribun|kompas|tempo|detik|"
        r"cnn|bbc|reuters)\b", t, re.I
    ):
        return "News / Newspaper"
    if re.search(
        r"(?:jakarta|yogyakarta|bandung|surabaya|malang|depok|"
        r"london|new york|oxford|cambridge)\s*:\s*[^,.;]+", t, re.I
    ):
        return "Book"
    if re.search(r"https?://|www\.", t, re.I):
        return "Website"
    return "Other"


def build_bibliography_statistics(references, gpt_results, manuscript_year):
    result_by_no = {int(item["number"]): item for item in (gpt_results or [])}
    detail_rows = []

    for i, ref in enumerate(references, start=1):
        ai = result_by_no.get(i, {})
        original_reference = clean_text(ref)
        corrected_reference = clean_text(ai.get("revised_bibliography_markdown", "")) or original_reference
        local = build_local_chicago_reference_correction(corrected_reference)
        parsed_doi = normalize_doi_from_text(original_reference)
        parsed_authors = extract_authors_from_chicago_reference(original_reference)
        verification = verify_reference_against_openalex(original_reference, parsed_doi, parsed_authors)

        source_type = normalize_source_type(
            ai.get("source_type") or heuristic_source_type(original_reference)
        )

        year = ai.get("year")
        if year is None:
            year = extract_reference_year(original_reference)
        try:
            year = int(year) if year is not None else None
        except Exception:
            year = None

        cutoff = manuscript_year - 9
        within_10 = bool(year is not None and cutoff <= year <= manuscript_year)

        detail_rows.append({
            "No.": i, "Source Type": source_type, "Year": year,
            "Within Last 10 Years": within_10 if year is not None else None,
            "Chicago Status": ai.get("status", "NOT CHECKED"),
            "Reference": original_reference,
            "GPT Revised": local["Corrected"],
            "Explanation": ai.get("explanation", ""),
            "Correction Note": local.get("Note", ""),
            "Placeholders": local.get("Placeholders", {}),
            "DOI Verified": verification["checked"],
            "DOI Suspicious": verification["suspicious"],
            "DOI Rescued": verification.get("rescued", False),
            "DOI Verification Reasons": " | ".join(verification["reasons"]),
            "Title Similarity": verification.get("title_similarity"),
            "Author Overlap": verification.get("author_overlap"),
            "OpenAlex Title": verification.get("crossref_title"),
            "OpenAlex Authors": ", ".join(verification.get("crossref_authors", [])[:5]),
        })

    counts = Counter(row["Source Type"] for row in detail_rows)
    total = len(detail_rows)
    composition_rows = [
        {"Source Type": st, "Count": c,
         "Percentage": round(c / total * 100, 1) if total else 0.0}
        for st, c in sorted(counts.items(), key=lambda x: (-x[1], x[0]))
    ]

    detected_year_rows = [r for r in detail_rows if r["Year"] is not None]
    recent_count = sum(bool(r["Within Last 10 Years"]) for r in detected_year_rows)
    unknown_year_count = total - len(detected_year_rows)

    return detail_rows, composition_rows, {
        "total": total,
        "cutoff": manuscript_year - 9,
        "recent_count": recent_count,
        "recent_pct_all": round(recent_count / total * 100, 1) if total else 0.0,
        "recent_pct_detected": round(recent_count / len(detected_year_rows) * 100, 1)
        if detected_year_rows else 0.0,
        "unknown_year_count": unknown_year_count,
    }


# ============================================================
# DOCX HELPERS + REPORT
# ============================================================

RED = RGBColor(0xC0, 0x00, 0x00)


def _set_run_font(run, size_pt=11, bold=False, italic=False, color=None):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def _add_run(p, text, size_pt=11, bold=False, italic=False, red=False):
    r = p.add_run(text)
    _set_run_font(r, size_pt=size_pt, bold=bold, italic=italic, color=RED if red else None)
    return r


_PLACEHOLDER_RE = re.compile(
    r"(author \?{3}|author \d+(?:, author \d+)+|vol\. \?{3}|no\. \?{3}"
    r"|pp\. \d+-\?{3}|pp\. \?{3}-\?{3}|doi: \?{3})",
    re.I,
)


def _add_markdown_runs(paragraph, text, make_red=False):
    parts = re.split(r"(\*[^*]+\*)", text or "")
    for part in parts:
        if not part:
            continue
        if part.startswith("*") and part.endswith("*") and len(part) >= 2:
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        else:
            run = paragraph.add_run(part)
        if make_red:
            run.font.color.rgb = RGBColor(255, 0, 0)


def add_markdown_to_paragraph(paragraph, text, make_red=False):
    pos = 0
    for m in _PLACEHOLDER_RE.finditer(text or ""):
        before = (text or "")[pos:m.start()]
        if before:
            _add_markdown_runs(paragraph, before, make_red=make_red)
        run = paragraph.add_run(m.group(0))
        run.italic = True
        run.bold = True
        run.font.color.rgb = RGBColor(255, 0, 0)
        pos = m.end()
    tail = (text or "")[pos:]
    if tail:
        _add_markdown_runs(paragraph, tail, make_red=make_red)


def _add_divider(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "BFBFBF")
    pBdr.append(bottom)
    pPr.append(pBdr)


def create_complete_chicago_report(
    checked_notes, bibliography_detail, match_rows=None,
    recency_stats=None, pdf_filename=None,
):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)

    match_rows = match_rows or []
    missing_footnote_numbers = {
        int(row["Footnote"]) for row in match_rows if not row.get("Matched", False)
    }

    if pdf_filename:
        doc.add_heading(Path(pdf_filename).stem, level=0)

    doc.add_heading("Footnotes", level=1)

    for note in checked_notes:
        number = int(note["number"])
        corrected = clean_text(note.get("ai_revised_footnote_markdown", "")) or clean_text(note.get("text", ""))
        is_missing = number in missing_footnote_numbers
        has_dan = bool(note.get("has_dan_warning", False))
        p = doc.add_paragraph()
        nr = p.add_run(f"{number}. ")
        if is_missing:
            nr.font.color.rgb = RGBColor(255, 0, 0)
        if has_dan:
            nr.bold = True
        add_markdown_to_paragraph(p, corrected, make_red=is_missing)
        if has_dan:
            for r in p.runs:
                r.bold = True

    doc.add_page_break()
    doc.add_heading("Bibliography", level=1)

    for row in bibliography_detail:
        original = clean_text(row.get("Reference", ""))
        corrected = clean_text(row.get("GPT Revised", "")) or original
        is_matched = bool(row.get("Matched in Footnotes", False))
        is_uncited = not is_matched
        truncated = bool(row.get("Truncated Authors", False))
        doi_susp = bool(row.get("DOI Suspicious", False))

        p_orig = doc.add_paragraph()
        add_markdown_to_paragraph(p_orig, original, make_red=(doi_susp or is_uncited or truncated))

        if is_uncited:
            r = p_orig.add_run("   \u2190 NOT CITED IN FOOTNOTES")
            r.bold = True
            r.italic = True
            r.font.color.rgb = RGBColor(255, 0, 0)

        if doi_susp:
            p_w = doc.add_paragraph()
            add_markdown_to_paragraph(
                p_w,
                "Corrected version withheld — the DOI does not match the "
                "claimed title/authors. Manual verification required.",
                make_red=True,
            )
            for r in p_w.runs:
                r.bold = True
            reasons = row.get("DOI Verification Reasons", "")
            if reasons:
                p_r = doc.add_paragraph()
                rr = p_r.add_run(f"\u26a0 Possible fabricated reference: {reasons}")
                rr.italic = True
                rr.bold = True
                rr.font.color.rgb = RGBColor(255, 0, 0)
            if row.get("OpenAlex Title"):
                p_oa = doc.add_paragraph()
                r2 = p_oa.add_run(f'  OpenAlex says: "{row["OpenAlex Title"]}"')
                r2.italic = True
            if row.get("OpenAlex Authors"):
                p_oa2 = doc.add_paragraph()
                r3 = p_oa2.add_run(f"  Authors: {row['OpenAlex Authors']}")
                r3.italic = True
        else:
            p_corr = doc.add_paragraph()
            add_markdown_to_paragraph(p_corr, "Corrected: " + corrected)
            note = row.get("Correction Note", "")
            if note:
                p_n = doc.add_paragraph()
                rn = p_n.add_run(f"Fix applied: {note}")
                rn.italic = True
            active_ph = [k for k, v in (row.get("Placeholders") or {}).items() if v]
            if active_ph:
                p_ph = doc.add_paragraph()
                rp = p_ph.add_run(
                    "Placeholder used — fill in before submission: "
                    + ", ".join(active_ph)
                )
                rp.italic = True
                rp.bold = True
                rp.font.color.rgb = RGBColor(255, 0, 0)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# COMBINED GPT REQUEST
# ============================================================

class CombinedChicagoFootnoteResult(BaseModel):
    number: int
    status: str
    revised_footnote_markdown: str


class CombinedChicagoBibliographyResult(BaseModel):
    number: int
    source_type: str
    year: Optional[int] = None
    status: str
    revised_bibliography_markdown: str


class CombinedChicagoResult(BaseModel):
    footnotes: list[CombinedChicagoFootnoteResult]
    bibliography: list[CombinedChicagoBibliographyResult]


def get_openai_client():
    api_key = None
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    if not api_key:
        api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return None
    return OpenAI(api_key=api_key)


def check_all_chicago_with_gpt(footnotes, references, client):
    footnote_payload = [
        {"number": int(n["number"]), "text": clean_text(n["text"]),
         "layout_warning": bool(n.get("layout_warning", False))}
        for n in footnotes
    ]
    bibliography_payload = [
        {"number": i, "text": clean_text(r)}
        for i, r in enumerate(references, start=1)
    ]

    prompt = r"""
You are a strict academic copyeditor specializing in CHICAGO NOTES AND
BIBLIOGRAPHY style.

You will receive TWO already-extracted sections from ONE manuscript:
A. FOOTNOTES
B. BIBLIOGRAPHY

Do not extract anything from a PDF. Do not search the web.
Do not invent missing bibliographic facts.

A. FOOTNOTES
For every supplied footnote:
1. Audit as a Chicago FULL NOTE.
2. Correct safely correctable formatting.
3. Preserve all supplied facts.
4. Never invent missing facts.
5. Use *single asterisks* for required italics.
6. Correct "dan" between personal author names to "and".

Footnote STATUS: OK | REVISED | MANUAL_CHECK

B. BIBLIOGRAPHY
For every supplied bibliography entry:
1. Identify source_type.
2. Extract publication year if explicitly present, else null.
3. Audit as a Chicago BIBLIOGRAPHY entry.
4. Preserve all supplied facts. Never invent missing facts.
5. Use *single asterisks* for required italics.
6. Invert only the first personal author's name (Surname, Given Name).
7. Cross-check the first-author name order against the matching full footnote
   when the same work appears in footnotes.
8. Treat identical DOI as the same source. Do not silently delete entries.
9. Repair malformed DOI to https://doi.org/10.xxxx/xxxxx.
10. Flag author-list truncation (et al., dkk., and others) but never invent names.
11. Use double quotation marks for contained titles.
12. Ensure correct terminal punctuation (never ".," or ".." after a DOI/URL).

Bibliography STATUS: OK | REVISED | MANUAL_CHECK

Return JSON only with keys "footnotes" and "bibliography".
One result per supplied item. Preserve supplied numbers.

INPUT:
""" + json.dumps(
        {"footnotes": footnote_payload, "bibliography": bibliography_payload},
        ensure_ascii=False, indent=2,
    )

    try:
        response = client.responses.parse(
            model="gpt-4o-mini", input=prompt, text_format=CombinedChicagoResult,
        )
        parsed = response.output_parsed
        if parsed is None:
            raise ValueError("The system returned no structured combined Chicago result.")
        return {
            "footnotes": [i.model_dump() for i in parsed.footnotes],
            "bibliography": [i.model_dump() for i in parsed.bibliography],
        }
    except Exception as exc:
        error = f"{type(exc).__name__}: {exc}"
        return {
            "footnotes": [
                {"number": int(n["number"]), "status": "MANUAL_CHECK",
                 "revised_footnote_markdown": "",
                 "explanation": f"Citation service error: {error}"}
                for n in footnotes
            ],
            "bibliography": [
                {"number": i, "source_type": heuristic_source_type(r),
                 "year": extract_reference_year(r), "status": "MANUAL_CHECK",
                 "revised_bibliography_markdown": "",
                 "explanation": f"Citation service error: {error}"}
                for i, r in enumerate(references, start=1)
            ],
        }


# ============================================================
# PDF EXTRACTION (MarkItDown)
# ============================================================

def extract_pdf_text(uploaded_file):
    md = MarkItDown(enable_plugins=False)
    pdf_bytes = uploaded_file.getvalue()
    stream = io.BytesIO(pdf_bytes)
    result = md.convert_stream(stream, file_extension=".pdf")
    full_text = result.text_content or ""
    cleaned = strip_running_headers_footers(full_text)

    reference_text, ref_found, post_found = slice_reference_section(cleaned)
    body_text, _ = slice_body_section(cleaned)

    return full_text, cleaned, body_text, reference_text, ref_found, post_found


# ============================================================
# API KEY RESOLUTION
# ============================================================

def _get_openalex_api_key():
    key = st.session_state.get("openalex_api_key", "").strip()
    if key:
        return key
    try:
        if "OPENALEX_API_KEY" in st.secrets:
            return st.secrets["OPENALEX_API_KEY"]
    except Exception:
        pass
    return os.getenv("OPENALEX_API_KEY", "")


# ============================================================
# PIPELINE — one PDF
# ============================================================

def process_single_chicago_pdf(uf, batch, client, manuscript_year):
    fn_result = extract_chicago_footnotes(batch["cleaned_text"])
    footnotes = attach_page_numbers_to_footnotes(fn_result["footnotes"])
    for n in footnotes:
        n["has_dan_warning"] = footnote_has_indonesian_author_conjunction(n["text"])

    ref_lines = batch["reference_text"].splitlines()
    references = split_references_from_lines(ref_lines)

    match_rows = match_footnotes_to_bibliography(footnotes, references)

    combined = check_all_chicago_with_gpt(footnotes, references, client)
    ai_notes = combined.get("footnotes", [])
    ai_by_no = {int(x["number"]): x for x in ai_notes if x.get("number") is not None}

    checked_notes = []
    for note in footnotes:
        ai = ai_by_no.get(int(note["number"]), {})
        status = str(ai.get("status", "MANUAL_CHECK")).upper().strip()
        if status not in {"OK", "REVISED", "MANUAL_CHECK"}:
            status = "MANUAL_CHECK"
        ck = dict(note)
        ck["ai_status"] = status
        ck["ai_revised_footnote_markdown"] = ai.get("revised_footnote_markdown", "")
        ck["ai_explanation"] = ai.get("explanation", "")
        ck["decision"] = ("\u2713 OK" if status == "OK"
                          else "\u26a0 REVISED" if status == "REVISED"
                          else "\u26a0 MANUAL CHECK")
        checked_notes.append(ck)

    bibliography_results = combined.get("bibliography", [])

    detail_rows, composition_rows, recency_stats = build_bibliography_statistics(
        references, bibliography_results, int(manuscript_year)
    )

    matched_bibliography_numbers = get_matched_bibliography_numbers(match_rows)
    duplicate_doi_numbers = get_duplicate_doi_reference_numbers(references)

    corrected_bibliography_rows = []
    for row in detail_rows:
        ref_no = int(row["No."])
        original_reference = row.get("Reference", "")
        cited_in_footnotes = ref_no in matched_bibliography_numbers
        duplicate_doi = ref_no in duplicate_doi_numbers

        row["Matched in Footnotes"] = cited_in_footnotes
        row["Duplicate DOI"] = duplicate_doi
        row["GPT Revised"] = normalize_malformed_locator_text(
            canonicalize_doi_url_in_citation(row.get("GPT Revised", "") or original_reference)
        )
        final_corr = clean_text(row["GPT Revised"]) or original_reference

        row["Source Type"] = normalize_source_type(heuristic_source_type(final_corr))
        final_year = extract_reference_year(final_corr)
        if final_year is not None:
            row["Year"] = final_year

        truncated = (bibliography_has_truncated_author_list(original_reference)
                     or bibliography_has_truncated_author_list(row["GPT Revised"]))
        row["Truncated Authors"] = truncated

        cs = str(row.get("Chicago Status", "NOT CHECKED")).upper().strip()
        display_status = ("MATCH" if cs == "OK"
                          else "REVISED" if cs == "REVISED"
                          else "MANUAL CHECK" if cs == "MANUAL_CHECK"
                          else cs)
        withheld = bool(row.get("DOI Suspicious")) and not bool(row.get("DOI Rescued"))

        corrected_bibliography_rows.append({
            "No.": ref_no,
            "Source Type": row.get("Source Type", "Other"),
            "Publication Year": row["Year"] if row.get("Year") is not None else "\u2014",
            "Original Version": original_reference,
            "Corrected Version": ("\u2014 WITHHELD (DOI mismatch) \u2014"
                                  if withheld else row["GPT Revised"]),
            "Placeholders": ", ".join(k for k, v in (row.get("Placeholders") or {}).items() if v) or "\u2014",
            "Status": display_status,
            "DOI Checked": "YES" if row.get("DOI Verified") else "NO",
            "DOI Suspicious": "\u26a0\ufe0f YES" if withheld else "\u2014",
            "OpenAlex Title": (row.get("OpenAlex Title") or "")[:60],
            "DOI Issues": row.get("DOI Verification Reasons", ""),
            "Footnote in Bibliography": "\u2611 Checked" if cited_in_footnotes else "\u2610 Unchecked",
            "Bibliography Missing from Footnotes": "No" if cited_in_footnotes else "Yes",
            "Duplicate DOI": "\u26a0 Yes" if duplicate_doi else "No",
            "Incomplete Author List": "\u26a0 Yes" if truncated else "No",
        })

    total_references = len(detail_rows)
    final_counts = Counter(r.get("Source Type", "Other") for r in detail_rows)
    composition_rows = [
        {"Source Type": st, "Count": c,
         "Percentage": round(c / total_references * 100, 1) if total_references else 0.0}
        for st, c in sorted(final_counts.items(), key=lambda x: (-x[1], x[0]))
    ]

    cutoff = int(manuscript_year) - 9
    detected_year_rows = [r for r in detail_rows if r.get("Year") is not None]
    for r in detail_rows:
        y = r.get("Year")
        r["Within Last 10 Years"] = (
            bool(cutoff <= int(y) <= int(manuscript_year)) if y is not None else None
        )
    recent_count = sum(bool(r.get("Within Last 10 Years")) for r in detected_year_rows)
    recency_stats = {
        "total": total_references,
        "cutoff": cutoff,
        "recent_count": recent_count,
        "recent_pct_all": round(recent_count / total_references * 100, 1) if total_references else 0.0,
        "recent_pct_detected": round(recent_count / len(detected_year_rows) * 100, 1) if detected_year_rows else 0.0,
        "unknown_year_count": total_references - len(detected_year_rows),
    }

    report_docx = create_complete_chicago_report(
        checked_notes=checked_notes,
        bibliography_detail=detail_rows,
        match_rows=match_rows,
        recency_stats=recency_stats,
        pdf_filename=batch["filename"],
    )

    batch.update({
        "footnotes": footnotes,
        "references": references,
        "match_rows": match_rows,
        "missing_rows": [r for r in match_rows if not r.get("Matched", False)],
        "checked_notes": checked_notes,
        "detail_rows": detail_rows,
        "composition_rows": composition_rows,
        "recency_stats": recency_stats,
        "corrected_rows": corrected_bibliography_rows,
        "report_docx": report_docx,
        "manuscript_year": int(manuscript_year),
        "ai_done": True,
    })


# ============================================================
# RENDER — APA-style flow
# ============================================================

def render():
    st.title("OmniCite Auditor — Chicago Style")

    client = get_openai_client()
    if client is None:
        st.error(
            "OPENAI_API_KEY was not found. "
            "Add it to .streamlit/secrets.toml or your environment variables."
        )
        st.stop()

    openalex_key = _get_openalex_api_key()
    if not openalex_key:
        st.warning(
            "OA key was not found — DOI verification will be skipped."
        )

    if "chicago_uploader_version" not in st.session_state:
        st.session_state["chicago_uploader_version"] = 0

    uploaded_files = st.file_uploader(
        "Upload manuscript PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload up to 10 PDF files at a time.",
        key=f"chicago_uploader_{st.session_state['chicago_uploader_version']}",
    )

    if uploaded_files and len(uploaded_files) > 10:
        st.error(
            f"Maximum 10 PDF files can be uploaded at a time. "
            f"You selected {len(uploaded_files)} files. Please remove "
            f"{len(uploaded_files) - 10} file(s)."
        )
        return

    if "chicago_batches" not in st.session_state:
        st.session_state["chicago_batches"] = {}

    if not uploaded_files:
        return

    current_year = datetime.now().year
    default_year = st.session_state.get("chicago_manuscript_year", current_year)
    manuscript_year = st.number_input(
        "Manuscript publication year",
        min_value=1900,
        max_value=current_year + 5,
        value=int(default_year),
        step=1,
        help=(
            "Used to compute the % of references published within the last "
            "10 years. The window is [year-9, year], inclusive."
        ),
        key="chicago_manuscript_year_input",
    )
    st.session_state["chicago_manuscript_year"] = int(manuscript_year)

    file_keys = []
    for idx, uf in enumerate(uploaded_files):
        key = f"{idx}::{uf.name}"
        file_keys.append((key, uf))

    active_keys = {k for k, _ in file_keys}
    for k in list(st.session_state["chicago_batches"].keys()):
        if k not in active_keys:
            del st.session_state["chicago_batches"][k]

    if st.button(
        "Extract & Review",
        type="primary",
        use_container_width=True,
        key="blue_btn_run_all_chicago_pdfs",
    ):
        overall = st.progress(0, text="Starting...")
        n = len(file_keys)
        step = 100 / max(n, 1)

        for i, (key, uf) in enumerate(file_keys, start=1):
            base_pct = int((i - 1) * step)

            overall.progress(
                base_pct + int(step * 0.10),
                text=f"[{i}/{n}] Extracting {uf.name}...",
            )
            try:
                (
                    full_text,
                    cleaned_text,
                    body_text,
                    reference_text,
                    ref_found,
                    post_found,
                ) = extract_pdf_text(uf)
            except Exception as exc:
                st.error(f"{uf.name} extraction failed: {exc}")
                continue

            if not ref_found:
                st.error(
                    f"{uf.name}: No valid 'Bibliography' or 'References' "
                    f"section could be identified. This manuscript cannot "
                    f"be audited."
                )
                continue

            batch = {
                "filename": uf.name,
                "full_text": full_text,
                "cleaned_text": cleaned_text,
                "body_text": body_text,
                "reference_text": reference_text,
                "ref_found": ref_found,
                "post_found": post_found,
                "manuscript_year": int(manuscript_year),
                "ai_done": False,
            }
            st.session_state["chicago_batches"][key] = batch

            overall.progress(
                base_pct + int(step * 0.35),
                text=f"[{i}/{n}] Auditing Chicago footnotes & bibliography...",
            )
            try:
                process_single_chicago_pdf(
                    uf, batch, client, int(manuscript_year)
                )
            except Exception as exc:
                st.error(f"{uf.name} Chicago check failed: {exc}")
                continue

            overall.progress(
                base_pct + int(step * 1.00),
                text=f"[{i}/{n}] {uf.name} done.",
            )

        overall.progress(100, text="All manuscripts processed.")
        overall.empty()

    any_done = any(
        b.get("ai_done") for b in st.session_state["chicago_batches"].values()
    )
    if not any_done:
        return

    selector_options = [
        k for k, _ in file_keys
        if k in st.session_state["chicago_batches"]
    ]

    if not selector_options:
        st.info(
            "No manuscript is available for review yet. "
            "Please process a PDF successfully first."
        )
        return

    def _fmt(k):
        b = st.session_state["chicago_batches"].get(k)
        if not b:
            return str(k).split("::", 1)[-1]
        suffix = "" if b.get("ai_done") else "  (not yet processed)"
        return b.get("filename", str(k).split("::", 1)[-1]) + suffix

    selected_key = st.selectbox(
        "Select manuscript to review",
        selector_options,
        format_func=_fmt,
        key="chicago_selected_pdf_key",
    )

    batch = st.session_state["chicago_batches"].get(selected_key)
    if not batch or not batch.get("ai_done"):
        st.info(
            "This manuscript has not been processed yet. "
            "Click 'Extract & Review' above."
        )
        return

    batch["manuscript_year"] = int(
        st.session_state.get("chicago_manuscript_year", current_year)
    )

    with st.expander("View bibliography section", expanded=False):
        st.text_area(
            "Bibliography slice",
            batch["reference_text"],
            height=300,
            key=f"chicago_dbg_ref_{selected_key}",
        )

    with st.expander("View extracted footnotes", expanded=False):
        st.text_area(
            "Footnotes",
            "\n\n".join(
                f"{n['number']}. {n['text']}"
                for n in batch.get("footnotes", [])
            ) or "(none detected)",
            height=300,
            key=f"chicago_dbg_fn_{selected_key}",
        )

    checked_notes = batch.get("checked_notes", [])
    detail_rows = batch.get("detail_rows", [])
    recency_stats = batch.get("recency_stats") or {}
    match_rows = batch.get("match_rows", [])
    composition_rows = batch.get("composition_rows", [])
    corrected_rows = batch.get("corrected_rows", [])

    total_refs = recency_stats.get("total", len(detail_rows))
    citation_count = len(batch.get("footnotes", []))

    doi_checked = sum(1 for r in detail_rows if r.get("DOI Verified"))
    doi_suspicious = sum(1 for r in detail_rows if r.get("DOI Suspicious"))
    doi_suspicious_pct = (
        doi_suspicious / total_refs * 100 if total_refs else 0.0
    )

    overview_df = pd.DataFrame([
        {"Metric": "Total References", "Value": str(total_refs)},
        {
            "Metric": "Citations > 15",
            "Value": (
                f"Yes ({citation_count})"
                if citation_count > 15
                else f"No ({citation_count})"
            ),
        },
        {
            "Metric": "% Last 10 Years",
            "Value": f"{recency_stats.get('recent_pct_all', 0.0)}%",
        },
        {
            "Metric": "Footnotes Missing from Bibliography",
            "Value": str(len(batch.get("missing_rows", []))),
        },
        {
            "Metric": "Bibliography Missing from Footnotes",
            "Value": str(
                total_refs
                - len(get_matched_bibliography_numbers(match_rows))
            ),
        },
        {
            "Metric": "DOI Checked (OpenAlex)",
            "Value": str(doi_checked),
        },
        {
            "Metric": "DOI Suspicious (possible fabrication)",
            "Value": f"{doi_suspicious} ({doi_suspicious_pct:.1f}%)",
        },
    ])

    source_lookup = {
        r["Source Type"]: f"{r['Count']} ({r['Percentage']}%)"
        for r in composition_rows
    }
    composition_df = pd.DataFrame([
        {
            "Source Type": st_,
            "Count / Percentage": source_lookup.get(st_, "0 (0.0%)"),
        }
        for st_ in [
            "Journal Article",
            "Book",
            "Government / Legal",
            "Report",
            "News / Newspaper",
            "Website",
            "Other",
        ]
    ])

    left, right = st.columns(2, gap="large")
    with left:
        st.dataframe(overview_df, use_container_width=True, hide_index=True)
    with right:
        st.dataframe(composition_df, use_container_width=True, hide_index=True)

    show_fn = st.toggle(
        "Footnote Comparison",
        value=False,
        key=f"chicago_show_fn_{selected_key}",
    )
    if show_fn:
        fn_rows = []
        for note in checked_notes:
            raw = str(note.get("ai_status", "MANUAL_CHECK")).upper().strip()
            disp = (
                "MATCH" if raw == "OK"
                else "REVISED" if raw == "REVISED"
                else "MANUAL CHECK"
            )
            fn_rows.append({
                "No.": note.get("number", ""),
                "Original Footnote": clean_text(note.get("text", "")),
                "Corrected AI Version": (
                    clean_text(note.get("ai_revised_footnote_markdown", ""))
                    or clean_text(note.get("text", ""))
                ),
                "Status": disp,
            })
        if fn_rows:
            st.dataframe(
                pd.DataFrame(fn_rows),
                use_container_width=True,
                hide_index=True,
                height=230,
            )
        else:
            st.info("No footnotes available for comparison.")

    show_bib = st.toggle(
        "Bibliography Comparison",
        value=False,
        key=f"chicago_show_bib_{selected_key}",
    )
    if show_bib:
        if corrected_rows:
            df = pd.DataFrame(corrected_rows)
            preferred = [
                "No.", "Source Type", "Publication Year",
                "Original Version", "Corrected Version", "Status",
                "Footnote in Bibliography",
                "Bibliography Missing from Footnotes",
                "Duplicate DOI", "Incomplete Author List",
                "Placeholders", "DOI Checked", "DOI Suspicious",
                "OpenAlex Title", "DOI Issues",
            ]
            existing = [c for c in preferred if c in df.columns]
            rest = [c for c in df.columns if c not in existing]
            df = df[existing + rest]
            st.caption(
                f"Showing {len(df)} of "
                f"{len(batch.get('references', []))} extracted entries."
            )
            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True,
                height=280,
            )
        else:
            st.info("No bibliography entries available.")

    report_docx = batch.get("report_docx")
    if report_docx:
        safe_name = re.sub(
            r"[^\w\-]+", "_", batch.get("filename", "manuscript")
        )
        st.download_button(
            label="📄 Download Chicago Diagnostic Report (.docx)",
            data=report_docx,
            file_name=f"{safe_name}_chicago_diagnostic_report.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
            key=f"chicago_download_{selected_key}",
        )

    st.markdown(
        """
        <style>
        div[class*="st-key-reset_btn"] button {
            background-color: #dc2626 !important;
            color: #ffffff !important;
            border: 1px solid #b91c1c !important;
            font-weight: 600 !important;
            transition: background-color 0.15s ease;
        }
        div[class*="st-key-reset_btn"] button:hover {
            background-color: #b91c1c !important;
            color: #ffffff !important;
            border-color: #991b1b !important;
        }
        div[class*="st-key-reset_btn"] button:focus {
            box-shadow: 0 0 0 0.2rem rgba(220, 38, 38, 0.4) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    if st.button(
        "🔄 Start Fresh — Clear All Uploads & Results",
        use_container_width=True,
        key="reset_btn_start_fresh",
    ):
        for k in list(st.session_state.keys()):
            if k == "chicago_batches" or k.startswith("chicago_batches"):
                del st.session_state[k]
            if k.startswith("chicago_dbg_"):
                del st.session_state[k]
            if k == "chicago_selected_pdf_key":
                del st.session_state[k]

        st.session_state["chicago_batches"] = {}
        st.session_state["chicago_uploader_version"] = (
            st.session_state.get("chicago_uploader_version", 0) + 1
        )
        st.session_state["chicago_manuscript_year"] = datetime.now().year

        st.rerun()