# chicago.py
# OmniCite Auditor — Chicago Notes & Bibliography Module
# Imported and rendered by app.py via:  import chicago; chicago.render()

from pathlib import Path

import re
import os
import json
import fitz          # PyMuPDF — REQUIRED for two-column + footnote geometry
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
from openai import OpenAI
from pydantic import BaseModel
from collections import Counter
from datetime import datetime
from typing import Optional

from openalex_config import get_openalex_api_key


# ============================================================
# OPENAI CLIENT
# ============================================================

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


# ============================================================
# OPENALEX DOI VERIFICATION
# ============================================================

def _get_openalex_api_key():
    return get_openalex_api_key()


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

    title = data.get("title") or ""
    year = data.get("publication_year")
    authors = []
    for a in data.get("authorships", []) or []:
        name = (a.get("author") or {}).get("display_name")
        if name:
            authors.append(name)
    return {"title": title, "authors": authors, "year": year}


def _normalize_for_compare(s):
    if not s:
        return ""
    s = re.sub(r"\s+", " ", s).strip().lower()
    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"[^a-z0-9 ]", "", s)
    return re.sub(r"\s+", " ", s).strip()


def _title_similarity(a, b):
    import difflib
    a_n = _normalize_for_compare(a)
    b_n = _normalize_for_compare(b)
    if not a_n or not b_n:
        return 0.0
    return difflib.SequenceMatcher(None, a_n, b_n).ratio()


def _normalize_title_for_compare(s):
    if not s:
        return ""
    s = s.strip()
    s = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", s)
    s = s.lower()
    s = re.sub(r"\s*-\s*", "-", s)
    s = s.replace("-", " ")
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
        ref_tokens = set(ref_norm.split())
        oa_tokens = set(oa_norm.split())
        if ref_tokens and oa_tokens:
            shared = len(ref_tokens & oa_tokens)
            title_ok = shared / max(1, min(len(ref_tokens), len(oa_tokens))) >= 0.80
    if not title_ok:
        return False
    if not openalex_authors:
        return True
    ref_tokens = set(ref_norm.split())
    for full_name in openalex_authors:
        surname = _surname_from_full_name(full_name)
        if not surname:
            continue
        if surname in ref_tokens:
            return True
    return False


def _extract_chicago_title(reference):
    """Extract the title. Chicago journals use SINGLE quotes for articles."""
    m = re.search(r"[\"\u201c](.+?)[\"\u201d]", reference)
    if m:
        return m.group(1).strip().rstrip(",")
    m = re.search(r"['\u2018](.+?)['\u2019]", reference)
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
            result["reasons"].append(
                f"DOI resolves to a different title (similarity {sim:.0%})."
            )

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
        if _is_same_work_from_reference(
            reference, meta.get("title") or "", meta.get("authors") or []
        ):
            result["suspicious"] = False
            result["reasons"] = []
            result["rescued"] = True
            result["rescue_note"] = (
                "DOI is genuine — title and author confirmed in the "
                "manuscript reference (fuzzy match)."
            )
    return result


def extract_authors_from_chicago_reference(reference):
    text = clean_text(reference)
    text = re.sub(r"^\s*\d+[\.\)]?\s*", "", text)
    m = re.match(r"^([A-ZÀ-ÖØ-Ý][^.,]+)", text)
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
# DOI TEXT NORMALIZATION
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
        r"""(?ix)((?:(?:https?://)?(?:dx\.)?doi\.org/)+|doi\s*:\s*)?
        10\.\d{4,9}\s*/\s*[-._;()/:A-Z0-9\s]+"""
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
    locator_at_end = re.search(r"(https?://\S+|www\.\S+)\s*$", value, flags=re.I)
    if locator_at_end:
        prefix = value[:locator_at_end.start()]
        locator = re.sub(r"[.,;:]+$", "", locator_at_end.group(1).strip())
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
# PLACEHOLDERS
# ============================================================

PLACEHOLDER_DOI = "doi: ???"


def build_local_chicago_reference_correction(reference):
    original = clean_text(reference)
    ref = original
    notes = []
    placeholder_flags = {
        "missing_pp": False, "missing_doi": False, "missing_authors": False,
    }
    if not re.match(r"^[A-ZÀ-ÖØ-Ý]", ref):
        ref = "author ???. " + ref
        placeholder_flags["missing_authors"] = True
        notes.append("Missing author list — placeholder inserted.")
    m = re.search(r"\bpp?\.\s*(\d+)\b(?!\s*[-–—]\s*\d)", ref)
    if m:
        ref = ref[:m.end()] + "-???" + ref[m.end():]
        placeholder_flags["missing_pp"] = True
        notes.append(f"Single page number — expanded to range placeholder: {m.group(1)}-???")
    if not re.search(r"10\.\d{4,9}/", ref) and not re.search(r"https?://", ref):
        ref = ref.rstrip(".").rstrip() + f". {PLACEHOLDER_DOI}"
        placeholder_flags["missing_doi"] = True
        notes.append(f"Missing DOI/URL — placeholder inserted: {PLACEHOLDER_DOI}")
    ref = ref.rstrip()
    if ref and not ref.endswith("."):
        ref = ref + "."
    changed = ref != original
    return {
        "Corrected": ref,
        "Status": "REVISED" if changed else "MATCH",
        "Note": " | ".join(notes) if notes else "",
        "Placeholders": placeholder_flags,
    }


# ============================================================
# BASIC TEXT CLEANING
# ============================================================

def clean_text(text):
    if not text:
        return ""
    text = text.replace("\u00ad", "")
    text = text.replace("‐", "-")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def clean_pdf_text(text):
    if not text:
        return ""
    text = text.replace("\u00ad", "")
    text = text.replace("‐", "-")
    lines = []
    for line in text.splitlines():
        line = re.sub(r"[ \t]+", " ", line).strip()
        if line:
            lines.append(line)
    return "\n".join(lines)


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
    debug_rows = []
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
        for line in candidates:
            detected = detect_footnote_number(line)
            debug_rows.append({
                "Page": page_number,
                "y0": round(line["y0"], 2),
                "x0": round(line["x0"], 2),
                "Font": round(line["avg_size"], 2),
                "Body Font": body_size,
                "Starts Note": str(detected[0]) if detected else "",
                "Text": line["text"],
            })
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

    return {
        "footnotes": all_notes,
        "count": len(all_notes),
        "bibliography_page": (
            bibliography_page + 1 if bibliography_page < 999999 else None
        ),
        "debug": debug_rows,
        "page_candidates": page_candidates,
    }


# ============================================================
# BIBLIOGRAPHY EXTRACTION — PyMuPDF column-aware
# ============================================================

def looks_like_running_header_footer(text, y0, y1, page_height):
    clean = re.sub(r"\s+", " ", text).strip()
    if re.fullmatch(r"\d{1,4}", clean):
        if y0 < page_height * 0.08 or y1 > page_height * 0.94:
            return True
    if y0 < page_height * 0.08:
        if re.search(r"\bvol\.?\s*\d+", clean, re.I):
            return True
        if re.search(r"\bvolume\s+\d+", clean, re.I):
            return True
    return False


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


# ============================================================
# REFERENCE SPLITTING — geometry mode vs heuristic mode
# ============================================================

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


def _split_glued_line(text: str) -> list[str]:
    """Split a single physical line that glued several references together."""
    if not text:
        return []
    if not re.search(r"\.\s+[A-Z\u00c0-\u00d6\u00d8-\u00dd]", text):
        return [text.strip()]

    candidates = []
    for m in re.finditer(r"(?<=\.)\s+(?=[A-Z\u00c0-\u00d6\u00d8-\u00dd])", text):
        candidates.append(m.start())
    if not candidates:
        return [text.strip()]

    boundaries = [0]
    for pos in candidates:
        before = text[:pos]
        after = text[pos:].lstrip()
        if re.search(r"https?://[^\s]*$", before):
            continue
        if _URL_OR_DOI_RE.match(after):
            continue
        if _looks_like_reference_start(after):
            boundaries.append(pos)
    boundaries.append(len(text))

    pieces = []
    for i in range(len(boundaries) - 1):
        piece = text[boundaries[i]:boundaries[i + 1]].strip()
        if piece:
            pieces.append(piece)

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

def _joined_lookahead(lines, start_index, max_lines=3):
    """
    Concatenate up to `max_lines` starting at `start_index`.
    Handles author names that wrap across two or three visual lines,
    e.g. "Firdaus," followed by "Muhammad," on the next line.
    """
    parts = []
    for j in range(start_index, min(start_index + max_lines, len(lines))):
        parts.append(lines[j]["text"].strip())
    return " ".join(parts)


def _starts_new_reference_at(lines, index):
    """
    True if line `index` — possibly joined with the next two lines to
    handle wrapped author names — begins a new bibliography entry.
    """
    joined = _joined_lookahead(lines, index, max_lines=3)
    return _looks_like_reference_start(joined)


def _repair_bibliography_pdf_breaks(text: str) -> str:
    """Repair PDF wrapping artifacts *without swallowing the next reference*.

    Important: never concatenate an arbitrary alphabetic token after a DOI.
    A previous version could turn:
        ...10.18196/jmh.v30i2.18628. Firmantoro, Kiki...
    into:
        ...10.18196/jmh.v30i2.18628.Firmantoro...
    which hid the bibliography boundary.
    """
    value = clean_text(text or "")
    if not value:
        return ""

    # Repair only very high-confidence DOI numeric wraps, e.g.
    #   10.18196/jmh.v30i2.186 28  ->  10.18196/jmh.v30i2.18628
    # Do NOT join alphabetic words: they may be the next author's surname.
    numeric_doi_wrap = re.compile(
        r"(10\.\d{4,9}/[-._;()/:A-Za-z0-9]*\d)\s+(\d{1,6})(?=(?:[.,;)]|\s|$))",
        re.I,
    )
    previous = None
    while value != previous:
        previous = value
        value = numeric_doi_wrap.sub(r"\1\2", value)

    # Repair whitespace immediately after DOI punctuation only when it is still
    # clearly part of a numeric DOI suffix. Never consume a capitalised surname.
    value = re.sub(
        r"(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+[._/-])\s+(\d+)",
        r"\1\2", value, flags=re.I,
    )

    # Common URL wraps. These are limited to an explicit slash or hyphen so a
    # following author name cannot be absorbed into the URL.
    value = re.sub(r"(https?://\S+/)\s+(?=[A-Za-z0-9])", r"\1", value, flags=re.I)
    value = re.sub(r"(https?://\S+-)\s+(?=[A-Za-z0-9])", r"\1", value, flags=re.I)

    # Canonicalise duplicate DOI prefixes only.
    value = re.sub(
        r"https?://(?:dx\.)?doi\.org/\s*https?://(?:dx\.)?doi\.org/",
        "https://doi.org/", value, flags=re.I,
    )
    return clean_text(value)


def _probable_bibliography_start(text: str) -> bool:
    """Permissive validator used *after* geometry says a line is at the first-line margin."""
    s = clean_text(text or "")
    if not s or _URL_OR_DOI_RE.match(s):
        return False

    # Standard inverted bibliography name: Surname, Given ...
    if re.match(
        r"^[A-Z\u00c0-\u00d6\u00d8-\u00dd][A-Za-z\u00c0-\u00ff'\u2019\-]+,\s*"
        r"[A-Z\u00c0-\u00d6\u00d8-\u00dd]",
        s,
    ):
        return True

    # A surname may occupy a whole visual line: "Firdaus,".
    if re.match(r"^[A-Z\u00c0-\u00d6\u00d8-\u00dd][A-Za-z\u00c0-\u00ff'\u2019\-]+,\s*$", s):
        return True

    # Indonesian single-name author followed by title: "Syafliansah. Metode ..."
    if re.match(
        r"^[A-Z\u00c0-\u00d6\u00d8-\u00dd][A-Za-z\u00c0-\u00ff'\u2019\-]{2,}\.\s+"
        r"[A-Z\u00c0-\u00d6\u00d8-\u00dd]",
        s,
    ):
        return True

    # Preserve existing recognisers for numbered, quoted, and corporate entries.
    return _looks_like_reference_start(s)


def _column_start_margins(lines, cluster_tolerance=2.5):
    """Infer bibliography first-line x positions independently for left/right columns."""
    by_col = {"LEFT": [], "RIGHT": []}
    for line in lines:
        by_col.setdefault(line.get("column", "LEFT"), []).append(float(line.get("x0", 0.0)))

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
        return [
            {"x": sum(g) / len(g), "count": len(g)}
            for g in groups
        ]

    margins = {}
    for col, values in by_col.items():
        cs = clusters(values)
        if not cs:
            margins[col] = None
            continue
        # A first-line margin repeats. Prefer the left-most repeated cluster.
        repeated = [c for c in cs if c["count"] >= 2]
        margins[col] = min(c["x"] for c in (repeated or cs))
    return margins


def split_references_from_lines(lines):
    """
    Split bibliography into entries.

    Geometry mode deliberately treats hanging-indent reset as the primary signal.
    Textual author detection is supporting evidence, not a requirement that the
    previous physical line must end in punctuation. This avoids glued references
    when a DOI/URL/name wraps at the end of a PDF line.
    """
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
                normalised.append({"text": text, "x0": 0.0, "y0": 0.0, "page": None, "column": "LEFT"})

    if not normalised:
        return []

    # ----------------------- HEURISTIC FALLBACK -----------------------
    if not has_geometry:
        expanded = []
        for record in normalised:
            for piece in _split_glued_line(record["text"]):
                new = dict(record)
                new["text"] = piece
                expanded.append(new)

        refs, current = [], []
        for idx, line in enumerate(expanded):
            text = line["text"]
            if not current:
                current = [text]
                continue
            starts_new = (
                not _URL_OR_DOI_RE.match(text)
                and (_starts_new_reference_at(expanded, idx) or _probable_bibliography_start(text))
                and bool(re.search(r"[.?!]\s*$", current[-1]))
            )
            if starts_new:
                refs.append(_repair_bibliography_pdf_breaks(" ".join(current)))
                current = [text]
            else:
                current.append(text)
        if current:
            refs.append(_repair_bibliography_pdf_breaks(" ".join(current)))
        return [r for r in refs if r]

    # ------------------------- GEOMETRY MODE --------------------------
    margins = _column_start_margins(normalised)
    margin_tolerance = 4.5
    refs, current = [], []
    current_column = None

    for idx, line in enumerate(normalised):
        text = line["text"]
        col = line["column"]
        base = margins.get(col)
        at_start_margin = base is not None and abs(line["x0"] - base) <= margin_tolerance
        url_or_doi = bool(_URL_OR_DOI_RE.match(text))

        # Look ahead because an author's inverted name can wrap over visual lines.
        content_start = _starts_new_reference_at(normalised, idx) or _probable_bibliography_start(text)

        # Column transition is important: the first item in a new column is a new
        # reference if it is at that column's first-line margin and looks plausible.
        column_changed = current and current_column is not None and col != current_column

        starts_new = bool(
            current
            and at_start_margin
            and not url_or_doi
            and content_start
        )

        # If geometry is very clear at a column transition, do not require the
        # previous physical line to end in punctuation.
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

    # Final guard: split any residual same-line glue, but only at highly probable
    # author starts. This is intentionally after geometry segmentation.
    final_refs = []
    for ref in refs:
        repaired = _repair_bibliography_pdf_breaks(ref)
        pieces = _split_glued_line(repaired)
        for piece in pieces:
            piece = clean_text(piece)
            if not piece:
                continue
            # One more pass catches chains such as DOI. Author ... DOI. Author.
            subpieces = _split_glued_line(piece)
            final_refs.extend(clean_text(x) for x in subpieces if clean_text(x))

    return [clean_text(r) for r in final_refs if clean_text(r)]


# ============================================================
# 'dan' WARNING
# ============================================================

def footnote_has_indonesian_author_conjunction(text):
    return bool(re.search(r"\s+dan\s+", clean_text(text), flags=re.I))


# ============================================================
# FOOTNOTE ↔ BIBLIOGRAPHY MATCHING
# ============================================================

STOPWORDS = {
    "the", "and", "of", "in", "on", "for", "to", "a", "an", "with",
    "dan", "yang", "di", "dalam", "pada", "untuk", "dari", "oleh",
    "no", "vol", "volume", "issue", "https", "http", "www", "doi",
}


def normalized_tokens(text):
    text = clean_text(text).lower()
    text = re.sub(r"https?://\S+|www\.\S+", " ", text)
    text = re.sub(r"\b10\.\d{4,9}/\S+", " ", text)
    text = re.sub(r"[^0-9a-zà-öø-ÿ]+", " ", text, flags=re.I)
    return [t for t in text.split() if len(t) >= 3 and t not in STOPWORDS]


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

    foot_tokens = identity_tokens(footnote_text)
    bib_tokens = identity_tokens(reference_text)
    if not foot_tokens or not bib_tokens:
        return 0.0
    foot_set = set(foot_tokens)
    bib_set = set(bib_tokens)
    shared = foot_set & bib_set
    jaccard = len(shared) / max(1, len(foot_set | bib_set))
    coverage = len(shared) / max(1, min(len(foot_set), len(bib_set)))
    foot_early = set(foot_tokens[:12])
    bib_early = set(bib_tokens[:12])
    early_overlap = len(foot_early & bib_early) / max(1, min(len(foot_early), len(bib_early)))
    foot_years = set(re.findall(r"\b(?:19|20)\d{2}\b", footnote_text))
    bib_years = set(re.findall(r"\b(?:19|20)\d{2}\b", reference_text))
    year_bonus = 0.08 if (foot_years and bib_years and foot_years & bib_years) else 0.0
    return min(1.0, 0.40 * jaccard + 0.35 * coverage + 0.25 * early_overlap + year_bonus)


def match_footnotes_to_bibliography(footnotes, references, threshold=0.30):
    rows = []
    for note in footnotes:
        best_index = None
        best_score = 0.0
        for idx, reference in enumerate(references):
            score = footnote_bibliography_score(note["text"], reference)
            if score > best_score:
                best_score = score
                best_index = idx
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
    values = []
    for match in matches:
        m = re.match(r"(\d{4})", match.group(0))
        if m:
            values.append(int(m.group(1)))
    return values[0] if values else None


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
# DOCX HELPERS
# ============================================================

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


def _add_red_italic_run(paragraph, text, size_pt=11):
    _add_run(paragraph, text, size_pt=size_pt, bold=True, italic=True, red=True)


def _docx_set_default_font(document, font_name="Times New Roman", size_pt=11):
    style = document.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(size_pt)


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


# ============================================================
# DOCX REPORT — mirrors APA layout
# ============================================================

def create_complete_chicago_report(
    checked_notes,
    bibliography_detail,
    match_rows=None,
    recency_stats=None,
    pdf_filename=None,
    composition_rows=None,
    manuscript_year=None,
):
    """
    Build the Chicago diagnostic report with the SAME structure as
    build_apa_report_docx():

      1. Title + subtitle + meta
      2. Section 1 — Summary (metrics table + source-type table)
      3. Section 2 — Footnotes  (Original / Corrected / Note / Status)
      4. Section 3 — Bibliography (Original / Corrected / Comment / Status)
         with a horizontal rule between entries
    """
    doc = Document()
    _docx_set_default_font(doc, "Times New Roman", 11)

    match_rows = match_rows or []
    recency_stats = recency_stats or {}
    composition_rows = composition_rows or []
    manuscript_year = manuscript_year or datetime.now().year

    # ---- Header -------------------------------------------------------
    title = doc.add_heading(level=0)
    tr = title.add_run(pdf_filename or "manuscript.pdf")
    _set_run_font(tr, size_pt=18, bold=True)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    sr = subtitle.add_run(
        "Chicago Notes & Bibliography — Citation & Reference Diagnostic Report"
    )
    _set_run_font(sr, size_pt=11, italic=True)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = doc.add_paragraph()
    mr = meta.add_run(
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Engine: OmniCite-Chicago-v1"
    )
    _set_run_font(mr, size_pt=9, italic=True)
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # ---- Compute summary values ---------------------------------------
    total_refs = int(recency_stats.get("total", len(bibliography_detail)))
    total_footnotes = len(checked_notes)

    missing_footnote_numbers = {
        int(row["Footnote"])
        for row in match_rows
        if not row.get("Matched", False)
    }
    footnotes_missing_from_bib = len(missing_footnote_numbers)

    matched_bib_numbers = get_matched_bibliography_numbers(match_rows)
    bib_missing_from_footnotes = max(0, total_refs - len(matched_bib_numbers))

    doi_checked = sum(1 for r in bibliography_detail if r.get("DOI Verified"))
    doi_suspicious = sum(1 for r in bibliography_detail if r.get("DOI Suspicious"))
    doi_suspicious_pct = (doi_suspicious / total_refs * 100) if total_refs else 0.0

    recent_pct = recency_stats.get("recent_pct_all", 0.0)
    cutoff = recency_stats.get("cutoff", manuscript_year - 9)

    # ---- Section 1 — Summary ------------------------------------------
    h1 = doc.add_heading(level=1)
    hr = h1.add_run("1. Summary")
    _set_run_font(hr, size_pt=14, bold=True)

    summary_rows = [
        ("Manuscript publication year", str(manuscript_year), False),
        ("Total bibliography entries", str(total_refs), False),
        ("Total footnotes", str(total_footnotes), False),
        ("Footnotes missing from bibliography",
         str(footnotes_missing_from_bib), footnotes_missing_from_bib > 0),
        ("Bibliography entries not cited in footnotes",
         str(bib_missing_from_footnotes), bib_missing_from_footnotes > 0),
        ("DOI checked (OpenAlex)", str(doi_checked), False),
        ("DOI suspicious (possible fabricated references)",
         f"{doi_suspicious} ({doi_suspicious_pct:.1f}%)", doi_suspicious > 0),
        (f"% references within last 10 years "
         f"({cutoff}-{manuscript_year})",
         f"{recent_pct:.1f}%", False),
    ]

    summary_table = doc.add_table(rows=1, cols=2)
    summary_table.style = "Light Grid Accent 1"
    for cell, text in zip(summary_table.rows[0].cells, ["Metric", "Value"]):
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

    # ---- 1.1 Source type distribution ---------------------------------
    h11 = doc.add_heading(level=2)
    hr11 = h11.add_run("1.1 Source Type Distribution")
    _set_run_font(hr11, size_pt=12, bold=True)

    source_lookup = {r.get("Source Type", "Other"): r for r in composition_rows}
    source_order = [
        "Journal Article", "Book", "Book Chapter", "Conference",
        "Government / Legal", "Report", "News / Newspaper", "Website",
        "Thesis", "Dissertation", "Magazine", "Dataset", "Other",
    ]

    src_table = doc.add_table(rows=1, cols=3)
    src_table.style = "Light Grid Accent 1"
    for cell, text in zip(src_table.rows[0].cells,
                          ["Source Type", "Count", "Percentage"]):
        cell.text = ""
        _set_run_font(cell.paragraphs[0].add_run(text), size_pt=10, bold=True)

    for stype in source_order:
        r = source_lookup.get(stype)
        cnt = r.get("Count", 0) if r else 0
        pct = r.get("Percentage", 0.0) if r else 0.0
        cells = src_table.add_row().cells
        for cell, val in zip(cells, [stype, str(cnt), f"{pct:.1f}%"]):
            cell.text = ""
            _set_run_font(cell.paragraphs[0].add_run(val), size_pt=10)

    cells = src_table.add_row().cells
    for cell, val in zip(cells, ["Total", str(total_refs), "100.0%"]):
        cell.text = ""
        _set_run_font(cell.paragraphs[0].add_run(val), size_pt=10, bold=True)

    doc.add_page_break()

    # ---- Section 2 — Footnotes ----------------------------------------
    h2 = doc.add_heading(level=1)
    hr2 = h2.add_run("2. Footnotes")
    _set_run_font(hr2, size_pt=14, bold=True)

    caption = doc.add_paragraph()
    cr = caption.add_run(
        "Format: footnote number — original footnote — corrected footnote"
    )
    _set_run_font(cr, size_pt=9, italic=True)

    if not checked_notes:
        p = doc.add_paragraph()
        _set_run_font(p.add_run("No footnotes were detected."), italic=True)
    else:
        for note in checked_notes:
            number = int(note["number"])
            original = clean_text(note.get("text", ""))
            revised_raw = clean_text(note.get("ai_revised_footnote_markdown", ""))
            revised = revised_raw or original

            raw_status = str(note.get("ai_status", "MANUAL_CHECK")).upper().strip()
            display_status = (
                "MATCH" if raw_status == "OK"
                else "REVISED" if raw_status == "REVISED"
                else "MANUAL CHECK"
            )

            is_missing = number in missing_footnote_numbers
            has_dan = bool(note.get("has_dan_warning", False))

            head = doc.add_paragraph()
            head.paragraph_format.space_before = Pt(4)
            head.paragraph_format.space_after = Pt(2)
            hrun = head.add_run(f"{number}.")
            _set_run_font(hrun, size_pt=11, bold=True)

            pages = note.get("pages") or []
            if pages:
                pr = head.add_run(f"  (p. {', '.join(map(str, pages))})")
                _set_run_font(pr, size_pt=9, italic=True)

            if is_missing:
                _add_red_italic_run(head, "  [NOT MATCHED TO BIBLIOGRAPHY]")
            if has_dan:
                _add_red_italic_run(head, "  [\"dan\" BETWEEN AUTHORS]")

            p_orig = doc.add_paragraph()
            p_orig.paragraph_format.left_indent = Inches(0.25)
            p_orig.paragraph_format.space_after = Pt(2)
            _add_run(p_orig, "Original:  ", bold=True, size_pt=11)
            if is_missing or has_dan:
                _add_run(p_orig, original, size_pt=11, red=True)
            else:
                _add_run(p_orig, original, size_pt=11)

            p_corr = doc.add_paragraph()
            p_corr.paragraph_format.left_indent = Inches(0.25)
            p_corr.paragraph_format.space_after = Pt(2)
            _add_run(p_corr, "Corrected: ", bold=True, size_pt=11)

            if not revised_raw:
                _add_run(p_corr, "— no change —", size_pt=11, italic=True)
            else:
                add_markdown_to_paragraph(p_corr, revised)

            note_text = note.get("ai_explanation", "")
            if note_text:
                np = doc.add_paragraph()
                np.paragraph_format.left_indent = Inches(0.25)
                np.paragraph_format.space_after = Pt(2)
                _add_run(np, f"Note: {note_text}", size_pt=10, italic=True)

            p_status = doc.add_paragraph()
            p_status.paragraph_format.left_indent = Inches(0.25)
            p_status.paragraph_format.space_after = Pt(4)
            _add_run(p_status, "Status: ", bold=True, size_pt=11)
            _add_run(
                p_status, display_status, size_pt=11, bold=True,
                red=(display_status != "MATCH"),
            )

            _add_divider(doc)

    doc.add_page_break()

    # ---- Section 3 — Bibliography -------------------------------------
    h3 = doc.add_heading(level=1)
    hr3 = h3.add_run("3. Bibliography (Chicago Notes & Bibliography)")
    _set_run_font(hr3, size_pt=14, bold=True)

    caption = doc.add_paragraph()
    cr2 = caption.add_run(
        "Format: original reference — corrected reference — comment — status"
    )
    _set_run_font(cr2, size_pt=9, italic=True)

    if not bibliography_detail:
        p = doc.add_paragraph()
        _set_run_font(p.add_run("No bibliography entries were detected."), italic=True)
    else:
        for row in bibliography_detail:
            reference_no = int(row.get("No.", 0))
            original = clean_text(row.get("Reference", ""))
            corrected = clean_text(row.get("GPT Revised", "")) or original

            is_uncited = reference_no not in matched_bib_numbers
            truncated_authors = bool(row.get("Truncated Authors", False))
            doi_suspicious = bool(row.get("DOI Suspicious", False))
            duplicate_doi = bool(row.get("Duplicate DOI", False))

            chicago_status = str(row.get("Chicago Status", "NOT CHECKED")).upper().strip()
            display_status = (
                "MATCH" if chicago_status == "OK"
                else "REVISED" if chicago_status == "REVISED"
                else "MANUAL CHECK" if chicago_status == "MANUAL_CHECK"
                else chicago_status
            )
            withheld = doi_suspicious and not bool(row.get("DOI Rescued", False))

            head = doc.add_paragraph()
            head.paragraph_format.space_before = Pt(6)
            head.paragraph_format.space_after = Pt(2)
            hrun = head.add_run(f"{reference_no}.")
            _set_run_font(hrun, size_pt=11, bold=True)

            if is_uncited:
                _add_red_italic_run(head, "  [NOT CITED IN FOOTNOTES]")
            if duplicate_doi:
                _add_red_italic_run(head, "  [DUPLICATE DOI]")
            if truncated_authors:
                _add_red_italic_run(head, "  [INCOMPLETE AUTHOR LIST]")
            if withheld:
                _add_red_italic_run(head, "  [DOI MISMATCH — WITHHELD]")

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
            _add_run(
                p_orig, original, size_pt=11,
                red=(is_uncited or doi_suspicious or truncated_authors),
            )

            if withheld:
                p_corr = doc.add_paragraph()
                p_corr.paragraph_format.left_indent = Inches(0.25)
                p_corr.paragraph_format.space_after = Pt(2)
                _add_run(p_corr, "Corrected: ", bold=True, size_pt=11)
                _add_run(
                    p_corr,
                    "— withheld (DOI mismatch) —",
                    size_pt=11, italic=True, red=True,
                )
            else:
                p_corr = doc.add_paragraph()
                p_corr.paragraph_format.left_indent = Inches(0.25)
                p_corr.paragraph_format.space_after = Pt(2)
                p_corr.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                _add_run(p_corr, "Corrected: ", bold=True, size_pt=11)
                add_markdown_to_paragraph(p_corr, corrected)

            p_comment = doc.add_paragraph()
            p_comment.paragraph_format.left_indent = Inches(0.25)
            p_comment.paragraph_format.space_after = Pt(2)
            _add_run(p_comment, "Comment: ", bold=True, size_pt=11)

            comment = str(row.get("Explanation", "")).strip()
            if not comment:
                if doi_suspicious and not row.get("DOI Rescued", False):
                    comment = (
                        "DOI does not resolve in OpenAlex to the claimed "
                        "title/authors. Corrected version withheld."
                    )
                elif row.get("DOI Rescued", False):
                    comment = (
                        "DOI verified — title and author matched "
                        "(fuzzy match). No change needed."
                    )
                elif display_status == "MATCH":
                    comment = "Reference is consistent with Chicago style."
                elif display_status == "REVISED":
                    comment = "Reference was revised according to Chicago style."
                else:
                    comment = "Manual verification is recommended."

            _add_run(
                p_comment, comment, size_pt=10, italic=True,
                red=(display_status in {"MANUAL CHECK", "NOT CHECKED"} or doi_suspicious),
            )

            if row.get("OpenAlex Title"):
                p_oa = doc.add_paragraph()
                p_oa.paragraph_format.left_indent = Inches(0.25)
                p_oa.paragraph_format.space_after = Pt(2)
                r = p_oa.add_run(f'  OpenAlex says: "{row["OpenAlex Title"]}"')
                r.italic = True
                r.font.size = Pt(10)

            if row.get("OpenAlex Authors"):
                p_oa2 = doc.add_paragraph()
                p_oa2.paragraph_format.left_indent = Inches(0.25)
                p_oa2.paragraph_format.space_after = Pt(2)
                r = p_oa2.add_run(f"  OpenAlex authors: {row['OpenAlex Authors']}")
                r.italic = True
                r.font.size = Pt(10)

            if row.get("DOI Verification Reasons"):
                p_reason = doc.add_paragraph()
                p_reason.paragraph_format.left_indent = Inches(0.25)
                p_reason.paragraph_format.space_after = Pt(2)
                r = p_reason.add_run(f"  ⚠ {row['DOI Verification Reasons']}")
                r.italic = True
                r.bold = True
                r.font.color.rgb = RED
                r.font.size = Pt(10)

            active_ph = [k for k, v in (row.get("Placeholders") or {}).items() if v]
            if active_ph:
                p_ph = doc.add_paragraph()
                p_ph.paragraph_format.left_indent = Inches(0.25)
                p_ph.paragraph_format.space_after = Pt(2)
                r = p_ph.add_run(
                    "Placeholder used — fill in before submission: "
                    + ", ".join(active_ph)
                )
                r.italic = True
                r.bold = True
                r.font.color.rgb = RED
                r.font.size = Pt(10)

            p_status = doc.add_paragraph()
            p_status.paragraph_format.left_indent = Inches(0.25)
            p_status.paragraph_format.space_after = Pt(4)
            _add_run(p_status, "Status: ", bold=True, size_pt=11)
            _add_run(
                p_status, display_status, size_pt=11, bold=True,
                red=(display_status != "MATCH"),
            )

            _add_divider(doc)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# PAGE NUMBER ENRICHMENT FOR FOOTNOTES
# ============================================================

def attach_page_numbers_to_footnotes(footnotes, page_candidates):
    number_to_pages = {}
    for page_data in page_candidates or []:
        page_number = page_data.get("page")
        if page_number is None:
            continue
        for line in page_data.get("footnote_lines", []):
            detected = line_starts_numbered_note(line)
            if detected is None:
                continue
            number = int(detected[0])
            number_to_pages.setdefault(number, set()).add(int(page_number))

    enriched = []
    for note in footnotes or []:
        item = dict(note)
        number = int(item["number"])
        existing = item.get("pages", [])
        if isinstance(existing, (int, float)):
            existing = [int(existing)]
        elif existing is None:
            existing = []
        pages = set()
        for value in existing:
            try:
                pages.add(int(value))
            except Exception:
                pass
        pages.update(number_to_pages.get(number, set()))
        for key in ("page", "start_page", "end_page"):
            value = item.get(key)
            if value is not None:
                try:
                    pages.add(int(value))
                except Exception:
                    pass
        item["pages"] = sorted(pages)
        enriched.append(item)
    return enriched


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


def check_all_chicago_with_gpt(footnotes, references, client):
    footnote_payload = [
        {
            "number": int(note["number"]),
            "text": clean_text(note["text"]),
            "layout_warning": bool(note.get("layout_warning", False)),
        }
        for note in footnotes
    ]
    bibliography_payload = [
        {"number": i, "text": clean_text(reference)}
        for i, reference in enumerate(references, start=1)
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
            "footnotes": [item.model_dump() for item in parsed.footnotes],
            "bibliography": [item.model_dump() for item in parsed.bibliography],
        }
    except Exception as exc:
        error = f"{type(exc).__name__}: {exc}"
        return {
            "footnotes": [
                {"number": int(note["number"]), "status": "MANUAL_CHECK",
                 "revised_footnote_markdown": "",
                 "explanation": f"Citation service error: {error}"}
                for note in footnotes
            ],
            "bibliography": [
                {"number": i, "source_type": heuristic_source_type(reference),
                 "year": extract_reference_year(reference), "status": "MANUAL_CHECK",
                 "revised_bibliography_markdown": "",
                 "explanation": f"Citation service error: {error}"}
                for i, reference in enumerate(references, start=1)
            ],
        }


# ============================================================
# PIPELINE — one PDF (PyMuPDF extraction + GPT review)
# ============================================================

def process_single_chicago_pdf(uploaded_file, batch, client, manuscript_year):
    fn_result = extract_chicago_footnotes(uploaded_file)
    footnotes = attach_page_numbers_to_footnotes(
        fn_result.get("footnotes", []),
        fn_result.get("page_candidates", []),
    )
    for n in footnotes:
        n["has_dan_warning"] = footnote_has_indonesian_author_conjunction(n["text"])

    bibliography_result = extract_bibliography_lines(uploaded_file)
    if bibliography_result.get("found"):
        references = split_references_from_lines(bibliography_result["lines"])
        references = [clean_text(r) for r in references if clean_text(r)]
    else:
        references = []

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
        ck["decision"] = (
            "✓ OK" if status == "OK"
            else "⚠ REVISED" if status == "REVISED"
            else "⚠ MANUAL CHECK"
        )
        checked_notes.append(ck)

    bibliography_results = combined.get("bibliography", [])

    detail_rows, composition_rows, recency_stats = build_bibliography_statistics(
        references, bibliography_results, int(manuscript_year)
    )

    matched_bibliography_numbers = get_matched_bibliography_numbers(match_rows)
    duplicate_doi_numbers = get_duplicate_doi_reference_numbers(references)

    corrected_bibliography_rows = []
    for row in detail_rows:
        reference_no = int(row["No."])
        original_reference = row.get("Reference", "")
        cited_in_footnotes = reference_no in matched_bibliography_numbers
        duplicate_doi = reference_no in duplicate_doi_numbers

        row["Matched in Footnotes"] = cited_in_footnotes
        row["Duplicate DOI"] = duplicate_doi
        row["GPT Revised"] = normalize_malformed_locator_text(
            canonicalize_doi_url_in_citation(
                row.get("GPT Revised", "") or original_reference
            )
        )
        final_corrected_reference = clean_text(row["GPT Revised"]) or original_reference

        row["Source Type"] = normalize_source_type(
            heuristic_source_type(final_corrected_reference)
        )
        final_year = extract_reference_year(final_corrected_reference)
        if final_year is not None:
            row["Year"] = final_year

        truncated_authors = (
            bibliography_has_truncated_author_list(original_reference)
            or bibliography_has_truncated_author_list(row["GPT Revised"])
        )
        row["Truncated Authors"] = truncated_authors

        chicago_status = str(row.get("Chicago Status", "NOT CHECKED")).upper().strip()
        display_status = (
            "MATCH" if chicago_status == "OK"
            else "REVISED" if chicago_status == "REVISED"
            else "MANUAL CHECK" if chicago_status == "MANUAL_CHECK"
            else chicago_status
        )
        withheld = bool(row.get("DOI Suspicious")) and not bool(row.get("DOI Rescued"))

        corrected_bibliography_rows.append({
            "No.": reference_no,
            "Source Type": row.get("Source Type", "Other"),
            "Publication Year": row["Year"] if row.get("Year") is not None else "—",
            "Original Version": original_reference,
            "Corrected Version": (
                "— WITHHELD (DOI mismatch) —" if withheld else row["GPT Revised"]
            ),
            "Placeholders": ", ".join(
                k for k, v in (row.get("Placeholders") or {}).items() if v
            ) or "—",
            "Status": display_status,
            "DOI Checked": "YES" if row.get("DOI Verified") else "NO",
            "DOI Suspicious": "⚠️ YES" if withheld else "—",
            "OpenAlex Title": (row.get("OpenAlex Title") or "")[:60],
            "DOI Issues": row.get("DOI Verification Reasons", ""),
            "Footnote in Bibliography": "☑ Checked" if cited_in_footnotes else "☐ Unchecked",
            "Bibliography Missing from Footnotes": "No" if cited_in_footnotes else "Yes",
            "Duplicate DOI": "⚠ Yes" if duplicate_doi else "No",
            "Incomplete Author List": "⚠ Yes" if truncated_authors else "No",
        })

    total_references = len(detail_rows)
    final_source_counts = Counter(r.get("Source Type", "Other") for r in detail_rows)
    composition_rows = [
        {"Source Type": source_type, "Count": count,
         "Percentage": round(count / total_references * 100, 1) if total_references else 0.0}
        for source_type, count in sorted(
            final_source_counts.items(), key=lambda item: (-item[1], item[0])
        )
    ]

    cutoff = int(manuscript_year) - 9
    detected_year_rows = [row for row in detail_rows if row.get("Year") is not None]
    for row in detail_rows:
        year_value = row.get("Year")
        row["Within Last 10 Years"] = (
            bool(cutoff <= int(year_value) <= int(manuscript_year))
            if year_value is not None else None
        )
    recent_count = sum(bool(row.get("Within Last 10 Years")) for row in detected_year_rows)
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
        composition_rows=composition_rows,
        manuscript_year=int(manuscript_year),
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
# RENDER — APA-style upload → year → button → selector flow
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
        st.warning("OA key was not found — DOI verification will be skipped.")

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
                uf.seek(0)
                pdf_bytes = uf.read()
                uf.seek(0)

                bib_result = extract_bibliography_lines(uf)
                uf.seek(0)

                if bib_result.get("found") and bib_result.get("lines"):
                    reference_text = "\n".join(
                        line["text"] for line in bib_result["lines"]
                    )
                else:
                    reference_text = ""

                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                full_pages = [page.get_text("text") for page in doc]
                if bib_result.get("found") and bib_result.get("start_page"):
                    body_pages = full_pages[: bib_result["start_page"] - 1]
                else:
                    body_pages = full_pages
                body_text = "\n".join(body_pages)
                full_text = "\n".join(full_pages)
                doc.close()
                uf.seek(0)

                ref_found = bool(
                    bib_result.get("found") and len(bib_result.get("lines", [])) > 0
                )
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
                "cleaned_text": body_text,
                "body_text": body_text,
                "reference_text": reference_text,
                "ref_found": ref_found,
                "post_found": False,
                "manuscript_year": int(manuscript_year),
                "ai_done": False,
            }
            st.session_state["chicago_batches"][key] = batch

            overall.progress(
                base_pct + int(step * 0.35),
                text=f"[{i}/{n}] Auditing Chicago footnotes & bibliography...",
            )
            try:
                uf.seek(0)
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
    doi_suspicious_pct = doi_suspicious / total_refs * 100 if total_refs else 0.0

    overview_df = pd.DataFrame([
        {"Metric": "Total References", "Value": str(total_refs)},
        {"Metric": "Citations > 15",
         "Value": f"Yes ({citation_count})" if citation_count > 15 else f"No ({citation_count})"},
        {"Metric": "% Last 10 Years", "Value": f"{recency_stats.get('recent_pct_all', 0.0)}%"},
        {"Metric": "Footnotes Missing from Bibliography",
         "Value": str(len(batch.get("missing_rows", [])))},
        {"Metric": "Bibliography Missing from Footnotes",
         "Value": str(total_refs - len(get_matched_bibliography_numbers(match_rows)))},
        {"Metric": "DOI Checked (OpenAlex)", "Value": str(doi_checked)},
        {"Metric": "DOI Suspicious (possible fabrication)",
         "Value": f"{doi_suspicious} ({doi_suspicious_pct:.1f}%)"},
    ])

    source_lookup = {r["Source Type"]: f"{r['Count']} ({r['Percentage']}%)" for r in composition_rows}
    composition_df = pd.DataFrame([
        {"Source Type": st_, "Count / Percentage": source_lookup.get(st_, "0 (0.0%)")}
        for st_ in ["Journal Article", "Book", "Government / Legal",
                    "Report", "News / Newspaper", "Website", "Other"]
    ])

    left, right = st.columns(2, gap="large")
    with left:
        st.dataframe(overview_df, use_container_width=True, hide_index=True)
    with right:
        st.dataframe(composition_df, use_container_width=True, hide_index=True)

    show_fn = st.toggle("Footnote Comparison", value=False,
                        key=f"chicago_show_fn_{selected_key}")
    if show_fn:
        fn_rows = []
        for note in checked_notes:
            raw = str(note.get("ai_status", "MANUAL_CHECK")).upper().strip()
            disp = ("MATCH" if raw == "OK"
                    else "REVISED" if raw == "REVISED"
                    else "MANUAL CHECK")
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
            st.dataframe(pd.DataFrame(fn_rows), use_container_width=True,
                         hide_index=True, height=230)
        else:
            st.info("No footnotes available for comparison.")

    show_bib = st.toggle("Bibliography Comparison", value=False,
                         key=f"chicago_show_bib_{selected_key}")
    if show_bib:
        if corrected_rows:
            df = pd.DataFrame(corrected_rows)
            preferred = [
                "No.", "Source Type", "Publication Year",
                "Original Version", "Corrected Version", "Status",
                "Footnote in Bibliography", "Bibliography Missing from Footnotes",
                "Duplicate DOI", "Incomplete Author List", "Placeholders",
                "DOI Checked", "DOI Suspicious", "OpenAlex Title", "DOI Issues",
            ]
            existing = [c for c in preferred if c in df.columns]
            rest = [c for c in df.columns if c not in existing]
            df = df[existing + rest]
            st.caption(
                f"Showing {len(df)} of "
                f"{len(batch.get('references', []))} extracted entries."
            )
            st.dataframe(df, use_container_width=True, hide_index=True, height=280)
        else:
            st.info("No bibliography entries available.")

    report_docx = batch.get("report_docx")
    if report_docx:
        safe_name = re.sub(r"[^\w\-]+", "_", batch.get("filename", "manuscript"))
        st.download_button(
            label="📄 Download Chicago Diagnostic Report (.docx)",
            data=report_docx,
            file_name=f"{safe_name}_chicago_diagnostic_report.docx",
            mime=("application/vnd.openxmlformats-officedocument."
                  "wordprocessingml.document"),
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
        }
        div[class*="st-key-reset_btn"] button:hover {
            background-color: #b91c1c !important;
            color: #ffffff !important;
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