# ieee.py
# OmniCite Auditor — IEEE Style Module
# Imported and rendered by app.py via:  import ieee; ieee.render()

import re
import os
import io
import json
import difflib
import traceback
from datetime import datetime
from collections import Counter

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
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
                "page_height": page_height, "page_width": page_width,
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
                    "Reason": "Repeated header/footer",
                })
                continue

            if _safe_fullmatch(r"\s*(?:page\s*)?\d+\s*", text, flags=re.IGNORECASE):
                if is_top or is_bottom:
                    removed_running_text.append({
                        "Page": page_data["page"], "Text": text,
                        "Reason": "Page number",
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
                        "flags": flags, "italic": italic, "bbox": span.get("bbox"),
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

REFERENCE_HEADINGS = ["references", "reference", "bibliography", "daftar pustaka"]

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
# IEEE REFERENCE SPLITTING
# =========================================================

IEEE_MARKER_RE = re.compile(r"^\s*\[(\d+)\]\s*")
IEEE_PLAIN_NUM_RE = re.compile(r"^\s*(\d{1,3})\.\s+(?=[A-Z])")


def is_page_break(line):
    return bool(re.fullmatch(r"<<<PAGE_BREAK:\d+>>>", line.strip()))


def split_references(reference_text):
    raw_lines = reference_text.splitlines()
    lines = []
    for raw_line in raw_lines:
        line = raw_line.strip()
        if not line:
            continue
        if is_page_break(line):
            continue
        line = line.replace("\u00ad", "")
        line = re.sub(r"\s+", " ", line).strip()
        if line:
            lines.append(line)

    references = []
    current = []

    def flush():
        nonlocal current
        if not current:
            return
        ref = re.sub(r"\s+", " ", " ".join(current)).strip()
        ref = re.sub(r"\s*\[(\d+)\]\s*$", "", ref).strip()
        if ref:
            references.append(ref)
        current = []

    for line in lines:
        if IEEE_MARKER_RE.match(line) or IEEE_PLAIN_NUM_RE.match(line):
            flush()
            current = [line]
        else:
            if not current:
                current = [line]
            else:
                current.append(line)
    flush()

    cleaned = []
    for ref in references:
        ref = re.sub(r"^\s*\[\d+\]\s*", "", ref).strip()
        ref = re.sub(r"^\s*\d{1,3}\.\s+", "", ref).strip()
        if ref:
            cleaned.append(ref)
    return cleaned


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

    pp_match = re.search(r"\bpp\.\s*([\w\-–—]+(?:\s*[–—-]\s*[\w\-–—]+)?)", reference)
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
    if re.search(r"\b(proc\.|proceedings|conference|symposium|workshop)\b", low):
        result["source_type"] = "Conference Paper"
    elif re.search(r"\b(ieee|acm)\s+trans\.", low) or re.search(r"\bjournal\b", low):
        result["source_type"] = "Journal Article"
    elif re.search(r"\b(vol\.|no\.|pp\.)\b", low) and result["venue"]:
        result["source_type"] = "Journal Article"
    elif re.search(r"\b(book|monograph|handbook)\b", low):
        result["source_type"] = "Book"
    elif re.search(r"\b(tech\.?\s*rep\.?|technical report|report)\b", low):
        result["source_type"] = "Technical Report"
    elif re.search(r"\b(ed\.|eds\.|chapter)\b", low):
        result["source_type"] = "Book Chapter"
    elif re.search(r"https?://", low) and not result["doi"]:
        result["source_type"] = "Web Page"
    elif result["doi"]:
        result["source_type"] = "Journal Article"

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


def _insert_placeholder_after_venue(ref, placeholder):
    m = re.search(r",\s*(?:vol\.|no\.|pp\.|doi:|https?://)", ref, re.I)
    if m:
        return ref[:m.start()] + f", {placeholder}" + ref[m.start():]
    return ref.rstrip(".").rstrip() + f", {placeholder}."


def _insert_placeholder_after_volume(ref, placeholder):
    m = re.search(r"\bvol\.\s*\S+", ref, re.I)
    if m:
        return ref[:m.end()] + f", {placeholder}" + ref[m.end():]
    return _insert_placeholder_after_venue(ref, placeholder)


def _insert_placeholder_after_issue(ref, placeholder):
    m = re.search(r"\bno\.\s*\S+", ref, re.I)
    if m:
        return ref[:m.end()] + f", {placeholder}" + ref[m.end():]
    return _insert_placeholder_after_volume(ref, placeholder)


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

    parsed_now = parse_ieee_reference(ref)

    if parsed_now["source_type"] == "Journal Article":
        if not parsed_now["volume"] and "vol." not in ref.lower():
            ref = _insert_placeholder_after_venue(ref, PLACEHOLDER_VOL)
            placeholder_flags["missing_vol"] = True
            notes.append(f"Missing volume — placeholder inserted: {PLACEHOLDER_VOL}")

        if not parsed_now["issue"] and "no." not in ref.lower():
            ref = _insert_placeholder_after_volume(ref, PLACEHOLDER_NO)
            placeholder_flags["missing_no"] = True
            notes.append(f"Missing issue — placeholder inserted: {PLACEHOLDER_NO}")

        if not parsed_now["pages"] and "pp." not in ref.lower():
            ref = _insert_placeholder_after_issue(ref, PLACEHOLDER_PAGES)
            placeholder_flags["missing_pp"] = True
            notes.append(f"Missing pages — placeholder inserted: {PLACEHOLDER_PAGES}")
        elif _is_single_page_range(parsed_now["pages"]):
            single = re.search(
                r"\b(p+p?\.)\s*(\d+)\b",
                ref,
                re.I,
            )
            if single:
                ref = (
                    ref[:single.start()]
                    + f"pp. {single.group(2)}-???"
                    + ref[single.end():]
                )
                placeholder_flags["missing_pp"] = True
                notes.append(
                    f"Single page number — expanded to range placeholder: "
                    f"pp. {single.group(2)}-???"
                )

    if not parsed_now["doi"] and not parsed_now["url"]:
        if "doi:" not in ref.lower() and "http" not in ref.lower():
            ref = ref.rstrip(".").rstrip() + f", {PLACEHOLDER_DOI}."
            placeholder_flags["missing_doi"] = True
            notes.append(f"Missing DOI/URL — placeholder inserted: {PLACEHOLDER_DOI}")

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
# AI REVIEW (OPTIONAL)
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


def review_ieee_references_with_ai(references):
    client = _get_openai_client()
    if client is None or not references:
        return []

    payload = [{"number": i, "reference": clean_text(r)} for i, r in enumerate(references, start=1)]

    prompt = f"""
You are checking IEEE-style REFERENCE-LIST entries.

IEEE STYLE RULES:
- Author names: "A. B. Smith" (initials first). Full given names are wrong.
- All authors must be listed; "et al." is NOT allowed.
- Article titles: double quotes, Title Case.
- Journal / conference names: italic (list them in "italic_elements").
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


def build_ieee_reference_comparison(references, ai_results=None, manuscript_year=None):
    ai_results = ai_results or []
    by_no = {
        int(x.get("number", -1)): x
        for x in ai_results
        if str(x.get("number", "")).isdigit()
    }
    rows = []
    for i, original in enumerate(references, start=1):
        ai = by_no.get(i, {})
        original_clean = strip_markdown_markers(clean_text(original))

        structure_errors = ieee_structure_errors(original_clean)

        corrected = strip_markdown_markers(
            clean_text(ai.get("revised_reference", ""))
        ) or original_clean
        if corrected != original_clean and _reference_is_hallucinated(original_clean, corrected):
            corrected = original_clean

        local = build_local_ieee_reference_correction(corrected)
        corrected = local["Corrected"]

        parsed = parse_ieee_reference(corrected)

        verification = verify_reference_against_openalex(original_clean, parsed)

        source_type = normalize_source_type(ai.get("source_type"))
        if source_type == "Other" and not ai.get("source_type"):
            source_type = parsed["source_type"]

        italic_elements = (ai.get("italic_elements") or "").strip()
        if not italic_elements:
            italic_elements = _fallback_ieee_italic_elements(source_type, parsed)

        year = ai.get("year")
        if not isinstance(year, int):
            year = int(parsed["year"]) if parsed["year"] else None

        missing = ai.get("missing_required_elements", []) or []
        if isinstance(missing, str):
            missing = [missing]
        missing = list({*missing, *ieee_missing_elements(corrected)})

        status = _ieee_status_label(ai.get("status")) if ai else "NOT AI CHECKED"

        rows.append({
            "No.": i,
            "Source Type": source_type,
            "Year": year,
            "Original Reference": original_clean,
            "Corrected Version": corrected,
            "Italicized in IEEE": italic_elements,
            "Status": status,
            "Missing Required Elements": ", ".join(str(x) for x in missing),
            "AI Explanation": ai.get("explanation", ""),
            "Original Structure Errors": " | ".join(structure_errors),
            "Original Has Structure Error": bool(structure_errors),
            "Correction Note": local.get("Note", ""),
            "DOI Verified": verification["checked"],
            "DOI Suspicious": verification["suspicious"],
            "DOI Verification Reasons": " | ".join(verification["reasons"]),
            "Title Similarity": verification.get("title_similarity"),
            "Author Overlap": verification.get("author_overlap"),
            "OpenAlex Title": verification.get("crossref_title"),
            "OpenAlex Authors": ", ".join(verification.get("crossref_authors", [])[:5]),
            "Placeholders": local.get("Placeholders", {}),
        })
    return rows


def _fallback_ieee_italic_elements(source_type, parsed):
    if source_type in {"Journal Article", "Conference Paper"} and parsed.get("venue"):
        return parsed["venue"]
    if source_type in {"Book", "Book Chapter", "Technical Report"} and parsed.get("title"):
        return parsed["title"]
    return ""


def _ieee_status_label(status):
    status = str(status or "MANUAL_CHECK").upper().strip()
    if status in {"OK", "PASS", "MATCH"}:
        return "MATCH"
    if status in {"REVISED", "NEEDS REVIEW", "NEEDS_REVIEW"}:
        return "REVISED"
    return "MANUAL CHECK"


# =========================================================
# ANALYZE
# =========================================================

def analyze_ieee_locally(uploaded_file, manuscript_year):
    uploaded_file.seek(0)
    full_text, pages, removed_running_text = extract_pdf_text(uploaded_file)
    uploaded_file.seek(0)
    style_spans = extract_pdf_style_spans(uploaded_file)
    full_text = clean_text(full_text)

    reference_text, heading, body_text = find_reference_section(full_text)
    if reference_text is None:
        raise ValueError("I could not detect a References section.")

    references = split_references(reference_text)
    parsed_references = [parse_ieee_reference(r) for r in references]
    citations, citation_clusters = extract_ieee_citations(body_text)

    for c in citations:
        c["page"] = _page_for_offset(full_text, c.get("offset"))

    citation_stats = calculate_citation_statistics(citations, citation_clusters)
    matching_results = match_citations_to_references(citations, parsed_references)
    orphan_citations = find_orphan_citations(citations, parsed_references)
    uncited_references = detect_uncited_references(citations, parsed_references)
    duplicates = detect_duplicates(parsed_references)
    recency = calculate_reference_recency(references, manuscript_year)

    local_reference_checks = []
    for i, ref in enumerate(references, start=1):
        status, _, _, parsed = check_ieee_reference(ref, style_spans)
        local_reference_checks.append({
            "Reference #": i,
            "IEEE Status": status,
            "Source Type": parsed.get("source_type", "Other"),
            "Structure Errors": " | ".join(ieee_structure_errors(ref)),
        })

    return {
        "filename": uploaded_file.name,
        "heading": heading,
        "references": references,
        "parsed_references": parsed_references,
        "citations": citations,
        "citation_clusters": citation_clusters,
        "citation_stats": citation_stats,
        "matching_results": matching_results,
        "orphan_citations": orphan_citations,
        "uncited_references": uncited_references,
        "duplicates": duplicates,
        "recency": recency,
        "local_reference_checks": local_reference_checks,
        "removed_running_text": removed_running_text,
        "manuscript_year": manuscript_year,
        "ai_complete": False,
    }


def check_ieee_reference(reference, style_spans=None):
    issues = ieee_structure_errors(reference)
    warnings = []
    parsed = parse_ieee_reference(reference)
    if issues:
        status = "Fail"
    elif warnings:
        status = "Review"
    else:
        status = "Pass"
    return status, issues, warnings, parsed


def enrich_ieee_with_ai(result):
    reference_ai = review_ieee_references_with_ai(result["references"])
    result["reference_comparison"] = build_ieee_reference_comparison(
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
    _set_run_font(
        r,
        size_pt=size_pt,
        bold=bold,
        italic=italic,
        color=RED if red else None,
    )
    return r


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


def _emit_with_placeholders(paragraph, text, italic=False):
    pos = 0
    for m in _PLACEHOLDER_RE.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], size_pt=11, italic=italic)

        matched = m.group(0)

        page_partial = _PAGE_PARTIAL_RE.match(matched)
        if page_partial:
            known = page_partial.group(1)
            prefix = matched[: matched.index(known)]
            _add_run(paragraph, prefix, size_pt=11, italic=italic)
            _add_run(paragraph, f"{known}-???", size_pt=11,
                     italic=True, bold=True, red=True)
        else:
            _add_run(paragraph, matched, size_pt=11,
                     italic=True, bold=True, red=True)

        pos = m.end()
    if pos < len(text):
        _add_run(paragraph, text[pos:], size_pt=11, italic=italic)


def _add_ieee_reference_with_italics(paragraph, text, italic_elements):
    tokens = [t.strip() for t in (italic_elements or "").split(",") if t.strip()]
    if not tokens:
        _emit_with_placeholders(paragraph, text, italic=False)
        return

    flat = [_esc(t) for t in tokens]
    pattern_str = "|".join(sorted(flat, key=len, reverse=True))
    pattern = _safe_compile(pattern_str, re.I)
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


def build_ieee_docx(result):
    doc = Document()
    _docx_set_default_font(doc)

    filename = result.get("filename", "manuscript.pdf")

    title = doc.add_heading(level=0)
    tr = title.add_run(filename)
    _set_run_font(tr, size_pt=18, bold=True)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    sr = subtitle.add_run("IEEE Reference Style — Citation & Reference Correction Report")
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

    citations = result.get("citations", [])
    clusters = result.get("citation_clusters", [])

    if not citations:
        _add_run(doc.add_paragraph(), "No bracketed in-text citations were detected.", italic=True)
    else:
        for cluster in clusters:
            numbers = cluster["numbers"]
            raw = cluster["raw"]
            canonical = cluster["canonical"]
            page = None
            for n in numbers:
                for c in citations:
                    if c["number"] == n:
                        page = c.get("page")
                        break
                if page:
                    break

            needs_fix, reasons = evaluate_cluster(cluster)

            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(3)
            _add_run(p, f"Page {page if page else '?'}  ", bold=True, size_pt=11)
            _add_run(p, f"Lookup: {raw}", size_pt=11, red=needs_fix)
            if needs_fix:
                _add_run(p, f"   ← {reasons[0]}", italic=True, size_pt=10, red=True)

            p2 = doc.add_paragraph()
            p2.paragraph_format.space_after = Pt(8)
            _add_run(p2, "Original : ", bold=True, size_pt=11)
            _add_run(p2, raw, size_pt=11, red=needs_fix)
            p3 = doc.add_paragraph()
            p3.paragraph_format.space_after = Pt(8)
            _add_run(p3, "Corrected: ", bold=True, size_pt=11)
            _add_run(p3, canonical, size_pt=11)

    doc.add_page_break()

    # ---------- Section 2: Reference List ----------
    h2 = doc.add_heading(level=1)
    hr2 = h2.add_run("2. Reference List (IEEE Style)")
    _set_run_font(hr2, size_pt=14, bold=True)

    reference_rows = result.get("reference_comparison") or build_ieee_reference_comparison(
        result.get("references", []), [], result.get("manuscript_year")
    )

    matching = result.get("matching_results", []) or []
    uncited_numbers = {row.get("Reference #") for row in matching if not row.get("Cited")}

    if not reference_rows:
        _add_run(doc.add_paragraph(), "No references were detected.", italic=True)
    else:
        for row in reference_rows:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.left_indent = Inches(0.4)
            p.paragraph_format.first_line_indent = Inches(-0.4)

            _add_run(p, f"[{row.get('No.')}] ", bold=True, size_pt=11)

            original = row.get("Original Reference", "")
            original_bad = row.get("Original Has Structure Error", False)
            if original_bad:
                _add_run(p, original, size_pt=11, red=True)
                _add_run(
                    p,
                    f"   ← IEEE structure issue: {row.get('Original Structure Errors','')}",
                    size_pt=10, italic=True, red=True,
                )
            else:
                _add_run(p, original, size_pt=11)

            if row.get("No.") in uncited_numbers:
                _add_run(p, "   ← NOT CITED IN TEXT", size_pt=10, bold=True, italic=True, red=True)

            if row.get("DOI Suspicious"):
                withheld = doc.add_paragraph()
                withheld.paragraph_format.left_indent = Inches(0.4)
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
                p2.paragraph_format.space_after = Pt(8)
                p2.paragraph_format.left_indent = Inches(0.4)
                _add_run(p2, "Corrected: ", bold=True, size_pt=11)
                corrected = row.get("Corrected Version", "")
                italic_elements = row.get("Italicized in IEEE", "")
                _add_ieee_reference_with_italics(p2, corrected, italic_elements)

                correction_note = row.get("Correction Note", "")
                if correction_note:
                    note_p = doc.add_paragraph()
                    note_p.paragraph_format.left_indent = Inches(0.4)
                    _add_run(note_p, f"Fix applied: {correction_note}",
                             size_pt=10, italic=True)

                placeholders = row.get("Placeholders") or {}
                active_ph = [k for k, v in placeholders.items() if v]
                if active_ph:
                    ph_p = doc.add_paragraph()
                    ph_p.paragraph_format.left_indent = Inches(0.4)
                    _add_run(ph_p,
                             "Placeholder used — fill in before submission: "
                             + ", ".join(active_ph),
                             size_pt=10, italic=True, bold=True, red=True)

            if row.get("DOI Suspicious"):
                warn_p = doc.add_paragraph()
                warn_p.paragraph_format.left_indent = Inches(0.4)
                reasons = row.get("DOI Verification Reasons", "")
                _add_run(warn_p, f"⚠ Possible fabricated reference: {reasons}",
                         size_pt=10, italic=True, bold=True, red=True)
                if row.get("OpenAlex Title"):
                    _add_run(warn_p, f"\n  OpenAlex says: \"{row['OpenAlex Title']}\"",
                             size_pt=10, italic=True)
                if row.get("OpenAlex Authors"):
                    _add_run(warn_p, f"\n  Authors: {row['OpenAlex Authors']}",
                             size_pt=10, italic=True)

            missing = row.get("Missing Required Elements", "")
            if missing:
                note_p = doc.add_paragraph()
                note_p.paragraph_format.left_indent = Inches(0.4)
                _add_run(note_p, f"Missing element(s): {missing}",
                         size_pt=10, italic=True, red=True)

    # ---------- Section 3: Cross-link Discrepancies ----------
    doc.add_page_break()
    h3 = doc.add_heading(level=1)
    hr3 = h3.add_run("3. Cross-link Discrepancies")
    _set_run_font(hr3, size_pt=14, bold=True)

    orphan = result.get("orphan_citations", []) or []
    uncited = result.get("uncited_references", []) or []

    if orphan:
        sub = doc.add_heading(level=2)
        _add_run(sub.add_run("Citations Missing from References"), bold=True, size_pt=12)
        for o in orphan:
            p = doc.add_paragraph()
            _add_run(p, f"{o.get('Citation','')}  ", size_pt=11, red=True)
            _add_run(p, f"— {o.get('Problem','')}", size_pt=10, italic=True, red=True)

    if uncited:
        sub = doc.add_heading(level=2)
        _add_run(sub.add_run("References Missing from Citations"), bold=True, size_pt=12)
        for u in uncited:
            p = doc.add_paragraph()
            _add_run(p, f"[{u.get('Reference #')}] {u.get('Reference','')}  ",
                     size_pt=11, red=True)
            _add_run(p, f"— {u.get('Problem','')}", size_pt=10, italic=True, red=True)

    if not orphan and not uncited:
        _add_run(doc.add_paragraph(), "No cross-link discrepancies detected.", italic=True)

    # ---------- Section 4: Summary ----------
    doc.add_page_break()
    h4 = doc.add_heading(level=1)
    hr4 = h4.add_run("4. Summary")
    _set_run_font(hr4, size_pt=14, bold=True)

    stats = result.get("citation_stats", {})
    recency = result.get("recency", {})

    total_refs_now = len(result.get("references", []))
    doi_checked = sum(1 for r in reference_rows if r.get("DOI Verified"))
    doi_suspicious = sum(1 for r in reference_rows if r.get("DOI Suspicious"))
    doi_suspicious_pct = (
        doi_suspicious / total_refs_now * 100 if total_refs_now else 0
    )
    withheld_count = doi_suspicious
    single_page_count = sum(
        1 for r in reference_rows
        if (r.get("Placeholders") or {}).get("missing_pp")
    )

    summary_lines = [
        f"Total references: {total_refs_now}",
        f"Total in-text citation markers: {stats.get('total', 0)}",
        f"Unique cited references: {stats.get('unique', 0)}",
        f"Grouped citation clusters detected: {stats.get('clusters', 0)}",
        f"Crowded clusters needing collapse: {stats.get('crowded_clusters', 0)}",
        f"Orphan citations (no matching reference): {len(orphan)}",
        f"Uncited references: {len(uncited)}",
        f"DOI checked via OpenAlex: {doi_checked}",
        f"DOI suspicious (possible fabricated references): "
        f"{doi_suspicious} ({doi_suspicious_pct:.1f}%)",
        f"Corrections withheld due to DOI mismatch: {withheld_count}",
        f"References with incomplete page range: {single_page_count}",
        f"% references within last 10 years "
        f"({recency.get('start_year')}-{recency.get('end_year')}): "
        f"{recency.get('recent_percentage', 0):.1f}%",
    ]
    for line in summary_lines:
        _add_run(doc.add_paragraph(), line, size_pt=11)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


# =========================================================
# RENDER (called from app.py)
# =========================================================

def render():

    st.title("OmniCite Auditor - IEEE Style")
    st.caption("Numbered bracketed in-text citations ↔ IEEE reference list")

    uploaded_files = st.file_uploader(
        "Upload manuscript PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload up to 5 manuscripts in one batch.",
        key="ieee_file_uploader",
    )

    manuscript_year = st.number_input(
        "Manuscript publication year", min_value=1900, max_value=2100,
        value=datetime.now().year, step=1,
        key="ieee_manuscript_year",
    )

    if "ieee_batch_results" not in st.session_state:
        st.session_state["ieee_batch_results"] = {}

    if uploaded_files:
        if len(uploaded_files) > 5:
            st.error(f"You uploaded {len(uploaded_files)} manuscripts. The maximum batch size is 5 PDFs.")
            st.stop()

        with st.container(key="blue_btn_extract_ieee"):
            extract_batch = st.button(
                f"Extract & Review ({len(uploaded_files)} manuscript{'s' if len(uploaded_files) != 1 else ''})",
                use_container_width=True,
                key="ieee_extract_button",
            )

        if extract_batch:
            local_results = {}
            progress = st.progress(0, text="Extracting IEEE citations and references locally...")
            for file_index, uploaded_file in enumerate(uploaded_files, start=1):
                key = f"{file_index}::{uploaded_file.name}"
                try:
                    progress.progress(
                        int(((file_index - 1) / len(uploaded_files)) * 100),
                        text=f"Extracting {file_index}/{len(uploaded_files)}: {uploaded_file.name}",
                    )
                    local_results[key] = analyze_ieee_locally(uploaded_file, int(manuscript_year))
                except Exception:
                    st.error(f"{uploaded_file.name} failed:\n\n{traceback.format_exc()}")
                    local_results[key] = {
                        "filename": uploaded_file.name,
                        "error": "see traceback above",
                        "manuscript_year": int(manuscript_year),
                    }
            progress.empty()
            st.session_state["ieee_batch_results"] = local_results

        results = st.session_state.get("ieee_batch_results", {})

        if results:
            valid_results = {k: v for k, v in results.items() if not v.get("error")}

            if not valid_results:
                for failed in results.values():
                    if failed.get("error"):
                        st.error(f"{failed.get('filename', 'Manuscript')}: {failed['error']}")
            else:
                selector_keys = list(valid_results.keys())
                selected_key = st.selectbox(
                    "Select manuscript to review",
                    selector_keys,
                    format_func=lambda k: valid_results[k].get("filename", k),
                    key="ieee_selected_manuscript",
                )
                result = valid_results[selected_key]

                citations = result["citations"]
                references = result["references"]
                matching = result["matching_results"]
                orphan = result["orphan_citations"]
                uncited = result["uncited_references"]

                st.caption(
                    f"References: {len(references)}  |  "
                    f"In-text citation markers: {len(citations)}"
                )

                with st.expander(f"In-text Citations ({len(citations)})", expanded=False):
                    if citations:
                        df_cit = pd.DataFrame([
                            {
                                "Page": c.get("page"),
                                "Citation": c["raw"],
                                "Context": (c.get("context") or "").strip(),
                            }
                            for c in citations
                        ])
                        st.dataframe(df_cit, use_container_width=True, hide_index=True,
                                     height=min(300, 38 * (len(citations) + 1)))
                    else:
                        st.info("No bracketed IEEE citations detected in body text.")

                clusters = result.get("citation_clusters", [])
                if clusters:
                    rows = []
                    for cl in clusters:
                        needs_fix, reasons = evaluate_cluster(cl)
                        page = None
                        for c in citations:
                            if c["number"] in cl["numbers"]:
                                page = c.get("page")
                                break
                        rows.append({
                            "Page": page,
                            "Original Form": cl["raw"],
                            "Corrected Form": cl["canonical"],
                            "Numbers Cited": ", ".join(str(n) for n in cl["numbers"]),
                            "Status": "REVISED" if needs_fix else "MATCH",
                            "Reason": " | ".join(reasons),
                        })
                    with st.expander(
                        f"Grouped In-text Citations ({len(clusters)}) "
                        f"— {sum(1 for r in rows if r['Status']=='REVISED')} need collapsing",
                        expanded=False,
                    ):
                        st.dataframe(pd.DataFrame(rows), use_container_width=True,
                                     hide_index=True, height=min(320, 38 * (len(rows) + 1)))

                if orphan:
                    with st.expander(f"Citations Missing from References ({len(orphan)})", expanded=False):
                        st.dataframe(pd.DataFrame(orphan), use_container_width=True, hide_index=True)

                if uncited:
                    with st.expander(f"References Missing from Citations ({len(uncited)})", expanded=False):
                        st.dataframe(pd.DataFrame(uncited), use_container_width=True, hide_index=True)

                with st.container(key="green_btn_ai_ieee"):
                    run_ai = st.button(
                        "Automated Processing Check",
                        use_container_width=True,
                        disabled=not bool(_get_openai_client()),
                        help=None if _get_openai_client() else "Set OPENAI_API_KEY to enable automated IEEE correction.",
                        key="run_ieee_ai_review",
                    )

                if run_ai:
                    ai_progress = st.progress(0, text="Running IEEE automated review...")
                    for idx, (key, batch_result) in enumerate(valid_results.items(), start=1):
                        ai_progress.progress(
                            int(((idx - 1) / len(valid_results)) * 100),
                            text=f"IEEE review {idx}/{len(valid_results)}: {batch_result['filename']}",
                        )
                        try:
                            enriched = enrich_ieee_with_ai(batch_result)
                        except Exception:
                            st.error(f"IEEE review failed for {batch_result.get('filename')}:\n\n{traceback.format_exc()}")
                            enriched = batch_result
                        st.session_state["ieee_batch_results"][key] = enriched
                    ai_progress.empty()
                    st.rerun()

                ai_ready = all(v.get("ai_complete", False) for v in valid_results.values())
                if ai_ready:
                    result = valid_results[selected_key]
                    reference_rows = result.get("reference_comparison") or []
                    stats = result["citation_stats"]
                    recency = result["recency"]

                    total_refs_now = len(result["references"])
                    doi_checked = sum(1 for r in reference_rows if r.get("DOI Verified"))
                    doi_suspicious = sum(1 for r in reference_rows if r.get("DOI Suspicious"))
                    doi_suspicious_pct = (
                        doi_suspicious / total_refs_now * 100 if total_refs_now else 0
                    )

                    metric_rows = [
                        {"Metric": "Total References", "Value": total_refs_now},
                        {"Metric": "Total In-text Citation Markers", "Value": stats.get("total", 0)},
                        {"Metric": "Unique Cited References", "Value": stats.get("unique", 0)},
                        {"Metric": "Grouped Citation Clusters", "Value": stats.get("clusters", 0)},
                        {"Metric": "Crowded Clusters (needs collapse)", "Value": stats.get("crowded_clusters", 0)},
                        {"Metric": "Orphan Citations", "Value": len(result["orphan_citations"])},
                        {"Metric": "Uncited References", "Value": len(result["uncited_references"])},
                        {"Metric": "DOI Checked (OpenAlex)", "Value": doi_checked},
                        {"Metric": "DOI Suspicious (possible fabrication)",
                         "Value": f"{doi_suspicious} ({doi_suspicious_pct:.1f}%)"},
                        {"Metric": "% Last 10 Years", "Value": f"{recency.get('recent_percentage', 0):.1f}%"},
                    ]
                    metric_df = pd.DataFrame(metric_rows)

                    source_counts = Counter(
                        (row.get("Source Type") or "Other") for row in reference_rows
                    )
                    source_rows = []
                    total_refs = len(result["references"])
                    for source_type in CANONICAL_SOURCE_TYPES:
                        count = source_counts.get(source_type, 0)
                        pct = count / total_refs * 100 if total_refs else 0
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

                    show_reference = st.toggle(
                        "Show Reference Correction",
                        value=False,
                        key="show_ieee_reference_correction_toggle",
                    )
                    if show_reference:
                        reference_display = pd.DataFrame([
                            {
                                "No.": row.get("No."),
                                "Corrected Version (IEEE)": (
                                    "— WITHHELD (DOI mismatch) —"
                                    if row.get("DOI Suspicious")
                                    else row.get("Corrected Version", "")
                                ),
                                "Placeholders": ", ".join(
                                    k for k, v in (row.get("Placeholders") or {}).items() if v
                                ) or "—",
                                "DOI Checked": "YES" if row.get("DOI Verified") else "NO",
                                "DOI Suspicious": "⚠️ YES" if row.get("DOI Suspicious") else "—",
                                "OpenAlex Title": (row.get("OpenAlex Title") or "")[:60],
                                "DOI Issues": row.get("DOI Verification Reasons", ""),
                            }
                            for row in reference_rows
                        ])
                        if not reference_display.empty:
                            st.dataframe(reference_display, use_container_width=True,
                                         hide_index=True, height=280)
                        else:
                            st.info("No references were available for automated review.")

                    try:
                        docx_bytes = build_ieee_docx(result)
                        safe_name = re.sub(r"[^\w\-]+", "_", result.get("filename", "manuscript"))
                        st.markdown('<div class="apa-green-button-marker"></div>', unsafe_allow_html=True)
                        st.download_button(
                            label="Download Diagnostic Report",
                            data=docx_bytes,
                            file_name=f"{safe_name}_diagnostic_report.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key="download_ieee_correction_docx",
                        )
                    except Exception as exc:
                        st.error(f"Could not build DOCX report: {exc}")