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
from datetime import datetime
from collections import Counter

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
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
    lines = text.splitlines()

    marker_re = re.compile(r"^\s*\[?\s*(\d{1,3})\s*\]")
    last_marker_idx = None
    for i in range(len(lines) - 1, -1, -1):
        if marker_re.match(lines[i]):
            last_marker_idx = i
            break

    if last_marker_idx is None:
        return _find_reference_section_topdown(text)

    heading_idx = None
    for i in range(last_marker_idx, -1, -1):
        if _heading_matches_reference_vocab(lines[i]):
            heading_idx = i
            break

    if heading_idx is None:
        reference_text = "\n".join(lines[last_marker_idx:])
        body_text = "\n".join(lines[:last_marker_idx])
        return reference_text, None, body_text

    end_idx = None
    saw_marker = False
    for i in range(heading_idx + 1, len(lines)):
        line = lines[i].strip()
        if not line:
            continue
        if re.fullmatch(r"<<<PAGE_BREAK:\d+>>>", line):
            continue
        if marker_re.match(line):
            saw_marker = True
            continue
        if saw_marker and is_post_reference_heading(line):
            end_idx = i
            break

    heading_found = lines[heading_idx].strip()
    body_text = "\n".join(lines[:heading_idx])

    if end_idx is not None:
        reference_text = "\n".join(lines[heading_idx + 1:end_idx])
    else:
        reference_text = "\n".join(lines[heading_idx + 1:])

    return reference_text, heading_found, body_text


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
# IEEE REFERENCE SPLITTING — bottom-up anchor
# =========================================================

_MARKER_TOKEN_RE = re.compile(
    r"(?:(?<=\s)|^)"
    r"\[?\s*(\d{1,3})\s*\]"
    r"(?=\s|[A-Z]|$)"
)


def _find_all_markers(line):
    hits = []
    for m in _MARKER_TOKEN_RE.finditer(line):
        try:
            num = int(m.group(1))
        except ValueError:
            continue
        if 1900 <= num <= 2099:
            continue
        if num > 500:
            continue
        hits.append((m.start(), num))
    return hits


def _looks_like_reference_start(text_after_marker):
    tail = text_after_marker.lstrip()
    if not tail:
        return False
    if tail[0].isupper() or tail[0] in '"\u201c':
        return True
    return False


def split_references(reference_text):
    if not reference_text:
        return []

    text = reference_text
    text = re.sub(r"<<<PAGE_BREAK:\d+>>>", " ", text)
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u00ad", "")
    text = re.sub(r"\s+", " ", text).strip()

    text = re.sub(r"([a-z])(\d{1,3})\]\s", r"\1 [\2] ", text)

    all_hits = _find_all_markers(text)
    if not all_hits:
        return []

    sequences = []
    current = []
    for pos, num in all_hits:
        if not current:
            current = [(pos, num)]
            continue
        prev_num = current[-1][1]
        if num == prev_num + 1:
            current.append((pos, num))
        elif num > prev_num + 1:
            sequences.append(current)
            current = [(pos, num)]

    if current:
        sequences.append(current)

    if not sequences:
        return []
    best = max(sequences, key=lambda s: (len(s), s[-1][0]))

    markers = best
    references = []
    for i, (pos, num) in enumerate(markers):
        start = pos
        end = markers[i + 1][0] if i + 1 < len(markers) else len(text)
        chunk = text[start:end].strip()
        chunk = re.sub(r"^\[?\s*\d{1,3}\s*\]\s*", "", chunk).strip()
        chunk = re.sub(r"\s+\d{1,3}\s*$", "", chunk).strip()
        if chunk:
            references.append(chunk)

    numbered = sorted(zip([n for _, n in markers], references), key=lambda x: x[0])
    return [ref for _, ref in numbered]


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


# =========================================================
# STAGED PIPELINE
# =========================================================

def process_single_ieee_pdf(uploaded_file, batch, client, manuscript_year):
    uploaded_file.seek(0)
    full_text, pages, removed_running_text = extract_pdf_text(uploaded_file)
    uploaded_file.seek(0)
    style_spans = extract_pdf_style_spans(uploaded_file)
    full_text = clean_text(full_text)

    reference_text, heading, body_text = find_reference_section(full_text)
    if reference_text is None:
        raise ValueError("I could not detect a References section.")

    references = split_references(reference_text)
    parsed_refs = [parse_ieee_reference(r) for r in references]
    local_source_types = [p.get("source_type", "Other") for p in parsed_refs]

    citations, clusters = extract_ieee_citations(body_text)
    for c in citations:
        c["page"] = _page_for_offset(full_text, c.get("offset"))

    verification_rows = []
    for ref, parsed in zip(references, parsed_refs):
        v = verify_reference_against_openalex(ref, parsed)
        verification_rows.append(v)

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

        ai_revised = strip_markdown_markers(clean_text(ai.get("revised_reference", "")))
        corrected = ai_revised or original_clean
        if corrected != original_clean and _reference_is_hallucinated(original_clean, corrected):
            corrected = original_clean

        local = build_local_ieee_reference_correction(corrected)
        corrected = local["Corrected"]

        parsed = parse_ieee_reference(corrected)
        v = verification_rows[i - 1]

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
        if ai.get("status") == "WITHHELD":
            corrected_display = "— WITHHELD —"
        elif v.get("suspicious"):
            corrected_display = "— WITHHELD (DOI mismatch) —"
        else:
            corrected_display = corrected

        rows.append({
            "No.": i,
            "Source Type": source_type,
            "Year": year,
            "Original Reference": original_clean,
            "Corrected Version": corrected_display,
            "Italicized in IEEE": italic_elements,
            "Status": status,
            "Missing Required Elements": ", ".join(str(x) for x in missing),
            "AI Explanation": ai.get("explanation", ""),
            "Original Structure Errors": " | ".join(structure_errors),
            "Original Has Structure Error": bool(structure_errors),
            "Correction Note": local.get("Note", ""),
            "DOI Verified": v["checked"],
            "DOI Suspicious": v["suspicious"],
            "DOI Verification Reasons": " | ".join(v["reasons"]),
            "Title Similarity": v.get("title_similarity"),
            "Author Overlap": v.get("author_overlap"),
            "OpenAlex Title": v.get("crossref_title"),
            "OpenAlex Authors": ", ".join(v.get("crossref_authors", [])[:5]),
            "Placeholders": local.get("Placeholders", {}),
        })

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

    citation_stats = calculate_citation_statistics(citations, clusters)
    matching = match_citations_to_references(citations, parsed_refs)
    orphan = find_orphan_citations(citations, parsed_refs)
    uncited = detect_uncited_references(citations, parsed_refs)
    duplicates = detect_duplicates(parsed_refs)
    recency = calculate_reference_recency(references, manuscript_year)

    batch.update({
        "filename": uploaded_file.name,
        "reference_text": reference_text,
        "heading": heading,
        "references": references,
        "parsed_references": parsed_refs,
        "citations": citations,
        "citation_clusters": clusters,
        "citation_stats": citation_stats,
        "matching_results": matching,
        "orphan_citations": orphan,
        "uncited_references": uncited,
        "duplicates": duplicates,
        "recency": recency,
        "removed_running_text": removed_running_text,
        "manuscript_year": int(manuscript_year),
        "reference_comparison": rows,
        "cluster_rows": cluster_rows,
        "verification_rows": verification_rows,
        "ai_done": True,
        "ai_complete": True,
    })


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


def _add_ieee_reference_with_italics(paragraph, text, italic_elements, missing_tokens=None, size_pt=11):
    """Emit corrected reference; italicize venue/title tokens; skip DOI spans;
    highlight placeholders AND missing marker tokens in red bold italic."""
    doi_spans = []
    for m in re.finditer(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", text):
        doi_spans.append((m.start(), m.end()))

    def _in_doi(pos):
        return any(s <= pos < e for s, e in doi_spans)

    if missing_tokens is None:
        missing_tokens = []

    tokens = [t.strip() for t in (italic_elements or "").split(",") if t.strip()]
    if not tokens:
        _highlight_missing_tokens_in_corrected(paragraph, text, missing_tokens, size_pt=size_pt)
        return

    flat = [_esc(t) for t in tokens]
    pattern_str = "|".join(sorted(flat, key=len, reverse=True))
    pattern = _safe_compile(pattern_str, re.I)
    if pattern is None:
        _highlight_missing_tokens_in_corrected(paragraph, text, missing_tokens, size_pt=size_pt)
        return

    pos = 0
    for m in pattern.finditer(text):
        if _in_doi(m.start()):
            continue
        if m.start() > pos:
            _highlight_missing_tokens_in_corrected(
                paragraph, text[pos:m.start()], missing_tokens, size_pt=size_pt
            )
        # Italic span — render with italic on the whole span, but still
        # highlight missing tokens inside it.
        _highlight_missing_tokens_in_corrected(
            paragraph, m.group(0), missing_tokens, size_pt=size_pt
        )
        pos = m.end()

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
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Engine: IEEE-P1-P4-v2"
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

    summary_rows = [
        ("Manuscript publication year", str(manuscript_year), False),
        ("Total references", str(total_refs), False),
        ("Total in-text citation markers", str(total_cits), False),
        ("  • Unique cited references", str(stats.get("unique", 0)), False),
        ("  • Grouped citation clusters", str(stats.get("clusters", 0)), False),
        ("Crowded clusters needing collapse",
         str(stats.get("crowded_clusters", 0)),
         stats.get("crowded_clusters", 0) > 0),
        ("Orphan citations (no matching reference)",
         str(len(orphan)), len(orphan) > 0),
        ("Uncited references",
         str(len(uncited)), len(uncited) > 0),
        ("DOI checked via OpenAlex", str(doi_checked), False),
        ("DOI suspicious (possible fabricated references)",
         f"{doi_suspicious} ({doi_suspicious_pct:.1f}%)",
         doi_suspicious > 0),
        ("Corrections withheld due to DOI mismatch",
         str(withheld_count), withheld_count > 0),
        ("References flagged MANUAL CHECK",
         str(manual_count), manual_count > 0),
        ("References with placeholder fields",
         str(placeholder_count), placeholder_count > 0),
        (f"% references within last 10 years ({window_start}-{window_end})",
         f"{recency.get('recent_percentage', 0):.1f}%", False),
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

            type_run = head.add_run("[Bracketed]")
            _set_run_font(type_run, size_pt=11, bold=True)

            page_no = row.get("Page")
            if page_no:
                _add_run(head, f"  — Page {page_no}", size_pt=10, italic=True)

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

                    italic_elements = row.get("Italicized in IEEE", "")

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
            overall.progress(base_pct + int(step * 0.10),
                             text=f"[{i}/{n}] Extracting references & verifying DOIs...")

            batch = {"filename": uf.name, "manuscript_year": int(manuscript_year),
                     "ai_done": False}
            try:
                process_single_ieee_pdf(uf, batch, client, int(manuscript_year))
            except Exception as exc:
                st.error(f"{uf.name} IEEE check failed: {exc}")
                batch["error"] = str(exc)
            st.session_state["ieee_batches"][key] = batch
            overall.progress(base_pct + int(step * 1.00),
                             text=f"[{i}/{n}] {uf.name} done.")

        overall.progress(100, text="All manuscripts processed.")
        overall.empty()
        st.success(f"Processed {n} manuscript{'s' if n != 1 else ''}.")

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
        st.text_area(
            "Reference slice",
            batch.get("reference_text", ""),
            height=300,
            key=f"dbg_ref_{selected_key}",
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

    metric_rows = [
        {"Metric": "Manuscript publication year", "Value": str(manuscript_year)},
        {"Metric": "Total references", "Value": total_refs_now},
        {"Metric": "Total in-text citation markers", "Value": stats.get("total", 0)},
        {"Metric": "Unique cited references", "Value": stats.get("unique", 0)},
        {"Metric": "Grouped citation clusters", "Value": stats.get("clusters", 0)},
        {"Metric": "Crowded clusters (needs collapse)", "Value": stats.get("crowded_clusters", 0)},
        {"Metric": "Orphan citations (no matching reference)", "Value": len(orphan)},
        {"Metric": "Uncited references", "Value": len(uncited)},
        {"Metric": "DOI checked via OpenAlex", "Value": doi_checked},
        {"Metric": "DOI suspicious (possible fabrication)",
         "Value": f"{doi_suspicious} ({doi_suspicious_pct:.1f}%)"},
        {"Metric": "Corrections withheld due to DOI mismatch", "Value": withheld_count},
        {"Metric": "References flagged MANUAL CHECK", "Value": manual_count},
        {"Metric": "References with placeholder fields", "Value": placeholder_count},
        {"Metric": f"% references within last 10 years ({window_start}-{window_end})",
         "Value": f"{recency.get('recent_percentage', 0):.1f}%"},
    ]
    metric_df = pd.DataFrame(metric_rows)

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