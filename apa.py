# apa.py
# OmniCite Auditor — APA 7th Edition module.
# Exposes render() to be called from app.py.
#
# Run via: streamlit run app.py   (app.py imports apa and calls apa.render())

import os
import io
import json
import re
import difflib
from collections import Counter
from datetime import datetime

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


# ============================================================
# HEADINGS
# ============================================================

REFERENCE_HEADINGS = [
    "references", "reference",
    "daftar pustaka", "daftar rujukan", "rujukan",
    "bibliografi", "bibliography", "referensi",
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
    "ucapan terima kasih", "ucapan terimakasih",
    "kontribusi penulis", "kontribusi author", "pernyataan kontribusi penulis",
    "kontribusi",
    "profil penulis", "biografi penulis",
    "konflik kepentingan", "pernyataan konflik kepentingan",
    "pendanaan", "pernyataan pendanaan",
    "ketersediaan data",
    "persetujuan etik", "persetujuan etika",
    "lampiran",
    "catatan", "catatan penulis",
}


def _normalize_heading(line: str) -> str:
    cleaned = line.strip()
    cleaned = re.sub(r"^[#>*\-\s]+", "", cleaned)
    cleaned = cleaned.strip("*_` ")
    cleaned = re.sub(r"\s+", " ", cleaned).lower()
    cleaned = cleaned.rstrip(":. ")
    return cleaned


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
    r"^~?\s*\d+\s*~\s*\d+\(\d+\),\s*\d+[-–]\d+\s*$",
    r"^\d+\s*~\s*\d+\(\d+\),\s*\d+[-–]\d+\s*$",
    r"^p?e?ISSN\s*\d+[-–]\d+.*$",
    r"^terakreditasi.*$",
    r"^\*?email koresponden.*$",
    r"^copyright ©.*$",
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
# SECTION SLICERS
# ============================================================

def slice_reference_section(text: str):
    lines = text.splitlines()
    ref_index = None
    post_index = None

    for i, line in enumerate(lines):
        norm = _normalize_heading(line)
        if not norm:
            continue

        if ref_index is None and norm in REFERENCE_HEADINGS:
            ref_index = i
            continue

        if ref_index is not None and norm in POST_REFERENCE_HEADINGS:
            post_index = i
            break

    if ref_index is None:
        return text, False, False

    end = post_index if post_index is not None else len(lines)
    sliced = "\n".join(lines[ref_index:end]).strip()

    return sliced, True, post_index is not None


def slice_body_section(text: str):
    lines = text.splitlines()
    ref_index = None

    for i, line in enumerate(lines):
        norm = _normalize_heading(line)
        if not norm:
            continue
        if norm in REFERENCE_HEADINGS:
            ref_index = i
            break

    if ref_index is None:
        return text, False

    body = "\n".join(lines[:ref_index]).strip()
    return body, True


# ============================================================
# LOCAL IN-TEXT CITATION EXTRACTION
# ============================================================

_YEAR_RE = re.compile(r"\b(?:19|20)\d{2}[a-z]?\b")


def _inside_parentheses(text, pos):
    depth = 0
    for i in range(pos):
        ch = text[i]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
    return depth > 0


def extract_parenthetical_citations(text):
    citations = []
    for content in re.findall(r"\(([^()]+)\)", text):
        if not _YEAR_RE.search(content):
            continue
        for part in content.split(";"):
            part = part.strip()
            ym = _YEAR_RE.search(part)
            if not ym:
                continue
            year = ym.group(0)[:4]
            author_part = part[:ym.start()].strip(" ,")
            et_al = bool(re.search(r"\bet\s+al\.", author_part, re.I))
            author_part_clean = re.sub(
                r"\bet\s+al\.", "", author_part, flags=re.I
            ).strip()
            authors = re.findall(
                r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\b",
                author_part_clean,
            )
            authors = [
                a for a in authors
                if a.lower() not in {"and", "according", "see", "cf"}
            ]
            if not authors:
                continue
            citations.append({
                "author": authors[0],
                "authors": authors,
                "year": year,
                "type": "parenthetical",
                "et_al": et_al,
                "raw": f"({part})",
                "parenthetical_content": content.strip(),
            })
    return citations


def extract_narrative_citations(text):
    citations = []
    occupied = []

    def _overlaps(start, end):
        return any(start >= s and end <= e for s, e in occupied)

    for m in re.finditer(
        r"\b("
        r"[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+"
        r"(?:\s*,\s*[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)+"
        r")"
        r"\s*,?\s*et\s+al\.\s*"
        r"\(((?:19|20)\d{2})[a-z]?\)",
        text,
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

    for m in re.finditer(
        r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\s+et\s+al\.\s*"
        r"\(((?:19|20)\d{2})[a-z]?\)",
        text,
    ):
        if _overlaps(m.start(), m.end()):
            continue
        citations.append({
            "author": m.group(1), "authors": [m.group(1)],
            "year": m.group(2), "type": "narrative",
            "et_al": True, "raw": m.group(0),
        })
        occupied.append((m.start(), m.end()))

    for m in re.finditer(
        r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\s+"
        r"\(((?:19|20)\d{2})[a-z]?\)",
        text,
    ):
        if _overlaps(m.start(), m.end()):
            continue
        citations.append({
            "author": m.group(1), "authors": [m.group(1)],
            "year": m.group(2), "type": "narrative",
            "et_al": False, "raw": m.group(0),
        })
        occupied.append((m.start(), m.end()))

    for m in re.finditer(
        r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)"
        r"\s+et\s+al\.\s*,\s*"
        r"((?:19|20)\d{2})[a-z]?\b",
        text,
    ):
        if _overlaps(m.start(), m.end()):
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

    for m in re.finditer(
        r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)"
        r"\s+(?:and|&)\s+"
        r"([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\s*,\s*"
        r"((?:19|20)\d{2})[a-z]?\b",
        text,
    ):
        if _overlaps(m.start(), m.end()):
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

    for m in re.finditer(
        r"(?<![,\.])\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\s*,\s*"
        r"((?:19|20)\d{2})[a-z]?\b",
        text,
    ):
        if _overlaps(m.start(), m.end()):
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


def citation_statistics(citations):
    narrative = sum(1 for c in citations if c["type"] == "narrative")
    parenthetical = sum(1 for c in citations if c["type"] == "parenthetical")
    total = narrative + parenthetical
    return {
        "total": total,
        "narrative": narrative,
        "parenthetical": parenthetical,
        "narrative_percentage": (narrative / total * 100 if total else 0),
        "parenthetical_percentage": (parenthetical / total * 100 if total else 0),
    }


# ============================================================
# SOURCE TYPES
# ============================================================

CANONICAL_SOURCE_TYPES = [
    "Journal Article", "Book", "Book Chapter", "Conference Proceeding",
    "Report", "Webpage / Online Document", "Other",
]

APA_REQUIRED_ELEMENTS = {
    "Journal Article": ["Author(s)", "Year", "Article title", "Journal title", "Volume", "Issue (when assigned)", "Page range or article number", "DOI when available"],
    "Book": ["Author(s) or editor(s)", "Year", "Book title", "Publisher", "DOI or URL when applicable"],
    "Book Chapter": ["Chapter author(s)", "Year", "Chapter title", "Editor(s)", "Book title", "Page range", "Publisher", "DOI or URL when applicable"],
    "Conference Proceeding": ["Author(s)", "Year", "Paper title", "Proceedings or conference title", "Publisher or organizer when applicable", "Page range when available", "DOI or URL when applicable"],
    "Report": ["Author or organization", "Year", "Report title", "Publisher or issuing organization when different from author", "Report number when available", "URL when applicable"],
    "Webpage / Online Document": ["Author or organization", "Date or n.d.", "Page/document title", "Website name when different from author", "URL"],
    "Other": ["Author or responsible organization", "Date when available", "Title", "Source information"],
}


_SOURCE_TYPE_ALIASES = {
    "journal": "Journal Article", "journal article": "Journal Article",
    "article": "Journal Article", "research article": "Journal Article",
    "book": "Book", "monograph": "Book",
    "book chapter": "Book Chapter", "chapter": "Book Chapter",
    "conference proceeding": "Conference Proceeding",
    "conference paper": "Conference Proceeding",
    "proceedings": "Conference Proceeding",
    "report": "Report", "technical report": "Report",
    "webpage": "Webpage / Online Document",
    "website": "Webpage / Online Document",
    "webpage / online document": "Webpage / Online Document",
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


# ============================================================
# REFERENCE PARSER + SOURCE TYPE DETECTION
# ============================================================

APA_DATE_RE = re.compile(r"\((?:(?:19|20)\d{2}[a-z]?|n\.d\.)\)", re.I)


def parse_reference(reference):
    result = {"raw": reference, "year": None, "first_author": None,
              "authors": [], "doi": None, "url": None}

    year_match = re.search(r"\(((?:19|20)\d{2})[a-z]?\)", reference)
    if not year_match:
        year_match = re.search(r"\b((?:19|20)\d{2})\b", reference)
    if year_match:
        result["year"] = year_match.group(1)

    author_block = (reference[:year_match.start()].strip()
                    if year_match else reference[:250])

    author_matches = re.findall(
        r"(?:^|,\s*)&?\s*([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)"
        r"\s*,\s*(?:[A-Z]\.\s*)+",
        author_block,
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
        url = url_match.group(0).rstrip(".,)")
        if re.match(r"^https?://(?:dx\.)?doi\.org/10\.", url, re.I):
            m = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", url)
            if m:
                result["doi"] = f"https://doi.org/{m.group(0).rstrip('.,;)')}"
                result["url"] = result["doi"]
        else:
            result["url"] = url
    else:
        dm = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", reference, re.I)
        if dm:
            result["doi"] = f"https://doi.org/{dm.group(0).rstrip('.,;)')}"
            result["url"] = result["doi"]

    return result


def extract_reference_year(reference):
    m = re.search(r"\(((?:19|20)\d{2})(?:[a-z])?\)", reference, re.I)
    return int(m.group(1)) if m else None


def detect_apa_source_type(reference):
    low = reference.lower()
    if re.search(
        r"\.\s*[^.]+?,\s*\d+\s*(?:\([^)]+\))?\s*,\s*\d+(?:\s*[–-]\s*\d+)?",
        reference,
    ):
        return "Journal Article"
    if re.search(r"\(eds?\.\)", reference, re.I):
        return "Book Chapter"
    if re.search(r"\b(proceedings|conference|symposium)\b", low):
        return "Conference Proceeding"
    if re.search(r"\b(report|working paper)\b", low):
        return "Report"
    if re.search(r"https?://", reference) and "doi.org" not in low:
        return "Webpage / Online Document"
    return "Other"


# ============================================================
# OPENALEX DOI VERIFICATION
# ============================================================

def _normalize_doi_for_lookup(doi):
    if not doi:
        return ""
    s = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", doi.strip(), flags=re.I)
    s = re.sub(r"^doi\s*:\s*", "", s, flags=re.I)
    s = s.rstrip(".,;)")
    m = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+", s)
    return m.group(0) if m else ""


def _fetch_openalex_metadata(doi, api_key):
    if not doi or not api_key:
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
    authors = [
        (a.get("author") or {}).get("display_name")
        for a in (data.get("authorships") or [])
        if (a.get("author") or {}).get("display_name")
    ]

    primary_location = data.get("primary_location") or {}
    source_block = primary_location.get("source") or {}
    source_display_name = source_block.get("display_name")

    biblio = data.get("biblio") or {}
    biblio_volume = biblio.get("volume")
    biblio_issue = biblio.get("issue")

    return {
        "title": title,
        "authors": authors,
        "year": year,
        "source_display_name": source_display_name,
        "biblio_volume": biblio_volume,
        "biblio_issue": biblio_issue,
    }


def _normalize_for_compare(s):
    s = re.sub(r"\s+", " ", s or "").strip().lower()
    return re.sub(r"[^a-z0-9 ]", "", s)


def _title_similarity(a, b):
    a_n, b_n = _normalize_for_compare(a), _normalize_for_compare(b)
    if not a_n or not b_n:
        return 0.0
    return difflib.SequenceMatcher(None, a_n, b_n).ratio()


def _surname_from_full_name(name):
    if not name:
        return ""
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
    return m.group(1).strip() if m else tail.split(".")[0].strip()


def verify_reference_against_openalex(reference, parsed, api_key):
    result = {
        "checked": False, "doi": parsed.get("doi"),
        "title_similarity": None, "author_overlap": None,
        "suspicious": False, "reasons": [],
        "openalex_title": None, "openalex_authors": [],
        "openalex_year": None,
        "openalex_source_display_name": None,
        "openalex_biblio_volume": None,
        "openalex_biblio_issue": None,
        "resolved": False,
    }
    doi = parsed.get("doi")
    if not doi:
        return result

    meta = _fetch_openalex_metadata(doi, api_key)
    if meta is None:
        return result

    if meta.get("_not_found"):
        result["checked"] = True
        result["suspicious"] = True
        result["reasons"].append("DOI does not resolve in OpenAlex.")
        return result

    result["checked"] = True
    result["resolved"] = True
    result["openalex_title"] = meta.get("title")
    result["openalex_authors"] = meta.get("authors", [])
    result["openalex_year"] = meta.get("year")
    result["openalex_source_display_name"] = meta.get("source_display_name")
    result["openalex_biblio_volume"] = meta.get("biblio_volume")
    result["openalex_biblio_issue"] = meta.get("biblio_issue")

    ref_title = _extract_apa_title(reference)
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
    oa_surnames = {_surname_from_full_name(n) for n in meta.get("authors", [])}
    oa_surnames.discard("")
    if oa_surnames:
        overlap = len(ref_surnames & oa_surnames) / max(1, len(oa_surnames))
        result["author_overlap"] = overlap
        if overlap < 0.50:
            result["suspicious"] = True
            result["reasons"].append(
                f"DOI author list does not match manuscript "
                f"({overlap:.0%} surname overlap)."
            )

    return result


# ============================================================
# ITALIC SPANS
# ============================================================

def compute_italic_spans_from_openalex(
    openalex_meta, original_reference, corrected_reference=None
):
    spans = []

    if not openalex_meta:
        openalex_meta = {}

    target = corrected_reference or original_reference or ""

    journal = openalex_meta.get("source_display_name")
    if journal:
        spans.append(journal)

    volume = openalex_meta.get("biblio_volume")
    if volume:
        m = re.search(
            rf"(?<!\d){re.escape(str(volume))}(?!\d)", target
        )
        if m:
            spans.append(m.group(0))
    else:
        m = re.search(
            r"\.\s*([^.]+?),\s*(\d+)\s*(?:\(\d+\))?\s*,\s*"
            r"\d+(?:\s*[–-]\s*\d+)?",
            target,
        )
        if m:
            journal_fallback = m.group(1).strip()
            volume_fallback = m.group(2).strip()
            if journal_fallback and journal_fallback not in spans:
                spans.append(journal_fallback)
            if volume_fallback and volume_fallback not in spans:
                spans.append(volume_fallback)

    issue = openalex_meta.get("biblio_issue")
    if issue:
        issue_token = f"({issue})"
        spans = [s for s in spans if s not in {issue, issue_token}]

    seen = set()
    unique = []
    for s in spans:
        if s and s not in seen:
            seen.add(s)
            unique.append(s)
    return unique


# ============================================================
# OPENAI HELPERS
# ============================================================

INPUT_PRICE_PER_MILLION = 0.15
OUTPUT_PRICE_PER_MILLION = 0.60


def _build_usage(response):
    i = response.usage.prompt_tokens
    o = response.usage.completion_tokens
    t = response.usage.total_tokens
    ic = (i / 1_000_000) * INPUT_PRICE_PER_MILLION
    oc = (o / 1_000_000) * OUTPUT_PRICE_PER_MILLION
    return {
        "model": "gpt-4o-mini",
        "input_tokens": i, "output_tokens": o, "total_tokens": t,
        "input_cost": ic, "output_cost": oc, "total_cost": ic + oc,
    }


def _safe_json_loads(text):
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
        text = re.sub(r"\s*```$", "", text)
    return json.loads(text)


# ============================================================
# OPENAI — REFERENCE EXTRACTION
# ============================================================

def extract_references(reference_text, client):
    system_prompt = """
You are an academic manuscript reference extraction system.

Your job is extraction, NOT correction.

Return compact valid JSON only.

Identify every bibliography/reference-list entry from the supplied text.

RULES:
- Preserve each reference exactly as much as possible.
- Do not correct APA formatting.
- Do not invent references.
- Do not merge or split references.
- Reference may continue across multiple lines.
- Two references may appear on the same line — split them.
- Use author + year patterns to identify boundaries.
- DOI/URL belongs to the preceding reference.
- Preserve wording as in the source.

Return exactly:
{
  "references": [
    { "reference": "complete extracted reference" }
  ]
}
"""
    user_prompt = (
        "Extract all bibliography references from the following "
        "reference section.\n\nREFERENCE SECTION:\n\n" + reference_text
    )

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        response_format={"type": "json_object"},
        temperature=0,
        max_tokens=16384,
    )

    finish_reason = response.choices[0].finish_reason
    content = response.choices[0].message.content or ""
    if not content:
        raise ValueError("OpenAI returned an empty response.")

    data = json.loads(content)
    usage = _build_usage(response)
    return data, finish_reason, len(content), usage


# ============================================================
# OPENAI — APA REFERENCE REVIEW
# ============================================================

def review_apa_references_with_ai(items, client):
    if not items:
        return []

    prompt = f"""
You are reconstructing APA 7th edition REFERENCE-LIST entries.

Each input item has:
  - "number": the reference's index
  - "reference": the ORIGINAL reference string extracted from the PDF
  - "openalex": optional metadata for the SAME reference, resolved from
    OpenAlex by its DOI. If present, it is AUTHORITATIVE for authors,
    title, and year. It may also include:
      - "source_display_name" (journal name)
      - "biblio_volume"        (volume number)
      - "biblio_issue"         (issue number)

RULES:

1. If "openalex" is present with a title:
   - RECONSTRUCT the reference from the OpenAlex metadata.
   - Authors: use OpenAlex authors, converted to APA 7 form
     (Surname, A. A., & Surname, B. B.).
   - Year: use OpenAlex year.
   - Title: use OpenAlex title. If the OpenAlex title is in ALL CAPS,
     convert it to sentence case (only the first word and proper nouns
     capitalized). Do NOT use title case.
   - Journal name: use OpenAlex "source_display_name" if present.
   - Volume: use OpenAlex "biblio_volume" if present.
   - Issue: use OpenAlex "biblio_issue" if present. If OpenAlex has no
     issue, preserve the issue number from the ORIGINAL reference if it
     is present there. The issue always goes inside parentheses directly
     after the volume:
        Journal Name, 7(5), 2209-2216.
     If no issue is available from either source, omit the parentheses.
   - Pages: preserve from the ORIGINAL reference.
   - DOI: preserve from the ORIGINAL reference as https://doi.org/...

2. If "openalex" is absent, null, or has no title:
   - Correct the ORIGINAL reference ONLY. Do not invent facts.
   - If the article/book title in the original is in ALL CAPS, convert
     it to sentence case.

3. Non-English titles:
   - You MUST add an English translation in parentheses for any title
     that is not in English.
   - Format:  Original-language title (English translation)
   - The translation must be a faithful, literal English rendering.
   - Detect the language of the title. If you are not confident it is
     English, treat it as non-English and translate.
   - Apply this to article titles, book titles, chapter titles, and
     report titles.
   - Do NOT translate journal names, publisher names, or conference
     names; keep them as-is.

4. Do NOT invent missing facts.
5. revised_reference must contain ONLY the corrected APA reference.

6. "italic_elements" must be a comma-separated list of EXACT substrings
   that appear verbatim inside revised_reference and must be italicized
   under APA 7 rules. Do NOT return descriptive words.
   Do NOT include the issue number inside parentheses.

APA 7 italic rules (for reference):
  - Journal Article: journal name + volume number are italic.
                     Issue number in parentheses is NOT italic.
  - Book: book title is italic.
  - Book Chapter: book title is italic.
  - Report: report title is italic.
  - Webpage: webpage title is italic.

SOURCE TYPE — exactly one of:
  "Journal Article", "Book", "Book Chapter", "Conference Proceeding",
  "Report", "Webpage / Online Document", "Other"

Return JSON only:
{{"results": [
  {{"number": int, "status": "OK"|"REVISED"|"MANUAL_CHECK",
    "revised_reference": str, "source_type": str,
    "italic_elements": str, "year": int|null,
    "missing_required_elements": [str], "explanation": str}}
]}}

INPUT:
{json.dumps(items, ensure_ascii=False)}
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system",
                 "content": "You are a precise APA 7 reference editor. JSON only."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        results = _safe_json_loads(
            response.choices[0].message.content
        ).get("results", [])
        for r in results:
            r["source_type"] = normalize_source_type(r.get("source_type"))
        return results
    except Exception as exc:
        return [{"number": 0, "status": "MANUAL_CHECK",
                 "revised_reference": "", "source_type": "Other",
                 "italic_elements": "", "year": None,
                 "missing_required_elements": [],
                 "explanation": f"OpenAI API error: {exc}"}]


# ============================================================
# OPENAI — APA IN-TEXT CITATION REVIEW
# ============================================================

def enforce_apa_citation_rules(original, revised, citation_type):
    """Deterministically enforce APA 7 conjunction rules after AI review."""
    text = (revised or original or "").strip()
    if citation_type == "parenthetical":
        text = re.sub(r"\s+(?:and|dan)\s+", " & ", text, flags=re.I)
    elif citation_type == "narrative":
        text = re.sub(r"\s+(?:&|dan)\s+", " and ", text, flags=re.I)
    return text


def review_apa_citations_with_ai(citations, reference_rows, client):
    if not citations:
        return []

    payload = [
        {"number": i, "citation": c.get("raw", ""), "type": c.get("type", "")}
        for i, c in enumerate(citations, start=1)
    ]

    reference_context = []
    for row in reference_rows:
        ref_text = row.get("Corrected Version", "")
        if not ref_text or ref_text.startswith("— WITHHELD"):
            ref_text = row.get("Original Reference", "")
        parsed = parse_reference(ref_text)
        reference_context.append({
            "year": parsed.get("year"),
            "authors": parsed.get("authors", []),
            "reference": ref_text,
        })

    prompt = f"""
You are checking APA 7th edition IN-TEXT citations.

Each item is ALREADY a separate citation. Correct each IN ISOLATION.
Do not merge sources. Do not invent authors or years.
Use the matching REFERENCE LIST CONTEXT to identify author surnames whenever possible.

APA 7 AUTHOR RULES:
- Use AUTHOR SURNAMES from the matching reference-list entry.
- Never preserve a full personal name in an in-text citation when the matching reference identifies the surname.
- 1 author: parenthetical (Surname, 2024); narrative Surname (2024).
- 2 authors: parenthetical MUST use "&"; narrative MUST use "and".
- 3+ authors: FirstSurname et al.
- Do not guess surnames merely from word position when the reference list provides the surname.

Examples:
(Ana Ittihada, 2026) + reference "Ittihada, A. (2026)" -> (Ittihada, 2026)
(Darius Ru'ung, 2021) + reference "Ru'ung, D. (2021)" -> (Ru'ung, 2021)
(Aji Sofanudin dan Wahab, 2020) -> (Sofanudin & Wahab, 2020)

Return JSON only:
{{"results": [
  {{"number": int, "status": "OK"|"REVISED"|"MANUAL_CHECK",
    "revised_citation": str, "explanation": str}}
]}}

REFERENCE LIST CONTEXT:
{json.dumps(reference_context, ensure_ascii=False)}

INPUT CITATIONS:
{json.dumps(payload, ensure_ascii=False)}
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a precise APA 7 citation editor. JSON only."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
        )
        return _safe_json_loads(response.choices[0].message.content).get("results", [])
    except Exception as exc:
        return [{"number": 0, "status": "MANUAL_CHECK", "revised_citation": "",
                 "explanation": f"OpenAI API error: {exc}"}]


# ============================================================
# DOCX REPORT EXPORT
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
    _set_run_font(r, size_pt=size_pt, bold=bold, italic=italic,
                  color=RED if red else None)
    return r


def _add_red_italic_run(paragraph, text, size_pt=11):
    _add_run(paragraph, text, size_pt=size_pt, bold=True, italic=True, red=True)


def _docx_set_default_font(document, font_name="Times New Roman", size_pt=11):
    style = document.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(size_pt)


def _italic_pattern(tokens):
    parts = []
    for t in tokens:
        t = (t or "").strip()
        if not t:
            continue
        if re.fullmatch(r"\d+", t):
            parts.append(rf"(?<!\d){re.escape(t)}(?!\d)")
        else:
            parts.append(re.escape(t))
    if not parts:
        return None
    return re.compile("|".join(parts), re.I)


def _emit_with_italic_tokens(paragraph, text, tokens, size_pt=11):
    pattern = _italic_pattern(tokens)
    if pattern is None:
        _add_run(paragraph, text, size_pt=size_pt)
        return

    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], size_pt=size_pt)
        _add_run(paragraph, m.group(0), size_pt=size_pt, italic=True)
        pos = m.end()
    if pos < len(text):
        _add_run(paragraph, text[pos:], size_pt=size_pt)


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


def _compute_citation_pages(full_text, citations, page_break_tag="<<<PAGE_BREAK"):
    if not full_text or page_break_tag not in full_text:
        return [None] * len(citations)

    pages = []
    current_page = 1
    current_lines = []
    for line in full_text.splitlines():
        stripped = line.strip()
        m = re.fullmatch(r"<<<PAGE_BREAK:(\d+)>>>", stripped)
        if m:
            pages.append((current_page, "\n".join(current_lines)))
            current_page = int(m.group(1))
            current_lines = []
        else:
            current_lines.append(line)
    pages.append((current_page, "\n".join(current_lines)))

    def norm(s):
        return re.sub(r"\s+", " ", s or "").strip().lower()

    page_lookup = [(num, norm(text)) for num, text in pages]

    result = []
    for c in citations:
        raw = norm(c.get("raw", ""))
        if not raw:
            result.append(None)
            continue
        found_page = None
        for num, page_text in page_lookup:
            if raw in page_text:
                found_page = num
                break
        result.append(found_page)
    return result


def build_apa_report_docx(result):
    """
    `result` is the batch dict. `manuscript_year` inside it controls
    the trailing-10-year window (inclusive of that year).
    """
    doc = Document()
    _docx_set_default_font(doc)

    title = doc.add_heading(level=0)
    tr = title.add_run(result.get("filename", "manuscript.pdf"))
    _set_run_font(tr, size_pt=18, bold=True)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    sr = subtitle.add_run(
        "APA 7th Edition — Citation & Reference Diagnostic Report"
    )
    _set_run_font(sr, size_pt=11, italic=True)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = doc.add_paragraph()
    mr = meta.add_run(
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    )
    _set_run_font(mr, size_pt=9, italic=True)
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    citation_rows = result.get("citation_rows", [])
    reference_rows = result.get("reference_rows", [])
    stats = result.get("citation_stats", {})
    manuscript_year = result.get("manuscript_year", datetime.now().year)
    full_text = result.get("full_text", "")

    total_refs = len(reference_rows)
    total_cits = stats.get("total", 0)

    ref_keys = {}
    for row in reference_rows:
        p = parse_reference(row.get("Original Reference", ""))
        if p["first_author"] and p["year"]:
            ref_keys[(p["first_author"].lower(), p["year"])] = row["No."]

    cit_keys = set()
    for c in result.get("citations", []):
        cit_keys.add((c["author"].lower(), c["year"]))

    cited_ref_nos = set()
    for row in reference_rows:
        p = parse_reference(row.get("Original Reference", ""))
        key = (p["first_author"].lower() if p["first_author"] else "", p["year"])
        if key in cit_keys:
            cited_ref_nos.add(row["No."])

    not_in_refs = []
    for c in result.get("citations", []):
        key = (c["author"].lower(), c["year"])
        if key not in ref_keys:
            not_in_refs.append(c)

    citation_page_map = _compute_citation_pages(
        full_text, result.get("citations", [])
    )

    h1 = doc.add_heading(level=1)
    hr = h1.add_run("1. Summary")
    _set_run_font(hr, size_pt=14, bold=True)

    doi_checked = sum(
        1 for r in reference_rows if r.get("DOI Checked") == "YES"
    )
    doi_suspicious = sum(
        1 for r in reference_rows if r.get("DOI Suspicious") != "—"
    )
    doi_suspicious_pct = (
        doi_suspicious / total_refs * 100 if total_refs else 0
    )
    doi_reconstructed = sum(
        1 for r in reference_rows if r.get("Source") == "OpenAlex"
    )

    # --- Revised: last-10-year window now anchored to manuscript_year ---
    window_start = manuscript_year - 9
    window_end = manuscript_year
    recent = sum(
        1 for row in reference_rows
        if isinstance(row.get("Year"), int)
        and window_start <= row["Year"] <= window_end
    )
    recent_pct = (recent / total_refs * 100) if total_refs else 0

    refs_missing_from_cits = total_refs - len(cited_ref_nos)
    cits_missing_from_refs = len(not_in_refs)

    summary_rows = [
        ("Manuscript publication year", str(manuscript_year), False),
        ("Total references", str(total_refs), False),
        ("Total in-text citations", str(total_cits), False),
        ("  • Narrative", str(stats.get("narrative", 0)), False),
        ("  • Parenthetical", str(stats.get("parenthetical", 0)), False),
        ("Citations NOT in reference list",
         str(cits_missing_from_refs), cits_missing_from_refs > 0),
        ("References NOT cited in text",
         str(refs_missing_from_cits), refs_missing_from_cits > 0),
        ("DOI checked via OpenAlex", str(doi_checked), False),
        ("References reconstructed from OpenAlex",
         str(doi_reconstructed), False),
        ("DOI suspicious (possible fabricated references)",
         f"{doi_suspicious} ({doi_suspicious_pct:.1f}%)",
         doi_suspicious > 0),
        (f"% references within last 10 years "
         f"({window_start}-{window_end})",
         f"{recent_pct:.1f}%", False),
    ]

    summary_table = doc.add_table(rows=1, cols=2)
    summary_table.style = "Light Grid Accent 1"
    hdr = summary_table.rows[0].cells
    for cell, text in zip(hdr, ["Metric", "Value"]):
        cell.text = ""
        _set_run_font(
            cell.paragraphs[0].add_run(text), size_pt=10, bold=True
        )

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

    h11 = doc.add_heading(level=2)
    hr11 = h11.add_run("1.1 Source Type Distribution")
    _set_run_font(hr11, size_pt=12, bold=True)

    src_counts = Counter(
        row.get("Source Type", "Other") for row in reference_rows
    )
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Light Grid Accent 1"
    for cell, text in zip(
        tbl.rows[0].cells, ["Source Type", "Count", "Percentage"]
    ):
        cell.text = ""
        _set_run_font(
            cell.paragraphs[0].add_run(text), size_pt=10, bold=True
        )

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
        _set_run_font(
            cell.paragraphs[0].add_run(val), size_pt=10, bold=True
        )

    doc.add_page_break()

    h2 = doc.add_heading(level=1)
    hr2 = h2.add_run("2. In-text Citation Correction")
    _set_run_font(hr2, size_pt=14, bold=True)

    caption = doc.add_paragraph()
    cr = caption.add_run(
        "Format: page number — original citation — corrected citation"
    )
    _set_run_font(cr, size_pt=9, italic=True)

    if not citation_rows:
        p = doc.add_paragraph()
        _set_run_font(
            p.add_run("No in-text citations were detected."), italic=True
        )
    else:
        for i, row in enumerate(citation_rows):
            c_dict = result.get("citations", [])[i] if i < len(
                result.get("citations", [])
            ) else None
            missing_from_refs = False
            if c_dict:
                key = (c_dict["author"].lower(), c_dict["year"])
                missing_from_refs = key not in ref_keys

            page_no = citation_page_map[i] if i < len(citation_page_map) else None

            head = doc.add_paragraph()
            head.paragraph_format.space_before = Pt(4)
            head.paragraph_format.space_after = Pt(2)
            hp = f"{row.get('No.', i+1)}. [{row.get('Type', '')}]"
            if page_no:
                hp += f" — Page {page_no}"
            hrun = head.add_run(hp)
            _set_run_font(hrun, size_pt=11, bold=True)

            if missing_from_refs:
                _add_red_italic_run(head, "  [NOT IN REFERENCES]")

            p_orig = doc.add_paragraph()
            p_orig.paragraph_format.space_after = Pt(2)
            p_orig.paragraph_format.left_indent = Inches(0.25)
            _add_run(p_orig, "Original:  ", bold=True, size_pt=11)
            if missing_from_refs:
                _add_run(
                    p_orig, row.get("Original Citation", ""),
                    size_pt=11, red=True,
                )
            else:
                _add_run(
                    p_orig, row.get("Original Citation", ""), size_pt=11
                )

            p_corr = doc.add_paragraph()
            p_corr.paragraph_format.space_after = Pt(2)
            p_corr.paragraph_format.left_indent = Inches(0.25)
            _add_run(p_corr, "Corrected: ", bold=True, size_pt=11)
            _add_run(
                p_corr, row.get("Revised Citation", ""), size_pt=11
            )

            notes = row.get("Notes", "")
            if notes:
                np = doc.add_paragraph()
                np.paragraph_format.left_indent = Inches(0.25)
                np.paragraph_format.space_after = Pt(2)
                _add_run(np, f"Note: {notes}", size_pt=10, italic=True)

            _add_divider(doc)

    doc.add_page_break()

    h3 = doc.add_heading(level=1)
    hr3 = h3.add_run("3. Reference List (APA 7th Edition)")
    _set_run_font(hr3, size_pt=14, bold=True)

    caption = doc.add_paragraph()
    cr2 = caption.add_run(
        "Format: original reference — corrected reference"
    )
    _set_run_font(cr2, size_pt=9, italic=True)

    if not reference_rows:
        p = doc.add_paragraph()
        _set_run_font(
            p.add_run("No references were detected."), italic=True
        )
    else:
        for row in reference_rows:
            not_cited = row["No."] not in cited_ref_nos

            head = doc.add_paragraph()
            head.paragraph_format.space_before = Pt(4)
            head.paragraph_format.space_after = Pt(2)
            hrun = head.add_run(f"{row.get('No.', '')}.")
            _set_run_font(hrun, size_pt=11, bold=True)

            if not_cited:
                _add_red_italic_run(head, "  [NOT CITED IN TEXT]")

            p_orig = doc.add_paragraph()
            p_orig.paragraph_format.space_after = Pt(2)
            p_orig.paragraph_format.left_indent = Inches(0.25)
            p_orig.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            _add_run(p_orig, "Original:  ", bold=True, size_pt=11)

            original = row.get("Original Reference", "")
            if row.get("DOI Suspicious") != "—" or not_cited:
                _add_run(p_orig, original, size_pt=11, red=True)
            else:
                _add_run(p_orig, original, size_pt=11)

            if row.get("Status") == "WITHHELD":
                withheld = doc.add_paragraph()
                withheld.paragraph_format.left_indent = Inches(0.25)
                withheld.paragraph_format.space_after = Pt(6)
                _add_run(
                    withheld,
                    row.get("Explanation") or "Automated correction withheld.",
                    size_pt=10, italic=True, bold=True, red=True,
                )
            elif row.get("DOI Suspicious") != "—":
                withheld = doc.add_paragraph()
                withheld.paragraph_format.left_indent = Inches(0.25)
                withheld.paragraph_format.space_after = Pt(6)
                withheld.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                _add_run(
                    withheld,
                    "Corrected version withheld — the DOI in this reference "
                    "does not match the claimed title/authors. Manual "
                    "verification required.",
                    size_pt=10, italic=True, bold=True, red=True,
                )
            else:
                p_corr = doc.add_paragraph()
                p_corr.paragraph_format.space_after = Pt(2)
                p_corr.paragraph_format.left_indent = Inches(0.25)
                p_corr.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                _add_run(p_corr, "Corrected: ", bold=True, size_pt=11)

                corrected = row.get("Corrected Version", "")
                italic_elements = row.get("Italicized in APA", "")
                tokens = [
                    t.strip()
                    for t in (italic_elements or "").split(",")
                    if t.strip()
                ]
                _emit_with_italic_tokens(
                    p_corr, corrected, tokens, size_pt=11
                )

                if row.get("Source") == "OpenAlex":
                    src_p = doc.add_paragraph()
                    src_p.paragraph_format.left_indent = Inches(0.25)
                    _add_run(
                        src_p,
                        "Reconstructed from OpenAlex DOI metadata.",
                        size_pt=10, italic=True,
                    )

            if row.get("DOI Suspicious") != "—":
                warn = doc.add_paragraph()
                warn.paragraph_format.left_indent = Inches(0.25)
                warn.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                _add_run(
                    warn,
                    f"⚠ Possible fabricated reference: "
                    f"{row.get('DOI Issues', '')}",
                    size_pt=10, italic=True, bold=True, red=True,
                )
                oa_title = row.get("OpenAlex Title", "")
                if oa_title:
                    _add_run(
                        warn,
                        f'\n   OpenAlex says: "{oa_title}"',
                        size_pt=10, italic=True,
                    )

            # Source-type-specific APA requirements
            req_p = doc.add_paragraph()
            req_p.paragraph_format.left_indent = Inches(0.25)
            _add_run(req_p, f"Source Type: {row.get('Source Type', 'Other')}", size_pt=10, bold=True)

            req2 = doc.add_paragraph()
            req2.paragraph_format.left_indent = Inches(0.25)
            _add_run(req2, "Mandatory Elements: ", size_pt=10, bold=True)
            _add_run(req2, "; ".join(row.get("Required Elements", [])), size_pt=10)

            missing = row.get("Missing Required Elements", [])
            miss_p = doc.add_paragraph()
            miss_p.paragraph_format.left_indent = Inches(0.25)
            if missing:
                _add_run(miss_p, "Missing Mandatory Elements: ", size_pt=10, bold=True, red=True)
                _add_run(miss_p, "; ".join(missing), size_pt=10, red=True)
            else:
                _add_run(miss_p, "Missing Mandatory Elements: None detected", size_pt=10)

            if row.get("Status") == "WITHHELD" and row.get("Explanation"):
                note_p = doc.add_paragraph()
                note_p.paragraph_format.left_indent = Inches(0.25)
                _add_run(note_p, row.get("Explanation"), size_pt=10, italic=True, red=True)

            status = row.get("Status", "")
            if status and status != "OK":
                sp = doc.add_paragraph()
                sp.paragraph_format.left_indent = Inches(0.25)
                _add_run(sp, f"Status: {status}", size_pt=10, italic=True)

            _add_divider(doc)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


# ============================================================
# PIPELINE — full APA processing for one PDF
# ============================================================

def process_single_pdf(uf, batch, client, openalex_api_key, manuscript_year):
    """
    Run the full APA pipeline for ONE PDF's batch dict.
    Mutates `batch` in place with the AI results.
    `manuscript_year` is the user-supplied year, used for the
    10-year recency window.
    """
    # ---- Stage A: extract references ----
    (
        ref_result,
        ref_finish,
        ref_output_len,
        ref_usage,
    ) = extract_references(batch["reference_text"], client)

    reference_list = [
        (item.get("reference", "") if isinstance(item, dict) else str(item))
        for item in ref_result.get("references", [])
    ]

    parsed_refs = [parse_reference(r) for r in reference_list]
    local_source_types = [detect_apa_source_type(r) for r in reference_list]

    # ---- Stage B: verify DOIs via OpenAlex ----
    verification_rows = []
    for ref, parsed in zip(reference_list, parsed_refs):
        v = verify_reference_against_openalex(ref, parsed, openalex_api_key)
        verification_rows.append(v)

    # ---- Stage C: decide which references are eligible for AI review ----
    # Cost-control policy:
    #   * Journal + no DOI: withhold; do NOT send to OpenAI.
    #   * Journal + suspicious/unresolved DOI: manual check; do NOT send to OpenAI.
    #   * Journal + verified DOI: eligible for AI reconstruction/correction.
    #   * Non-journal sources: manual check for now; do NOT send to OpenAI.
    review_payload = []
    preclassified_results = {}

    for i, (ref, parsed, source_type, v) in enumerate(
        zip(reference_list, parsed_refs, local_source_types, verification_rows), start=1
    ):
        if source_type == "Journal Article":
            if not parsed.get("doi"):
                preclassified_results[i] = {
                    "number": i, "status": "WITHHELD", "revised_reference": "",
                    "source_type": source_type, "italic_elements": "",
                    "year": parsed.get("year"), "missing_required_elements": [],
                    "explanation": "DOI not provided — automated verification/correction withheld.",
                }
                continue

            if v.get("suspicious") or not v.get("resolved"):
                reason = " | ".join(v.get("reasons", [])) or "DOI metadata could not be independently verified."
                preclassified_results[i] = {
                    "number": i, "status": "MANUAL_CHECK", "revised_reference": "",
                    "source_type": source_type, "italic_elements": "",
                    "year": parsed.get("year"), "missing_required_elements": [],
                    "explanation": reason,
                }
                continue

            openalex_block = {
                "title": v["openalex_title"],
                "authors": v.get("openalex_authors") or [],
                "year": v.get("openalex_year"),
                "source_display_name": v.get("openalex_source_display_name"),
                "biblio_volume": v.get("openalex_biblio_volume"),
                "biblio_issue": v.get("openalex_biblio_issue"),
            }
            review_payload.append({"number": i, "reference": ref, "openalex": openalex_block})
            continue

        preclassified_results[i] = {
            "number": i, "status": "MANUAL_CHECK", "revised_reference": "",
            "source_type": source_type, "italic_elements": "",
            "year": parsed.get("year"), "missing_required_elements": [],
            "explanation": "No independently verified metadata workflow is configured for this source type; manual verification required.",
        }

    # ---- Stage D: AI review only for eligible, independently verified references ----
    ref_ai = review_apa_references_with_ai(review_payload, client) if review_payload else []

    by_no = {
        int(x.get("number", -1)): x
        for x in (ref_ai or [])
        if str(x.get("number", "")).isdigit()
    }
    by_no.update(preclassified_results)

    # ---- Stage E: build ref rows ----
    ref_rows = []
    for i, ref in enumerate(reference_list, start=1):
        ai = by_no.get(i, {})
        ai_revised = (ai.get("revised_reference") or "").strip()
        corrected = ai_revised or ref
        stype = ai.get("source_type") or local_source_types[i - 1]
        v = verification_rows[i - 1]

        reconstructed = (
            v.get("resolved")
            and v.get("openalex_title")
            and not v["suspicious"]
            and corrected.strip() != ref.strip()
        )

        openalex_meta_for_italic = {
            "source_display_name": v.get("openalex_source_display_name"),
            "biblio_volume": v.get("openalex_biblio_volume"),
            "biblio_issue": v.get("openalex_biblio_issue"),
        } if v.get("resolved") else None

        local_spans = compute_italic_spans_from_openalex(
            openalex_meta_for_italic,
            ref,
            corrected_reference=corrected,
        )
        ai_spans = [
            t.strip() for t in (ai.get("italic_elements") or "").split(",")
            if t.strip()
        ]

        combined = []
        for s in local_spans + ai_spans:
            if s and s not in combined:
                combined.append(s)

        ref_rows.append({
            "No.": i,
            "Source Type": stype,
            "Year": ai.get("year") or extract_reference_year(corrected),
            "Original Reference": ref,
            "Corrected Version": (
                "— WITHHELD —"
                if ai.get("status") == "WITHHELD"
                else ("— WITHHELD (DOI mismatch) —" if v["suspicious"] else corrected)
            ),
            "Source": ("OpenAlex" if reconstructed else ("AI" if ai_revised else "Manual")),
            "DOI Checked": "YES" if v["checked"] else "NO",
            "DOI Suspicious": "⚠️ YES" if v["suspicious"] else "—",
            "OpenAlex Title": (v.get("openalex_title") or "")[:60],
            "DOI Issues": " | ".join(v.get("reasons", [])),
            "Status": (ai.get("status") or "MANUAL CHECK").upper(),
            "Explanation": ai.get("explanation", ""),
            "Italicized in APA": ", ".join(combined),
            "Missing Required Elements": ai.get("missing_required_elements", []),
            "Required Elements": APA_REQUIRED_ELEMENTS.get(stype, APA_REQUIRED_ELEMENTS["Other"]),
        })

    # ---- Stage F: AI citation review ----
    cit_ai = review_apa_citations_with_ai(batch["citations"], ref_rows, client)

    cit_by_no = {
        int(x.get("number", -1)): x
        for x in (cit_ai or [])
        if str(x.get("number", "")).isdigit()
    }

    cit_rows = []
    for i, c in enumerate(batch["citations"], start=1):
        ai = cit_by_no.get(i, {})
        revised = (ai.get("revised_citation") or "").strip() or c["raw"]
        revised = enforce_apa_citation_rules(c["raw"], revised, c["type"])
        cit_rows.append({
            "No.": i,
            "Type": c["type"].title(),
            "Original Citation": c["raw"],
            "Revised Citation": revised,
            "Status": (ai.get("status") or "MATCH").upper(),
            "Notes": ai.get("explanation", ""),
        })

    # ---- Stage G: attach to batch ----
    batch.update({
        "reference_list": reference_list,
        "reference_rows": ref_rows,
        "citation_rows": cit_rows,
        "verification_rows": verification_rows,
        "manuscript_year": int(manuscript_year),
        "ref_usage": ref_usage,
        "ref_finish": ref_finish,
        "ref_output_len": ref_output_len,
        "ai_done": True,
    })


# ============================================================
# PDF EXTRACTION (MarkItDown)
# ============================================================

def extract_pdf_text(uploaded_file):
    md = MarkItDown(enable_plugins=False)
    pdf_bytes = uploaded_file.getvalue()
    stream = io.BytesIO(pdf_bytes)

    result = md.convert_stream(stream, file_extension=".pdf")
    full_text = result.text_content or ""

    cleaned_text = strip_running_headers_footers(full_text)

    reference_text, ref_found, post_found = slice_reference_section(cleaned_text)
    body_text, _ = slice_body_section(cleaned_text)

    return (
        full_text,
        cleaned_text,
        body_text,
        reference_text,
        ref_found,
        post_found,
    )


# ============================================================
# API KEY RESOLUTION
# ============================================================

def _get_openai_api_key():
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY")


def _get_openalex_api_key():
    # 1) Preferred: key captured by app.py login
    key = st.session_state.get("openalex_api_key", "").strip()
    if key:
        return key
    # 2) Fallback: secrets.toml / env
    try:
        if "OPENALEX_API_KEY" in st.secrets:
            return st.secrets["OPENALEX_API_KEY"]
    except Exception:
        pass
    return os.getenv("OPENALEX_API_KEY", "")


# ============================================================
# RENDER — entry point called by app.py
# ============================================================

def render():
    st.title("OmniCite Auditor — APA Style")

    # ---- API keys ----
    openai_key = _get_openai_api_key()
    if not openai_key:
        st.error(
            "OPENAI_API_KEY was not found. "
            "Add it to .streamlit/secrets.toml or your environment variables."
        )
        st.stop()

    openalex_key = _get_openalex_api_key()
    if not openalex_key:
        st.warning(
            "OpenAlex API key was not found — DOI verification will be skipped."
        )

    client = OpenAI(api_key=openai_key)

    # ---- Uploader ----
    if "apa_uploader_version" not in st.session_state:
        st.session_state["apa_uploader_version"] = 0

    uploaded_files = st.file_uploader(
        "Upload manuscript PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"apa_uploader_{st.session_state['apa_uploader_version']}",
    )

    if "pdf_batches" not in st.session_state:
        st.session_state["pdf_batches"] = {}

    if not uploaded_files:
        return

    # ---- Manual manuscript publication year (drives the 10-year window) ----
    current_year = datetime.now().year
    default_year = st.session_state.get("apa_manuscript_year", current_year)
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
        key="apa_manuscript_year_input",
    )
    st.session_state["apa_manuscript_year"] = int(manuscript_year)

    file_keys = []
    for idx, uf in enumerate(uploaded_files):
        key = f"{idx}::{uf.name}"
        file_keys.append((key, uf))

    # Drop batches no longer uploaded
    active_keys = {k for k, _ in file_keys}
    for k in list(st.session_state["pdf_batches"].keys()):
        if k not in active_keys:
            del st.session_state["pdf_batches"][k]

    # ---- Run pipeline ----
    if st.button(
        "Extract & Review",
        type="primary",
        use_container_width=True,
        key="blue_btn_run_all_pdfs",
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

            citations = extract_all_citations(body_text)
            stats = citation_statistics(citations)

            batch = {
                "filename": uf.name,
                "full_text": full_text,
                "cleaned_text": cleaned_text,
                "body_text": body_text,
                "reference_text": reference_text,
                "ref_found": ref_found,
                "post_found": post_found,
                "citations": citations,
                "citation_stats": stats,
                "manuscript_year": int(manuscript_year),
                "ai_done": False,
            }
            st.session_state["pdf_batches"][key] = batch

            overall.progress(
                base_pct + int(step * 0.35),
                text=f"[{i}/{n}] Extracting references & verifying DOIs via OpenAlex...",
            )
            try:
                process_single_pdf(
                    uf, batch, client, openalex_key, int(manuscript_year)
                )
            except Exception as exc:
                st.error(f"{uf.name} APA check failed: {exc}")
                continue

            overall.progress(
                base_pct + int(step * 1.00),
                text=f"[{i}/{n}] {uf.name} done.",
            )

        overall.progress(100, text="All manuscripts processed.")
        overall.empty()
        st.success(f"Processed {n} manuscript{'s' if n != 1 else ''}.")

    # ========================================================
    # RESULTS
    # ========================================================

    any_done = any(
        b.get("ai_done") for b in st.session_state["pdf_batches"].values()
    )
    if not any_done:
        return

    selector_options = [k for k, _ in file_keys]

    def _fmt(k):
        b = st.session_state["pdf_batches"][k]
        suffix = "" if b.get("ai_done") else "  (not yet processed)"
        return b["filename"] + suffix

    selected_key = st.selectbox(
        "Select manuscript to review",
        selector_options,
        format_func=_fmt,
        key="selected_pdf_key",
    )

    batch = st.session_state["pdf_batches"].get(selected_key)
    if not batch or not batch.get("ai_done"):
        st.info(
            "This manuscript has not been processed yet. "
            "Click 'Extract & Review' above."
        )
        return

    # Refresh the batch's manuscript_year with whatever is in the widget now
    batch["manuscript_year"] = int(st.session_state.get("apa_manuscript_year", current_year))

    with st.expander("View reference section (sent to AI)", expanded=False):
        st.text_area(
            "Reference slice",
            batch["reference_text"],
            height=300,
            key=f"dbg_ref_{selected_key}",
        )

    # ---- Metrics ----
    reference_rows = batch["reference_rows"]
    citation_rows = batch["citation_rows"]
    stats = batch["citation_stats"]
    citations = batch["citations"]
    manuscript_year = batch["manuscript_year"]

    total_refs = len(reference_rows)
    total_cits = stats["total"]
    doi_checked = sum(1 for x in reference_rows if x["DOI Checked"] == "YES")
    doi_suspicious = sum(1 for x in reference_rows if x["DOI Suspicious"] != "—")
    doi_reconstructed = sum(
        1 for x in reference_rows if x.get("Source") == "OpenAlex"
    )

    ref_keys = set()
    for ref in batch["reference_list"]:
        p = parse_reference(ref)
        if p["first_author"] and p["year"]:
            ref_keys.add((p["first_author"].lower(), p["year"]))

    missing_cits = 0
    for c in citations:
        if (c["author"].lower(), c["year"]) not in ref_keys:
            missing_cits += 1

    cit_keys = {(c["author"].lower(), c["year"]) for c in citations}
    missing_refs = 0
    for ref in batch["reference_list"]:
        p = parse_reference(ref)
        if p["first_author"] and p["year"]:
            if (p["first_author"].lower(), p["year"]) not in cit_keys:
                missing_refs += 1

    # ---- Revised: window anchored to manuscript_year, inclusive ----
    window_start = manuscript_year - 9
    window_end = manuscript_year
    recent = sum(
        1 for row in reference_rows
        if isinstance(row["Year"], int)
        and window_start <= row["Year"] <= window_end
    )
    recent_pct = (recent / total_refs * 100) if total_refs else 0

    metrics = [
        ("Total References", total_refs),
        ("References > 15",
         f"Yes ({total_refs})" if total_refs > 15 else f"No ({total_refs})"),
        ("Total In-text Citations", total_cits),
        ("Narrative Citations", stats["narrative"]),
        ("Parenthetical Citations", stats["parenthetical"]),
        (f"% Last 10 Years ({window_start}–{window_end})",
         f"{recent_pct:.1f}% ({recent}/{total_refs})" if total_refs else "0.0%"),
        ("Citations Missing from References", missing_cits),
        ("References Missing from Citations", missing_refs),
        ("DOI Checked (OpenAlex)", doi_checked),
        ("References Reconstructed from OpenAlex", doi_reconstructed),
        ("DOI Suspicious (possible fabrication)",
         f"{doi_suspicious} "
         f"({doi_suspicious / total_refs * 100:.1f}%)" if total_refs else "0"),
    ]

    metric_df = pd.DataFrame(metrics, columns=["Metric", "Value"])

    src_counts = Counter(row["Source Type"] for row in reference_rows)
    src_rows = []
    for stype in CANONICAL_SOURCE_TYPES:
        cnt = src_counts.get(stype, 0)
        pct = (cnt / total_refs * 100) if total_refs else 0
        src_rows.append({
            "Source Type": stype,
            "Count / Percentage": f"{cnt} ({pct:.1f}%)",
        })
    src_df = pd.DataFrame(src_rows)

    left, right = st.columns([1, 1], gap="large")
    with left:
        st.dataframe(metric_df, use_container_width=True, hide_index=True)
    with right:
        st.dataframe(src_df, use_container_width=True, hide_index=True)

    # ---- Download ----
    try:
        docx_bytes = build_apa_report_docx(batch)
        safe_name = re.sub(r"[^\w\-]+", "_", batch.get("filename", "manuscript"))
        st.download_button(
            label="📄 Download APA Diagnostic Report (.docx)",
            data=docx_bytes,
            file_name=f"{safe_name}_APA_report.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
            key=f"download_apa_report_docx_{selected_key}",
        )
    except Exception as exc:
        st.error(f"Could not build DOCX report: {exc}")

    # ---- Start Fresh button (red, below download) ----
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
        # 1. Wipe every APA-related session-state entry.
        for k in list(st.session_state.keys()):
            if k == "pdf_batches" or k.startswith("pdf_batches"):
                del st.session_state[k]
            if k.startswith("dbg_ref_"):
                del st.session_state[k]
            if k == "selected_pdf_key":
                del st.session_state[k]

        # 2. Recreate the batch dict empty.
        st.session_state["pdf_batches"] = {}

        # 3. Bump the uploader version → new key → empty uploader.
        st.session_state["apa_uploader_version"] = (
            st.session_state.get("apa_uploader_version", 0) + 1
        )

        # 4. Optional: reset the manuscript year back to current year.
        st.session_state["apa_manuscript_year"] = datetime.now().year

        # 5. Rerun so everything re-renders clean.
        st.rerun()