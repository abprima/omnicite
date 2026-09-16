# apa.py
import re
from datetime import datetime
import io
import os
import json
import hashlib
import traceback
import difflib
import shutil

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
# THRESHOLDS
# =========================================================

TITLE_STRONG_MATCH    = 0.78
TITLE_REVIEW_MATCH    = 0.55
AUTHOR_STRONG_MATCH   = 0.60
AUTHOR_REVIEW_MATCH   = 0.35
SEVERE_TITLE_MISMATCH  = 0.30
SEVERE_AUTHOR_MISMATCH = 0.25

TITLE_SEARCH_MIN_SIM  = 0.75
TITLE_SEARCH_YEAR_BONUS = 0.3

DOI_RESOLVER_URL     = "https://doi.org"
DOI_RESOLVER_TIMEOUT = 10

# ── Max output tokens for extraction (gpt-4o max is 16384)
EXTRACTION_MAX_TOKENS = 16384

# ── Long manuscripts (chars) get split into two calls
SPLIT_EXTRACTION_THRESHOLD = 25000

CACHE_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), ".cache", "apa_v3"
)
os.makedirs(CACHE_DIR, exist_ok=True)

# ── Diagnostic dump location
DEBUG_DUMP_PATH = "/tmp/ai_raw_response.txt"


# =========================================================
# REFERENCE SECTION DETECTION HEADINGS
# =========================================================

REFERENCE_HEADINGS = [
    "references", "reference", "reference list", "reference section",
    "literature cited", "literature", "works cited", "works consulted",
    "bibliography", "bibliographies", "cited references",
    "daftar pustaka", "daftar rujukan", "daftar bacaan",
    "rujukan", "rujukan pustaka", "bahan rujukan",
    "bibliografi", "referensi", "kepustakaan",
    "sumber rujukan", "sumber pustaka", "sumber referensi",
    "ref", "refs",
]


POST_REFERENCE_HEADINGS = {
    # Acknowledgments
    "acknowledgement", "acknowledgements",
    "acknowledgment", "acknowledgments",
    "acknowledgement of funding", "acknowledgment of funding",
    "acknowledgement section", "acknowledgment section",
    "ucapan terima kasih", "ucapan terimakasih",
    "ucapan terima kasih dan apresiasi",

    # Author contribution
    "author contribution", "author contributions",
    "authors contribution", "authors contributions",
    "author's contribution", "author's contributions",
    "authors' contribution", "authors' contributions",
    "author’s contribution", "author’s contributions",
    "authors’ contribution", "authors’ contributions",
    "contribution", "contributions",
    "credit author statement", "credit authorship contribution statement",
    "author contribution statement", "authorship statement",
    "kontribusi penulis", "kontribusi author",
    "pernyataan kontribusi penulis", "pernyataan kontribusi",

    # Author profile / bios
    "author profile", "authors profile",
    "author profiles", "authors profiles",
    "profile", "profiles",
    "biography", "biographies",
    "author biography", "author biographies",
    "about the authors", "about the author",
    "author bio", "authors bio",
    "biodata penulis", "profil penulis",

    # Conflicts
    "conflict of interest", "conflicts of interest",
    "conflict of interests", "conflicts of interests",
    "competing interest", "competing interests",
    "declaration of competing interest",
    "declaration of competing interests",
    "declaration of interest", "declaration of interests",
    "declarations of interest", "declarations of interests",
    "declaration", "declarations",
    "pernyataan konflik kepentingan", "konflik kepentingan",
    "pernyataan kepentingan",

    # Funding
    "funding", "funding information", "funding statement",
    "funding sources", "funding acknowledgements",
    "financial support", "financial disclosure",
    "sources of funding", "role of the funding source",
    "pendanaan", "pernyataan pendanaan",
    "sumber pendanaan", "sumber dana",

    # Data
    "data availability", "data availability statement",
    "availability of data", "availability of data and materials",
    "data and code availability", "data sharing statement",
    "code availability", "supplementary data",
    "ketersediaan data",

    # Ethics
    "ethical approval", "ethics approval",
    "ethics statement", "ethical statement",
    "ethics declarations", "ethical declarations",
    "informed consent", "consent for publication",
    "consent to participate", "consent statement",
    "persetujuan etik", "persetujuan etis",
    "pernyataan etik", "keterangan etik",
    "informed consent statement",

    # Disclosures / AI
    "disclosure", "disclosures", "disclosure statement",
    "declaration of generative ai",
    "declaration of generative ai use",
    "use of ai", "ai use statement",
    "penggunaan teknologi ai", "pernyataan penggunaan ai",
    "pernyataan penggunaan teknologi ai",

    # Appendices
    "appendix", "appendices", "appendix a", "appendix b", "appendix c",
    "supplementary material", "supplementary materials",
    "supplemental material", "supplemental materials",
    "supplementary information", "supporting information",
    "supporting information file",
    "lampiran", "lampiran a", "lampiran b",

    # Notes
    "notes", "note",
    "author note", "author notes",
    "endnotes", "endnote",
    "footnotes", "footnote",
    "catatan", "catatan kaki",

    # Correspondence
    "corresponding author", "corresponding author details",
    "correspondence", "address correspondence to",
    "reprint requests", "reprints",
    "orcid",
}


def normalize_heading(text):
    text = text.strip().lower()
    text = re.sub(r"^\s*(?:\d+(?:\.\d+)*)[\.\s:-]+", "", text)
    text = re.sub(r"^[\*\#•\-\s]+|[\*\#•\-\s]+$", "", text)
    text = re.sub(r"\s+", " ", text).rstrip(":").strip()
    text = text.rstrip(".:-—–").strip()
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
# AGGRESSIVE JSON RECOVERY
# =========================================================

def _safe_json_loads(text):
    """
    Aggressive JSON recovery for AI responses.
    Handles: markdown fences, leading/trailing prose, concatenated objects,
    trailing commas, unescaped control chars, BOM.
    """
    if text is None:
        raise ValueError("Empty AI response (None).")

    text = str(text)

    if text.startswith("\ufeff"):
        text = text[1:]

    # Strip markdown fences anywhere
    text = re.sub(r"```(?:json)?", "", text, flags=re.I)

    # Trim to first { and last }
    first = text.find("{")
    last = text.rfind("}")
    if first >= 0 and last > first:
        text = text[first:last + 1]

    text = text.strip()

    # 1. Straight parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 2. Remove control chars
    cleaned = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # 3. Fix trailing commas
    repaired = re.sub(r",(\s*[}\]])", r"\1", cleaned)
    try:
        return json.loads(repaired)
    except json.JSONDecodeError:
        pass

    # 4. Brace-counting extraction
    depth = 0
    start = None
    in_string = False
    escape = False
    for i, ch in enumerate(cleaned):
        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch == "{":
            if start is None:
                start = i
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0 and start is not None:
                candidate = cleaned[start:i + 1]
                for repair_fn in (
                    lambda s: s,
                    lambda s: re.sub(r",(\s*[}\]])", r"\1", s),
                    lambda s: re.sub(r"[\x00-\x1f]", " ", s),
                ):
                    try:
                        return json.loads(repair_fn(candidate))
                    except json.JSONDecodeError:
                        continue
                start = None
                depth = 0

    preview = text[:500].replace("\n", "\\n")
    raise ValueError(
        f"Could not parse JSON from AI response. "
        f"Length={len(text)}. First 500 chars: {preview!r}"
    )


# =========================================================
# PDF EXTRACTION
# =========================================================

def extract_pdf_text(uploaded_file):
    pdf_bytes = uploaded_file.read()
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []
    for page_number, page in enumerate(document, start=1):
        text = page.get_text("text")
        pages.append({"page": page_number, "text": text})
    document.close()
    full_text = "\n".join(
        f"<<<PAGE_BREAK:{p['page']}>>>\n{p['text']}" for p in pages
    )
    uploaded_file.seek(0)
    return full_text, pages


def strip_markdown_markers(text):
    if not text:
        return text
    text = _safe_sub(r"\*\*\*(.+?)\*\*\*", r"\1", text, flags=re.S)
    text = _safe_sub(r"\*\*(.+?)\*\*", r"\1", text, flags=re.S)
    text = _safe_sub(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"\1", text, flags=re.S)
    text = _safe_sub(r"(?<!_)_(?!\s)(.+?)(?<!\s)_(?!_)", r"\1", text, flags=re.S)
    text = re.sub(r"[ \t]{2,}", " ", text).strip()
    return text


def clean_text(text):
    text = text.replace("\u00ad", "")
    text = text.replace("\u2013", "–").replace("\u2014", "—")
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)
    return text


# =========================================================
# OPENAI CLIENT
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


def _openai_model():
    return os.getenv("OPENAI_MODEL", "gpt-4o")


def _cache_path(namespace, key_payload):
    key = hashlib.sha256(
        f"{namespace}||{key_payload}".encode("utf-8")
    ).hexdigest()
    return os.path.join(CACHE_DIR, f"{key}.json")


def _cache_get(path):
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _cache_set(path, value):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(value, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _clear_cache():
    try:
        shutil.rmtree(CACHE_DIR, ignore_errors=True)
        os.makedirs(CACHE_DIR, exist_ok=True)
    except Exception:
        pass


def _debug_dump(label, content):
    """Write raw AI output to disk for debugging."""
    try:
        with open(DEBUG_DUMP_PATH, "a", encoding="utf-8") as f:
            f.write(f"\n\n===== {label} @ {datetime.now().isoformat()} =====\n")
            f.write(content or "<empty>")
        print(f"[DEBUG] Wrote {label} to {DEBUG_DUMP_PATH} "
              f"({len(content or '')} chars)")
    except Exception as exc:
        print(f"[DEBUG] Could not write dump: {exc}")


# =========================================================
# STAGE 1 — AI EXTRACTION
# =========================================================

# ── Single-call prompt (short PDFs)
EXTRACTION_PROMPT = """You are analyzing the full text of an academic manuscript.

TASK
Extract:
  (A) Every entry in the REFERENCE LIST (Daftar Rujukan / References / Bibliography).
  (B) Every IN-TEXT CITATION in the body (parenthetical and narrative).

REFERENCE LIST RULES
- One entry = one bibliographic reference.
- Preserve the original text of each reference EXACTLY (typos included).
- If a reference spans multiple lines, join it into one string.
- Do NOT merge two references. Do NOT split one reference.
- Do NOT invent references. Do NOT include headings.
- Skip acknowledgments, author bios, funding statements.
- Two references by the SAME author in the SAME year are TWO separate entries.
- If you cannot tell whether something is a reference, put it in "uncertain".

IN-TEXT CITATION RULES
- Include both parenthetical "(Author, 2020)" and narrative "Author (2020)".
- Do NOT include citations inside the reference list itself.
- Preserve the exact text of each citation.
- Multiple sources inside one parenthetical are ONE citation string.
- Report "type" as "parenthetical" or "narrative".

OUTPUT FORMAT (JSON only, no markdown, no indentation)
{"references": ["...", "..."], "citations": [{"raw": "...", "type": "parenthetical"}], "uncertain": []}

FULL MANUSCRIPT TEXT
====================
{manuscript_text}
====================
"""


# ── Split prompts (long PDFs, called twice)
EXTRACTION_PROMPT_REFS = """You are extracting the REFERENCE LIST ONLY.

RULES
- Output every reference entry as a single string.
- Preserve original text exactly (typos included).
- Join multi-line references into one string.
- Do NOT merge two references, do NOT split one reference.
- Do NOT include headings, acknowledgments, funding, author bios.
- Do NOT include in-text citations.
- Do NOT invent entries.

OUTPUT (JSON only, no markdown)
{"references": ["...", "..."]}

FULL MANUSCRIPT TEXT
====================
{manuscript_text}
====================
"""


EXTRACTION_PROMPT_CITS = """You are extracting IN-TEXT CITATIONS ONLY.

RULES
- Include parenthetical "(Author, 2020)" and narrative "Author (2020)".
- Exclude anything inside the reference list itself.
- Preserve exact text of each citation.
- Multiple sources inside one pair of parentheses = ONE citation string.

OUTPUT (JSON only, no markdown)
{"citations": [{"raw": "...", "type": "parenthetical"}, {"raw": "...", "type": "narrative"}]}

FULL MANUSCRIPT TEXT
====================
{manuscript_text}
====================
"""


def _call_openai_for_json(prompt, label="extraction"):
    """Single OpenAI call → parsed JSON dict, or {"error": "..."}."""
    client = _get_openai_client()
    if client is None:
        return {"error": "OpenAI client unavailable (missing key)."}

    try:
        response = client.chat.completions.create(
            model=_openai_model(),
            messages=[
                {"role": "system",
                 "content": "Return compact JSON only. No markdown."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
            seed=42,
            max_tokens=EXTRACTION_MAX_TOKENS,
        )
        raw_content = response.choices[0].message.content
    except Exception as exc:
        return {"error": f"API call failed: {exc}"}

    # Dump for debugging
    _debug_dump(label, raw_content)

    try:
        data = _safe_json_loads(raw_content)
    except ValueError as exc:
        preview = (raw_content or "")[:200].replace("\n", "\\n")
        return {
            "error": (
                f"AI returned non-JSON output for {label}. "
                f"Preview: {preview!r} | parse error: {exc}"
            )
        }

    if not isinstance(data, dict):
        return {"error": f"AI returned JSON but not an object ({label})."}

    return data


def _extract_split(manuscript_text):
    """Two smaller calls for long PDFs."""
    refs_resp = _call_openai_for_json(
        EXTRACTION_PROMPT_REFS.format(manuscript_text=manuscript_text),
        label="references",
    )
    if "error" in refs_resp:
        return {"error": f"Reference extraction failed: {refs_resp['error']}"}

    cits_resp = _call_openai_for_json(
        EXTRACTION_PROMPT_CITS.format(manuscript_text=manuscript_text),
        label="citations",
    )
    if "error" in cits_resp:
        return {"error": f"Citation extraction failed: {cits_resp['error']}"}

    references = [
        r.strip() for r in refs_resp.get("references", [])
        if isinstance(r, str) and r.strip() and len(r) > 15
    ]

    citations = []
    for c in cits_resp.get("citations", []) or []:
        if not isinstance(c, dict):
            continue
        raw = (c.get("raw") or "").strip()
        ctype = (c.get("type") or "").lower().strip()
        if not raw:
            continue
        if ctype not in ("parenthetical", "narrative"):
            ctype = "parenthetical" if raw.startswith("(") else "narrative"
        citations.append({"raw": raw, "type": ctype})

    return {"references": references, "citations": citations, "uncertain": []}


def _extract_with_ai(manuscript_text):
    """
    Stage 1: AI extracts references + in-text citations.
    Uses split calls for long PDFs, single call for short ones.
    Cached by content hash. Returns dict or {"error": "..."} on failure.
    """
    client = _get_openai_client()
    if client is None:
        return {"error": "OpenAI client unavailable (missing key)."}

    model = _openai_model()
    cache_key = f"{model}::{len(manuscript_text)}::{manuscript_text[:2000]}"
    path = _cache_path("extract", cache_key)

    cached = _cache_get(path)
    if cached and isinstance(cached, dict) and "references" in cached:
        return cached

    # ── Long PDF → two calls
    if len(manuscript_text) > SPLIT_EXTRACTION_THRESHOLD:
        result = _extract_split(manuscript_text)
        if "error" not in result and result.get("references"):
            _cache_set(path, result)
        return result

    # ── Short PDF → single call
    prompt = EXTRACTION_PROMPT.format(manuscript_text=manuscript_text)
    data = _call_openai_for_json(prompt, label="extraction")
    if "error" in data:
        return data

    references = data.get("references", []) or []
    citations  = data.get("citations", []) or []

    references = [r.strip() for r in references
                  if isinstance(r, str) and r.strip()]
    references = [r for r in references if len(r) > 15]

    cleaned_citations = []
    for c in citations:
        if not isinstance(c, dict):
            continue
        raw = (c.get("raw") or "").strip()
        ctype = (c.get("type") or "").strip().lower()
        if not raw:
            continue
        if ctype not in ("parenthetical", "narrative"):
            ctype = "parenthetical" if raw.startswith("(") else "narrative"
        cleaned_citations.append({"raw": raw, "type": ctype})

    result = {
        "references": references,
        "citations": cleaned_citations,
        "uncertain": data.get("uncertain", []) or [],
    }

    if references or cleaned_citations:
        _cache_set(path, result)

    return result


# =========================================================
# PARSING HELPERS
# =========================================================

def parse_reference(reference):
    result = {
        "raw": reference, "year": None, "first_author": None,
        "authors": [], "doi": None, "url": None,
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
        if re.match(r"^https?://(?:dx\.)?doi\.org/", url, re.I):
            result["doi"] = normalize_doi_or_url(url)
            result["url"] = result["doi"]
        else:
            result["url"] = url
    else:
        doi_match = re.search(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+",
                              reference, re.I)
        if doi_match:
            result["doi"] = normalize_doi_or_url(doi_match.group(0))
            result["url"] = result["doi"]

    return result


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
        'https://doi.org/', text,
    )
    doi_match = re.search(r'10\.\d{4,9}/[-._;()/:A-Za-z0-9]+', text, re.I)
    if doi_match:
        doi = doi_match.group(0).rstrip(".,;)")
        return f"https://doi.org/{doi}"
    url_match = re.search(r'https?://\S+', text, re.I)
    if url_match:
        return url_match.group(0).rstrip(".,;)")
    return text


def parse_citation(citation):
    raw = citation.get("raw", "")
    ctype = citation.get("type", "parenthetical")

    year_match = re.search(r"\b((?:19|20)\d{2})[a-z]?\b", raw)
    year = year_match.group(1) if year_match else None

    inner = raw
    if ctype == "parenthetical":
        inner = re.sub(r"^\(|\)$", "", raw).strip()

    author = None
    if ctype == "parenthetical":
        first = inner.split(";")[0].strip()
        m = re.match(r"^([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)", first)
        if m:
            author = m.group(1)
    else:
        m = re.match(r"^([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)", raw)
        if m:
            author = m.group(1)

    return {"raw": raw, "type": ctype, "author": author, "year": year}


def calculate_citation_statistics(citations):
    narrative = sum(1 for c in citations if c.get("type") == "narrative")
    parenthetical = sum(1 for c in citations if c.get("type") == "parenthetical")
    total = narrative + parenthetical
    return {
        "total": total, "narrative": narrative, "parenthetical": parenthetical,
        "narrative_percentage": (narrative / total * 100 if total else 0),
        "parenthetical_percentage": (parenthetical / total * 100 if total else 0),
    }


def normalize(value):
    if not value:
        return ""
    return re.sub(r"[^a-z0-9]", "", value.lower().strip())


def match_citations_to_references(citations, parsed_references):
    results = []
    keys = {(normalize(c.get("author")), c.get("year")) for c in citations}
    for i, ref in enumerate(parsed_references, start=1):
        key = (normalize(ref["first_author"]), ref["year"])
        results.append({
            "Reference #": i, "Author": ref["first_author"],
            "Year": ref["year"], "Cited": key in keys, "Reference": ref["raw"],
        })
    return results


def find_missing_references(citations, parsed_references):
    ref_keys = {(normalize(r["first_author"]), r["year"])
                for r in parsed_references}
    missing, seen = [], set()
    for c in citations:
        key = (normalize(c.get("author")), c.get("year"))
        if key not in ref_keys and key not in seen and c.get("author"):
            missing.append({
                "Citation": c.get("raw"),
                "Author": c.get("author"),
                "Year": c.get("year"),
            })
            seen.add(key)
    return missing


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
        rows.append({
            "Reference #": i,
            "Year": year if year is not None else "Not detected",
            "Recency": category,
            "Reference": ref,
        })
    total = len(references)
    return {
        "start_year": start_year, "end_year": manuscript_year, "total": total,
        "recent_count": recent, "older_count": older,
        "unknown_count": unknown, "future_count": future,
        "recent_percentage": recent / total * 100 if total else 0,
        "rows": rows,
        "year_summary": [{"Year": y, "References": c} for y, c in
                         sorted(year_counts.items(), key=lambda x: x[0],
                                reverse=True)],
    }


# =========================================================
# STAGE 2 — OPENALEX VERIFICATION
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


def _openalex_result_to_meta(item):
    if not item:
        return None
    authors = []
    for a in item.get("authorships", []) or []:
        name = (a.get("author") or {}).get("display_name")
        if name:
            authors.append(name)
    primary = item.get("primary_location") or {}
    source = primary.get("source") or {}
    biblio = item.get("biblio") or {}
    return {
        "_not_found": False,
        "title": item.get("title") or "",
        "year": item.get("publication_year"),
        "authors": authors,
        "doi": item.get("doi") or "",
        "journal": source.get("display_name"),
        "source_type": source.get("type"),
        "volume": biblio.get("volume"),
        "issue": biblio.get("issue"),
        "first_page": biblio.get("first_page"),
        "last_page": biblio.get("last_page"),
        "publisher": source.get("host_organization_name"),
        "is_oa": item.get("open_access", {}).get("is_oa"),
    }


def _verify_doi_resolution(doi):
    if not doi:
        return {"resolves": None, "reason": "No DOI supplied"}
    try:
        import requests
    except Exception:
        return {"resolves": None, "reason": "requests missing"}
    clean = _normalize_doi_for_lookup(doi)
    if not clean:
        return {"resolves": None, "reason": "DOI not parseable"}
    url = f"{DOI_RESOLVER_URL}/{clean}"
    headers = {"User-Agent": "OmniCite/3.0"}
    try:
        r = requests.get(url, headers=headers, allow_redirects=True,
                         timeout=DOI_RESOLVER_TIMEOUT)
    except Exception as exc:
        return {"resolves": None, "reason": f"resolver unavailable: {exc}"}
    if 200 <= r.status_code < 400:
        return {"resolves": True, "reason": "DOI resolves", "final_url": r.url}
    if r.status_code in (404, 410):
        return {"resolves": False, "reason": f"HTTP {r.status_code}"}
    return {"resolves": None, "reason": f"Inconclusive HTTP {r.status_code}"}


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
    ext_surnames = {_surname_from_full_name(n)
                    for n in (external.get("authors") or []) if n}
    ext_surnames.discard("")
    a = None
    if ref_surnames and ext_surnames:
        a = len(ref_surnames & ext_surnames) / max(1, len(ref_surnames | ext_surnames))
    sy = str(submitted.get("year") or "").strip()
    ey = str(external.get("year") or "").strip()
    year_match = (sy == ey) if sy and ey else None
    return {"title_similarity": t, "author_similarity": a, "year_match": year_match}


def _decide_doi_status_advanced(comparison):
    t = comparison.get("title_similarity")
    a = comparison.get("author_similarity")
    y = comparison.get("year_match")
    severe_title  = t is not None and t < SEVERE_TITLE_MISMATCH
    severe_author = a is not None and a < SEVERE_AUTHOR_MISMATCH
    year_wrong    = y is False

    if severe_title and severe_author and year_wrong:
        return ("WITHHELD",
                "DOI exists but title, authors, and year strongly conflict "
                "with OpenAlex metadata. Possible fabricated citation.")
    title_good  = t is None or t >= TITLE_STRONG_MATCH
    author_good = a is None or a >= AUTHOR_STRONG_MATCH
    year_good   = y is None or y is True
    if title_good and author_good and year_good:
        return ("VERIFIED", "DOI exists and matches external metadata.")
    title_review  = t is None or t >= TITLE_REVIEW_MATCH
    author_review = a is None or a >= AUTHOR_REVIEW_MATCH
    if title_review and author_review:
        return ("VERIFIED_METADATA_CORRECTED",
                "Source identifiable; some metadata fields need correction.")
    return ("MANUAL_CHECK",
            "DOI exists but reference does not match strongly enough.")


def verify_reference_against_openalex(reference, parsed):
    result = {
        "checked": False, "suspicious": False,
        "title_similarity": None, "author_similarity": None,
        "year_match": None, "status": "UNVERIFIED",
        "reasons": [], "crossref_title": None,
        "crossref_authors": [], "source_of_truth": None,
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
                "DOI does not resolve in OpenAlex OR doi.org. "
                "Possible fabricated DOI."
            )
            return result
        if resolution["resolves"] is True:
            result["checked"] = True
            result["status"] = "UNVERIFIED"
            result["reasons"].append("DOI resolves but not in OpenAlex.")
            return result
        result["checked"] = True
        result["status"] = "MANUAL_CHECK"
        result["reasons"].append("DOI verification inconclusive.")
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

    status, reason = _decide_doi_status_advanced(comparison)
    result["status"] = status
    result["reasons"].append(reason)
    if status == "WITHHELD":
        result["suspicious"] = True
    return result


def resolve_canonical_metadata(reference, parsed):
    result = {
        "found": False, "source": None, "confidence": "unverified",
        "meta": None, "verification": None,
    }

    doi = parsed.get("doi")
    if doi:
        meta = _fetch_openalex_metadata(doi)
        if meta and not meta.get("_not_found"):
            verification = verify_reference_against_openalex(reference, parsed)
            result["found"] = True
            result["source"] = "openalex_doi"
            result["confidence"] = (
                "high" if verification.get("status") in
                ("VERIFIED", "VERIFIED_METADATA_CORRECTED")
                else "medium"
            )
            result["meta"] = meta
            result["verification"] = verification
            return result

    title = _extract_apa_title(reference)
    if title and len(title) >= 15:
        candidates = _openalex_search_by_title(title, max_results=5)
        if candidates:
            best = None
            best_score = 0.0
            ref_year = str(parsed.get("year") or "").strip()
            ref_surnames = {s.lower() for s in (parsed.get("authors") or []) if s}
            for cand in candidates:
                cand_title = cand.get("title") or ""
                cand_year = str(cand.get("publication_year") or "").strip()
                t_sim = _title_similarity(title, cand_title)
                if t_sim < TITLE_SEARCH_MIN_SIM:
                    continue
                score = t_sim
                if ref_year and cand_year == ref_year:
                    score += TITLE_SEARCH_YEAR_BONUS
                cand_authors = []
                for a in cand.get("authorships", []) or []:
                    name = (a.get("author") or {}).get("display_name")
                    if name:
                        cand_authors.append(name)
                cand_surnames = {_surname_from_full_name(n) for n in cand_authors}
                cand_surnames.discard("")
                if ref_surnames and cand_surnames:
                    if not (ref_surnames & cand_surnames):
                        continue
                if score > best_score:
                    best_score = score
                    best = cand
            if best:
                meta = _openalex_result_to_meta(best)
                result["found"] = True
                result["source"] = "openalex_title"
                result["confidence"] = "medium"
                result["meta"] = meta
                result["verification"] = {
                    "checked": True, "suspicious": False,
                    "status": "VERIFIED_TITLE_MATCH",
                    "reasons": [f"Matched by title search (score={best_score:.2f})"],
                    "crossref_title": meta.get("title"),
                    "crossref_authors": meta.get("authors", []),
                    "title_similarity": _title_similarity(title, meta.get("title") or ""),
                    "author_similarity": None,
                    "year_match": (str(meta.get("year") or "") == ref_year
                                   if ref_year else None),
                    "source_of_truth": "openalex",
                }
                return result

    return result


# =========================================================
# STAGE 3 — AI APA FORMATTING FROM CANONICAL METADATA
# =========================================================

APA_FORMAT_PROMPT = """You are formatting ONE APA 7th edition reference-list entry.

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
  3. Convert author list to APA 7 form.
  4. Apply APA 7 capitalization (sentence case for titles, title case for journal).
  5. Preserve the [translation] bracket if present.
  6. Choose the exact italic spans.

SOURCE TYPE (choose one):
  "Journal Article", "Book", "Book Chapter", "Conference Proceeding",
  "Report", "Webpage / Online Document", "Other"

APA 7 ITALIC RULES:
  - Journal Article -> journal name + volume number (issue NOT italic, pages NOT italic)
  - Book           -> book title
  - Book Chapter   -> book title
  - Report         -> report title
  - Webpage        -> webpage title
  - Conference     -> proceedings title

CRITICAL — `italic_elements` MUST be a comma-separated list of LITERAL
substrings appearing VERBATIM inside `corrected_reference`.

Return JSON ONLY:
{{
  "corrected_reference": "...",
  "italic_elements": "literal, substrings, to, italicize",
  "source_type": "one of the source types above",
  "explanation": "short note"
}}

GROUND TRUTH (do NOT alter):
{ground_truth}

ORIGINAL REFERENCE (for reference only — do NOT copy its errors):
{original_reference}
"""


def format_reference_with_ai(openalex_meta, original_reference, source_type):
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

    prompt = APA_FORMAT_PROMPT.format(
        ground_truth=json.dumps(ground_truth, ensure_ascii=False, indent=2),
        original_reference=original_reference,
    )

    try:
        response = client.chat.completions.create(
            model=_openai_model(),
            messages=[
                {"role": "system",
                 "content": "You format APA 7 references from verified "
                            "metadata. Never invent facts. JSON only."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
            seed=42,
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


def _reference_is_hallucinated(original, corrected):
    if not corrected:
        return False
    def numbers(text):
        return set(re.findall(r"\b\d+(?:\.\d+)?\b", text or ""))
    def dois(text):
        return set(re.findall(r"10\.\d{4,9}/[-._;()/:A-Za-z0-9]+",
                              text or "", re.I))
    def urls(text):
        return set(re.findall(r"https?://\S+", text or "", re.I))
    if numbers(corrected) - numbers(original):
        return True
    if dois(corrected) - dois(original):
        return True
    if urls(corrected) - urls(original):
        return True
    return False


# =========================================================
# IN-TEXT CITATION CORRECTION
# =========================================================

_YEAR_TOKEN_RE = re.compile(r"\b((?:19|20)\d{2}[a-z]?)\b")


def revise_parenthetical_citation_local(content):
    parts = [p.strip() for p in content.split(";") if p.strip()]
    if not parts:
        return "(" + content.strip() + ")"
    corrected = []
    for part in parts:
        ym = _YEAR_TOKEN_RE.search(part)
        if not ym:
            corrected.append(part)
            continue
        year = ym.group(1)
        author_text = part[:ym.start()].strip(" ,")
        surnames = re.findall(
            r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\b", author_text
        )
        surnames = [s for s in surnames if s.lower() not in
                    {"and", "see", "cf", "et", "al"}]
        has_et_al = bool(re.search(r"\bet\s+al\.", author_text, re.I))
        if not surnames:
            corrected.append(part)
            continue
        if has_et_al or len(surnames) >= 3:
            author_final = f"{surnames[0]} et al."
        elif len(surnames) == 2:
            author_final = f"{surnames[0]} & {surnames[1]}"
        else:
            author_final = surnames[0]
        corrected.append(f"{author_final}, {year}")
    if len(corrected) > 1:
        corrected.sort(key=lambda s: s.lower())
    return "(" + "; ".join(corrected) + ")"


def revise_narrative_citation_local(raw):
    ym = _YEAR_TOKEN_RE.search(raw)
    if not ym:
        return raw
    year = ym.group(1)
    before = raw[:ym.start()].strip(" ,")
    surnames = re.findall(
        r"\b([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)\b", before
    )
    surnames = [s for s in surnames if s.lower() not in
                {"and", "see", "cf", "et", "al"}]
    if not surnames:
        return raw
    if len(surnames) >= 2:
        if "&" in before:
            return f"{surnames[0]} and {surnames[1]} ({year})"
        return raw
    return f"{surnames[0]} ({year})"


def build_citation_correction(citation):
    raw = citation.get("raw", "").strip()
    ctype = citation.get("type", "parenthetical")
    if ctype == "parenthetical":
        inner = re.sub(r"^\(|\)$", "", raw).strip()
        corrected = revise_parenthetical_citation_local(inner)
    else:
        corrected = revise_narrative_citation_local(raw)
    status = "MATCH" if corrected.strip() == raw.strip() else "REVISED"
    return {"Corrected": corrected, "Status": status}


def review_citations_with_ai(citations):
    client = _get_openai_client()
    if client is None or not citations:
        return []
    payload = [
        {"number": i, "citation": c.get("raw", ""), "type": c.get("type", "")}
        for i, c in enumerate(citations, start=1)
    ]
    prompt = f"""
Correct each APA 7 in-text citation IN ISOLATION.
- Two authors: parenthetical uses "&"; narrative uses "and".
- Three or more authors: "Surname et al."
- Never merge or split citations. Never invent authors/years.

Return JSON only:
{{"results": [{{"number": int, "status": "OK"|"REVISED"|"MANUAL_CHECK",
"revised_citation": str, "explanation": str}}, ...]}}

INPUT:
{json.dumps(payload, ensure_ascii=False)}
"""
    try:
        response = client.chat.completions.create(
            model=_openai_model(),
            messages=[
                {"role": "system",
                 "content": "Precise APA 7 in-text citation editor. JSON only."},
                {"role": "user", "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,
            seed=42,
        )
        return _safe_json_loads(response.choices[0].message.content).get("results", [])
    except Exception as exc:
        return [{"number": 0, "status": "MANUAL_CHECK",
                 "revised_citation": "", "explanation": f"API error: {exc}"}]


def _citation_is_hallucinated(original, corrected):
    def surnames(text):
        return set(re.findall(r"\b([A-ZÀ-ÖØ-Ý][a-zà-ÿ]+)\b", text or ""))
    def years(text):
        return set(re.findall(r"\b((?:19|20)\d{2})\b", text or ""))
    return (surnames(corrected) - surnames(original)) or \
           (years(corrected) - years(original))


def build_citation_comparison(citations, ai_results=None):
    by_no = {int(x.get("number", -1)): x
             for x in (ai_results or [])
             if str(x.get("number", "")).isdigit()}
    rows = []
    for i, citation in enumerate(citations, start=1):
        local = build_citation_correction(citation)
        ai = by_no.get(i, {})
        original = strip_markdown_markers(citation.get("raw", ""))
        ai_corrected = strip_markdown_markers(
            ai.get("revised_citation", "")
        ) if ai else ""

        use_ai = (
            ai_corrected
            and ai_corrected != original
            and not _citation_is_hallucinated(original, ai_corrected)
        )

        if use_ai:
            corrected = ai_corrected
            status = _apa_status_label(ai.get("status")) if ai else "REVISED"
            note = ai.get("explanation", "") or local.get("Status", "")
        else:
            corrected = local["Corrected"]
            status = local["Status"]
            note = ""

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


def _apa_status_label(status):
    status = str(status or "MANUAL_CHECK").upper().strip()
    if status in {"OK", "PASS", "MATCH", "VERIFIED"}:
        return "MATCH"
    if status in {"REVISED", "NEEDS REVIEW", "NEEDS_REVIEW",
                  "VERIFIED_METADATA_CORRECTED"}:
        return "REVISED"
    return "MANUAL CHECK"


# =========================================================
# REFERENCE COMPARISON TABLE
# =========================================================

def build_reference_comparison(references, manuscript_year):
    rows = []
    for i, original in enumerate(references, start=1):
        original_clean = strip_markdown_markers(clean_text(original))
        parsed = parse_reference(original_clean)
        canonical = resolve_canonical_metadata(original_clean, parsed)

        placeholders = {}
        correction_note = ""
        italic_elements = ""
        source_type = detect_apa_source_type(original_clean)
        corrected = original_clean
        status = "UNVERIFIED"

        if canonical["found"]:
            meta = canonical["meta"]
            verification = canonical["verification"] or {}
            source = canonical["source"]

            if (source == "openalex_doi"
                    and verification.get("suspicious")):
                corrected = "— WITHHELD (DOI mismatch) —"
                correction_note = (verification.get("reasons") or [""])[0]
                status = "WITHHELD"
            else:
                source_type = normalize_source_type(
                    _source_type_from_openalex(meta) or source_type
                )
                ai_formatted = format_reference_with_ai(
                    meta, original_clean, source_type
                )
                if ai_formatted.get("corrected_reference"):
                    corrected = ai_formatted["corrected_reference"]
                    italic_elements = ai_formatted["italic_elements"]
                    source_type = ai_formatted["source_type"]
                    correction_note = (ai_formatted.get("explanation")
                                       or f"Formatted from {source}.")
                else:
                    corrected = original_clean
                    correction_note = ai_formatted.get("explanation", "")
                if source == "openalex_doi":
                    status = verification.get("status", "VERIFIED")
                else:
                    status = "VERIFIED_TITLE_MATCH"
        else:
            corrected = original_clean.rstrip(".") + "."
            correction_note = "No canonical metadata found; reference kept as-is."
            status = "UNVERIFIED"

        year = extract_reference_year(corrected)
        verification = canonical.get("verification") or {}
        meta = canonical.get("meta")

        rows.append({
            "No.": i,
            "Source Type": source_type,
            "Year": year,
            "Original Reference": original_clean,
            "Corrected Version": corrected,
            "Italicized in APA": italic_elements,
            "Status": status,
            "Canonical Source": canonical.get("source") or "—",
            "Confidence": canonical.get("confidence", "unverified"),
            "Verification Status": verification.get("status", "UNVERIFIED"),
            "Verification Reason": " | ".join(verification.get("reasons", [])),
            "Source of Truth": verification.get("source_of_truth") or "local",
            "Correction Note": correction_note,
            "Placeholders": placeholders,
            "DOI Verified": verification.get("checked", False),
            "DOI Suspicious": verification.get("suspicious", False),
            "DOI Verification Reasons": " | ".join(verification.get("reasons", [])),
            "Title Similarity": verification.get("title_similarity"),
            "Author Similarity": verification.get("author_similarity"),
            "Year Match": verification.get("year_match"),
            "OpenAlex Title": meta.get("title") if meta else None,
            "OpenAlex Authors": ", ".join((meta.get("authors") or [])[:5]) if meta else None,
            "OpenAlex Journal": meta.get("journal") if meta else None,
            "OpenAlex Volume": meta.get("volume") if meta else None,
            "OpenAlex Issue": meta.get("issue") if meta else None,
            "OpenAlex Pages": (
                f"{meta.get('first_page')}–{meta.get('last_page')}"
                if meta and meta.get("first_page") else None
            ),
        })
    return rows


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


def detect_apa_source_type(reference):
    low = (reference or "").lower()
    if re.search(r"\(eds?\.\)", reference, re.I):
        return "Book Chapter"
    if re.search(r"\b(proceedings|conference|symposium)\b", low):
        return "Conference Proceeding"
    if re.search(r"\b(report|technical report|working paper)\b", low):
        return "Report"
    if re.search(r"https?://", reference) and "doi.org" not in low:
        return "Webpage / Online Document"
    if re.search(r",\s*\d+\s*(?:\([^)]+\))?\s*,\s*\d+", reference):
        return "Journal Article"
    return "Other"


# =========================================================
# DOCX EXPORT
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
    _set_run_font(r, size_pt=size_pt, bold=bold, italic=italic,
                  color=RED if red else None)
    return r


def _add_red_italic_run(paragraph, text, size_pt=11):
    _add_run(paragraph, text, size_pt=size_pt, bold=True, italic=True, red=True)


def _add_styled_reference(paragraph, text, italic_elements):
    if isinstance(italic_elements, list):
        tokens = [t.strip() for t in italic_elements if t and t.strip()]
    else:
        tokens = [t.strip() for t in (italic_elements or "").split(",") if t.strip()]
    if not tokens:
        _add_run(paragraph, text, size_pt=11)
        return

    escaped = sorted({_esc(t) for t in tokens}, key=len, reverse=True)
    pattern = _safe_compile(
        r"(?<!\w)(" + "|".join(escaped) + r")(?!\w)", re.I,
    )
    if pattern is None:
        _add_run(paragraph, text, size_pt=11)
        return

    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            _add_run(paragraph, text[pos:m.start()], size_pt=11)
        _add_run(paragraph, m.group(0), size_pt=11, italic=True)
        pos = m.end()
    if pos < len(text):
        _add_run(paragraph, text[pos:], size_pt=11)


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

    # --- Section 1: In-text Citation List
    h1 = doc.add_heading(level=1)
    _set_run_font(h1.add_run("1. In-text Citation List"), size_pt=14, bold=True)

    citation_rows = result.get("citation_comparison") or []
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

            raw = row.get("Original Citation", "")
            author, year = "", None
            m = re.match(r"^\(?([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)", raw)
            if m:
                author = m.group(1)
            ym = re.search(r"\b((?:19|20)\d{2})\b", raw)
            if ym:
                year = ym.group(1)
            missing = (normalize(author), year) in missing_citation_keys

            _add_run(p, "Original : ", bold=True)
            _add_run(p, raw)
            if missing:
                _add_red_italic_run(p, "  ← MISSING FROM REFERENCES")

            p2 = doc.add_paragraph()
            p2.paragraph_format.space_after = Pt(8)
            _add_run(p2, "Corrected: ", bold=True)
            _add_run(p2, row.get("Revised Citation", ""))

    doc.add_page_break()

    # --- Section 2: Reference List
    h2 = doc.add_heading(level=1)
    _set_run_font(h2.add_run("2. Reference List (APA 7th Edition)"),
                  size_pt=14, bold=True)

    reference_rows = result.get("reference_comparison") or []
    matching = result.get("matching_results", []) or []
    uncited = {row.get("Reference #") for row in matching if not row.get("Cited")}

    if not reference_rows:
        p = doc.add_paragraph()
        _set_run_font(p.add_run("No references were detected."), italic=True)
    else:
        def sort_key(row):
            txt = row.get("Corrected Version") or row.get("Original Reference", "")
            m = re.match(r"^([A-Za-zÀ-ÖØ-öø-ÿ'’\-]+)", txt)
            return m.group(1).lower() if m else txt.lower()

        for row in sorted(reference_rows, key=sort_key):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.5)

            original = row.get("Original Reference", "")
            if row.get("DOI Suspicious"):
                _add_run(p, original, red=True)
            else:
                _add_run(p, original)
            if row.get("No.") in uncited:
                _add_red_italic_run(p, "  ← NOT CITED IN TEXT")

            if row.get("DOI Suspicious"):
                withheld = doc.add_paragraph()
                withheld.paragraph_format.left_indent = Inches(0.5)
                _add_run(withheld,
                         "Corrected version withheld — DOI mismatch. "
                         "Manual verification required.",
                         size_pt=10, italic=True, bold=True, red=True)
            else:
                p2 = doc.add_paragraph()
                p2.paragraph_format.left_indent = Inches(0.5)
                _add_run(p2, "Corrected: ", bold=True)
                _add_styled_reference(p2, row.get("Corrected Version", ""),
                                      row.get("Italicized in APA", ""))
                src_tag = f"  [{row.get('Canonical Source', '—')} · " \
                          f"{row.get('Confidence', 'unverified')}]"
                _add_run(p2, src_tag, size_pt=9, italic=True)

                vstatus = row.get("Verification Status", "")
                if vstatus and vstatus not in ("VERIFIED",):
                    sp = doc.add_paragraph()
                    sp.paragraph_format.left_indent = Inches(0.5)
                    _add_run(sp,
                             f"Verification: {vstatus} — "
                             f"{row.get('Verification Reason', '')}",
                             size_pt=9, italic=True,
                             red=(vstatus in ("MANUAL_CHECK", "WITHHELD")))

                note = row.get("Correction Note", "")
                if note:
                    np_ = doc.add_paragraph()
                    np_.paragraph_format.left_indent = Inches(0.5)
                    _add_run(np_, f"Fix applied: {note}", size_pt=10, italic=True)

            if row.get("DOI Suspicious"):
                wp = doc.add_paragraph()
                wp.paragraph_format.left_indent = Inches(0.5)
                _add_run(wp,
                         f"⚠ Possible fabricated reference: "
                         f"{row.get('DOI Verification Reasons', '')}",
                         size_pt=10, italic=True, bold=True, red=True)

    # --- Section 3: Summary
    doc.add_page_break()
    h3 = doc.add_heading(level=1)
    _set_run_font(h3.add_run("3. Summary"), size_pt=14, bold=True)

    stats = result.get("citation_stats", {})
    recency = result.get("recency", {})
    total_refs = len(result.get("references", []))
    doi_checked = sum(1 for r in reference_rows if r.get("DOI Verified"))
    doi_sus = sum(1 for r in reference_rows if r.get("DOI Suspicious"))
    high_conf = sum(1 for r in reference_rows if r.get("Confidence") == "high")
    med_conf = sum(1 for r in reference_rows if r.get("Confidence") == "medium")
    unv_conf = sum(1 for r in reference_rows if r.get("Confidence") == "unverified")

    summary_rows = [
        ("Total references", str(total_refs), False),
        ("Total in-text citations", str(stats.get("total", 0)), False),
        ("  • Narrative", str(stats.get("narrative", 0)), False),
        ("  • Parenthetical", str(stats.get("parenthetical", 0)), False),
        ("Citations missing from references",
         str(len(result.get("missing_references", []))), False),
        ("References missing from citations", str(len(uncited)), False),
        ("  • HIGH confidence (DOI verified)", str(high_conf), False),
        ("  • MEDIUM confidence (title match)", str(med_conf), False),
        ("  • UNVERIFIED (local only)", str(unv_conf), False),
        ("DOI checked via OpenAlex", str(doi_checked), False),
        ("DOI suspicious (possible fabricated references)",
         str(doi_sus), doi_sus > 0),
        ("% references within last 10 years "
         f"({recency.get('start_year')}-{recency.get('end_year')})",
         f"{recency.get('recent_percentage', 0):.1f}%", False),
    ]

    table = doc.add_table(rows=1, cols=2)
    table.style = "Light Grid Accent 1"
    table.autofit = True
    hdr = table.rows[0].cells
    for cell, text in zip(hdr, ["Metric", "Value"]):
        cell.text = ""
        run = cell.paragraphs[0].add_run(text)
        _set_run_font(run, size_pt=10, bold=True)

    for label, value, warn in summary_rows:
        cells = table.add_row().cells
        cells[0].text = ""
        rl = cells[0].paragraphs[0].add_run(label)
        _set_run_font(rl, size_pt=10, bold=warn,
                      color=RED if warn else None)
        cells[1].text = ""
        rv = cells[1].paragraphs[0].add_run(value)
        _set_run_font(rv, size_pt=10, bold=warn,
                      color=RED if warn else None)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


# =========================================================
# MAIN PIPELINE
# =========================================================

def analyze_manuscript(uploaded_file, manuscript_year):
    uploaded_file.seek(0)
    full_text, _pages = extract_pdf_text(uploaded_file)
    full_text = clean_text(full_text)

    if not full_text.strip():
        raise ValueError("PDF contained no extractable text.")

    ai_extract = _extract_with_ai(full_text)
    if not ai_extract:
        raise ValueError("OpenAI extraction failed (no response).")
    if "error" in ai_extract:
        raise ValueError(ai_extract["error"])

    references = ai_extract.get("references", [])
    raw_citations = ai_extract.get("citations", [])

    if not references:
        raise ValueError("OpenAI found 0 references in the manuscript.")

    citations = [parse_citation(c) for c in raw_citations]
    reference_rows = build_reference_comparison(references, manuscript_year)
    citation_ai = review_citations_with_ai(citations)
    citation_rows = build_citation_comparison(citations, citation_ai)

    parsed_refs = [parse_reference(r) for r in references]
    matching = match_citations_to_references(citations, parsed_refs)
    missing = find_missing_references(citations, parsed_refs)
    citation_stats = calculate_citation_statistics(citations)
    recency = calculate_reference_recency(references, manuscript_year)

    return {
        "filename": uploaded_file.name,
        "references": references,
        "parsed_references": parsed_refs,
        "citations": citations,
        "citation_stats": citation_stats,
        "matching_results": matching,
        "missing_references": missing,
        "recency": recency,
        "reference_comparison": reference_rows,
        "citation_comparison": citation_rows,
        "manuscript_year": manuscript_year,
        "ai_complete": True,
    }


# =========================================================
# UI RENDER
# =========================================================

def _sanitize_error(err):
    """Never leak more than 200 chars of raw AI output."""
    if not err:
        return "Unknown error."
    err = str(err)
    # Trim raw dumps from the error string
    for marker in ("Preview:", "First 500 chars:"):
        if marker in err:
            err = err.split(marker)[0].strip()
            err += " (see server logs / /tmp/ai_raw_response.txt for full raw response)"
    if len(err) > 400:
        err = err[:400] + "..."
    return err


def render():
    st.title("OmniCite Auditor - APA 7th Edition")
    st.caption("AI extraction → OpenAlex verification → AI formatting")

    uploaded_files = st.file_uploader(
        "Upload manuscript PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload up to 5 manuscripts.",
        key="apa_file_uploader",
    )

    manuscript_year = st.number_input(
        "Manuscript publication year",
        min_value=1900, max_value=2100,
        value=datetime.now().year, step=1,
        key="apa_manuscript_year",
    )

    if "apa_batch_results" not in st.session_state:
        st.session_state["apa_batch_results"] = {}

    # ── Cache management (always visible)
    with st.expander("🔧 Cache & Debug", expanded=False):
        colA, colB = st.columns(2)
        with colA:
            if st.button("🗑 Clear AI cache", use_container_width=True):
                _clear_cache()
                st.success("Cache cleared. Re-run extraction.")
        with colB:
            if st.button("📄 Show debug dump", use_container_width=True):
                if os.path.exists(DEBUG_DUMP_PATH):
                    try:
                        with open(DEBUG_DUMP_PATH, "r", encoding="utf-8") as f:
                            content = f.read()
                        st.text_area("Raw AI responses",
                                     content[-8000:], height=300)
                    except Exception as exc:
                        st.error(f"Could not read debug dump: {exc}")
                else:
                    st.info("No debug dump yet. Run extraction first.")

    if not uploaded_files:
        return

    if len(uploaded_files) > 5:
        st.error("Maximum batch size is 5 PDFs.")
        st.stop()

    if st.button(
        f"Extract & Review ({len(uploaded_files)} PDF"
        f"{'s' if len(uploaded_files) != 1 else ''})",
        use_container_width=True,
        key="apa_extract_button",
    ):
        if _get_openai_client() is None:
            st.error("OPENAI_API_KEY is not set. This pipeline requires OpenAI.")
            st.stop()

        results = {}
        progress = st.progress(0, text="Starting...")
        for i, uf in enumerate(uploaded_files, start=1):
            key = f"{i}::{uf.name}"
            try:
                progress.progress(
                    int(((i - 1) / len(uploaded_files)) * 100),
                    text=f"Processing {i}/{len(uploaded_files)}: {uf.name}",
                )
                results[key] = analyze_manuscript(uf, int(manuscript_year))
            except Exception as exc:
                results[key] = {
                    "filename": uf.name,
                    "error": _sanitize_error(str(exc)),
                    "manuscript_year": int(manuscript_year),
                }
        progress.empty()
        st.session_state["apa_batch_results"] = results
        st.rerun()

    results = st.session_state.get("apa_batch_results", {})
    if not results:
        return

    valid = {k: v for k, v in results.items() if not v.get("error")}
    for k, v in results.items():
        if v.get("error"):
            st.error(f"{v['filename']}: {v['error']}")

    if not valid:
        return

    selected_key = st.selectbox(
        "Select manuscript to review",
        list(valid.keys()),
        format_func=lambda k: valid[k].get("filename", k),
        key="apa_selected_manuscript",
    )
    result = valid[selected_key]

    stats = result["citation_stats"]
    recency = result["recency"]
    reference_rows = result.get("reference_comparison") or []
    citation_rows  = result.get("citation_comparison") or []
    matching       = result.get("matching_results") or []
    missing        = result.get("missing_references") or []

    total_refs = len(result["references"])
    total_cits = len(result["citations"])
    cited_count = sum(1 for row in matching if row.get("Cited"))
    cits_missing_ref = len(missing)
    refs_missing_cit = max(0, total_refs - cited_count)

    doi_checked = sum(1 for r in reference_rows if r.get("DOI Verified"))
    doi_sus = sum(1 for r in reference_rows if r.get("DOI Suspicious"))
    doi_sus_pct = doi_sus / total_refs * 100 if total_refs else 0
    high_conf = sum(1 for r in reference_rows if r.get("Confidence") == "high")
    med_conf = sum(1 for r in reference_rows if r.get("Confidence") == "medium")
    unv_conf = sum(1 for r in reference_rows if r.get("Confidence") == "unverified")

    st.caption(f"References: {total_refs}  |  In-text citations: {total_cits}")

    metric_rows = [
        {"Metric": "Total References", "Value": total_refs},
        {"Metric": "References > 15",
         "Value": (f"Yes ({total_refs})" if total_refs > 15 else f"No ({total_refs})")},
        {"Metric": "Total In-text Citations", "Value": total_cits},
        {"Metric": "Narrative Citations", "Value": stats.get("narrative", 0)},
        {"Metric": "Parenthetical Citations", "Value": stats.get("parenthetical", 0)},
        {"Metric": "HIGH-confidence corrections", "Value": high_conf},
        {"Metric": "MEDIUM-confidence corrections", "Value": med_conf},
        {"Metric": "UNVERIFIED corrections", "Value": unv_conf},
        {"Metric": "% Last 10 Years",
         "Value": f"{recency.get('recent_percentage', 0):.1f}%"},
        {"Metric": "Citations Missing from References", "Value": cits_missing_ref},
        {"Metric": "References Missing from Citations", "Value": refs_missing_cit},
        {"Metric": "DOI Checked (OpenAlex)", "Value": doi_checked},
        {"Metric": "DOI Suspicious (possible fabrication)",
         "Value": f"{doi_sus} ({doi_sus_pct:.1f}%)"},
    ]
    metric_df = pd.DataFrame(metric_rows)

    source_counts = Counter((r.get("Source Type") or "Other") for r in reference_rows)
    source_rows = []
    for stype in CANONICAL_SOURCE_TYPES:
        c = source_counts.get(stype, 0)
        pct = c / total_refs * 100 if total_refs else 0
        source_rows.append({
            "Source Type": stype,
            "Count / Percentage": f"{c} ({pct:.1f}%)",
        })
    source_df = pd.DataFrame(source_rows)

    col1, col2 = st.columns([1, 1], gap="large")
    with col1:
        st.dataframe(metric_df, use_container_width=True, hide_index=True)
    with col2:
        st.dataframe(source_df, use_container_width=True, hide_index=True)

    if st.toggle("Show In-text Citation Correction", value=False,
                 key="show_intext"):
        st.markdown("#### In-text Citation Correction")
        if citation_rows:
            df = pd.DataFrame([
                {
                    "Type": r.get("Type", ""),
                    "Original Citation": r.get("Original Citation", ""),
                    "Revised Citation": r.get("Revised Citation", ""),
                    "Status": r.get("Status", "MATCH"),
                    "Notes": r.get("Notes", ""),
                }
                for r in citation_rows
            ])
            st.dataframe(df, use_container_width=True, hide_index=True, height=320)
        else:
            st.info("No in-text citations available.")

    if st.toggle("Show Reference Correction", value=False,
                 key="show_reference"):
        st.markdown("#### Reference Correction")
        if reference_rows:
            df = pd.DataFrame([
                {
                    "No.": r.get("No."),
                    "Source Type": r.get("Source Type", "Other"),
                    "Year": r.get("Year"),
                    "Original Reference": r.get("Original Reference", ""),
                    "Corrected Version": (
                        "— WITHHELD (DOI mismatch) —"
                        if r.get("DOI Suspicious")
                        else r.get("Corrected Version", "")
                    ),
                    "Confidence": r.get("Confidence", "unverified"),
                    "Canonical Source": r.get("Canonical Source", "—"),
                    "Status": r.get("Status", "—"),
                    "Verification": r.get("Verification Status", "—"),
                    "OpenAlex Title": (r.get("OpenAlex Title") or "")[:60],
                    "OpenAlex Journal": r.get("OpenAlex Journal") or "—",
                    "OpenAlex Vol/Issue/Pages": " / ".join(
                        str(x) for x in [
                            r.get("OpenAlex Volume") or "—",
                            r.get("OpenAlex Issue") or "—",
                            r.get("OpenAlex Pages") or "—",
                        ]
                    ),
                }
                for r in reference_rows
            ])
            st.dataframe(df, use_container_width=True, hide_index=True, height=320)
        else:
            st.info("No references available.")

    try:
        docx_bytes = build_correction_docx(result)
        safe = re.sub(r"[^\w\-]+", "_", result.get("filename", "manuscript"))
        st.download_button(
            label="Download Diagnostic Report",
            data=docx_bytes,
            file_name=f"{safe}_diagnostic_report.docx",
            mime=("application/vnd.openxmlformats-officedocument."
                  "wordprocessingml.document"),
            use_container_width=True,
            key="download_apa_correction_docx",
        )
    except Exception as exc:
        st.error(f"Could not build DOCX: {exc}")
