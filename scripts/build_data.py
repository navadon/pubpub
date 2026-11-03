#!/usr/bin/env python3
"""Build data/dashboard.json (and append to data/history.json) from OpenAlex or a Scopus CSV export.

Usage:
    python3 scripts/build_data.py [--staff data/staff.json] [--out data/dashboard.json]
                                  [--history data/history.json] [--raw-dir data/raw]
                                  [--start 2017-01-01] [--no-download] [--refresh-sjr]
                                  [--source auto|openalex|scopus] [--scopus-csv data/raw/scopus.csv]

Source: with --scopus-csv FILE the Scopus export is used. Otherwise, if <raw-dir>/scopus.csv exists
it is used automatically, unless --source openalex is given. No OpenAlex requests are made in
Scopus mode.

Environment (all optional):
    OPENALEX_MAILTO   contact e-mail sent to OpenAlex (polite pool)
    OPENALEX_API_KEY  OpenAlex API key

Python 3.9+; standard library plus openpyxl (only for data/raw/citescore.xlsx).
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

OPENALEX = "https://api.openalex.org"
SJR_URL = "https://www.scimagojr.com/journalrank.php"
USER_AGENT = "pubpub-dashboard/1.0 (CPE CMU publication dashboard; +https://github.com/)"
SJR_MAX_AGE_DAYS = 30
INCLUDED_TYPES = {"article", "review", "book-chapter"}
WORK_SELECT = ",".join([
    "id", "doi", "title", "publication_date", "publication_year", "type",
    "type_crossref", "is_retracted", "primary_location", "authorships",
])
ID_CHUNK = 40  # author IDs per OR filter

ROOT = Path(__file__).resolve().parent.parent
_blocked_hosts: Set[str] = set()
NOTES: List[str] = []


def log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


def note(msg: str) -> None:
    if msg not in NOTES:
        NOTES.append(msg)
    log("NOTE: " + msg)


# --------------------------------------------------------------------------- HTTP

# Retry-After above this is a spent daily budget, not a transient limit
MAX_RETRY_WAIT = 120
# Chiang Mai University (ROR 05m2fqn25); skips a costly search request
CMU_INSTITUTION_ID = "I48076826"


def http_get(url: str, retries: int = 5, timeout: int = 90) -> bytes:
    host = urllib.parse.urlparse(url).netloc
    if host in _blocked_hosts:
        raise RuntimeError("host previously refused access: " + host)
    delay = 2.0
    last: Optional[Exception] = None
    for attempt in range(retries):
        req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "*/*"})
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read()
        except urllib.error.HTTPError as e:
            last = e
            if e.code in (401, 403, 404):
                if e.code in (401, 403):
                    _blocked_hosts.add(host)
                raise RuntimeError("HTTP %d for %s" % (e.code, url))
            wait = delay
            ra = e.headers.get("Retry-After") if e.headers else None
            if ra and ra.isdigit():
                wait = max(wait, float(ra))
            if wait > MAX_RETRY_WAIT:
                raise RuntimeError(
                    "HTTP %d for %s; server asks to wait %.1f h. The OpenAlex daily budget is "
                    "probably used up: set OPENALEX_API_KEY or retry after midnight UTC."
                    % (e.code, url[:100], wait / 3600))
            log("HTTP %d, retry in %.0fs (%s)" % (e.code, wait, url[:100]))
        except (urllib.error.URLError, TimeoutError, ConnectionError) as e:
            last = e
            wait = delay
            log("network error %r, retry in %.0fs" % (e, wait))
        time.sleep(wait)
        delay = min(delay * 2, 60)
    raise RuntimeError("giving up on %s: %r" % (url, last))


def oa_get(path: str, params: Dict[str, str]) -> Dict[str, Any]:
    p = dict(params)
    if os.environ.get("OPENALEX_MAILTO"):
        p["mailto"] = os.environ["OPENALEX_MAILTO"]
    if os.environ.get("OPENALEX_API_KEY"):
        p["api_key"] = os.environ["OPENALEX_API_KEY"]
    url = OPENALEX + path + "?" + urllib.parse.urlencode(p, safe=":|,")
    data = json.loads(http_get(url).decode("utf-8"))
    time.sleep(0.15)
    return data


def short_id(openalex_url: Optional[str]) -> str:
    return (openalex_url or "").rstrip("/").rsplit("/", 1)[-1]


# --------------------------------------------------------------------------- CMU institution

def find_cmu_institution(raw_dir: Path) -> str:
    cache = raw_dir / "cmu_institution.json"
    if cache.exists():
        try:
            return json.loads(cache.read_text())["id"]
        except Exception:
            pass
    if CMU_INSTITUTION_ID:
        return CMU_INSTITUTION_ID
    data = oa_get("/institutions", {"search": "Chiang Mai University",
                                    "select": "id,display_name,country_code,ror",
                                    "per-page": "10"})
    for r in data.get("results", []):
        if r.get("display_name") == "Chiang Mai University" and r.get("country_code") == "TH":
            cid = short_id(r["id"])
            cache.write_text(json.dumps({"id": cid, "display_name": r["display_name"],
                                         "ror": r.get("ror")}, indent=1))
            return cid
    raise RuntimeError("Could not find Chiang Mai University in OpenAlex institutions")


# --------------------------------------------------------------------------- SJR

def norm_issn(s: Optional[str]) -> str:
    return re.sub(r"[^0-9Xx]", "", s or "").upper()


def fresh(path: Path) -> bool:
    return path.exists() and (time.time() - path.stat().st_mtime) < SJR_MAX_AGE_DAYS * 86400


def fetch_sjr_file(path: Path, year: Optional[int], allow_download: bool, force: bool) -> bool:
    """Make sure `path` holds an SJR CSV. Returns True if usable file exists."""
    if path.exists() and (not allow_download or (fresh(path) and not force)):
        return True
    if allow_download and urllib.parse.urlparse(SJR_URL).netloc not in _blocked_hosts:
        url = SJR_URL + "?" + ("year=%d&" % year if year else "") + "out=xls"
        try:
            log("Downloading " + url)
            body = http_get(url, retries=3)
            if b";" not in body[:2000] or b"<html" in body[:500].lower():
                raise RuntimeError("response is not a SCImago CSV")
            path.write_bytes(body)
            return True
        except Exception as e:
            log("SJR download failed: %s" % e)
    return path.exists()


def parse_sjr(path: Path) -> Dict[str, str]:
    text = path.read_bytes().decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text), delimiter=";")
    out: Dict[str, str] = {}
    rank = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}
    for row in reader:
        q = (row.get("SJR Best Quartile") or "").strip()
        if q not in rank:
            continue
        for raw in re.split(r"[,\s]+", row.get("Issn") or ""):
            k = norm_issn(raw)
            if len(k) == 8 and (k not in out or rank[q] < rank[out[k]]):
                out[k] = q
        t = "t:" + norm_title(row.get("Title") or "")
        if len(t) > 2 and (t not in out or rank[q] < rank[out[t]]):
            out[t] = q
    return out


def norm_title(s: str) -> str:
    """Journal title key for exports without ISSNs: 'IEEE Access' -> 'ieeeaccess'."""
    return re.sub(r"[^a-z0-9]", "", (s or "").lower().replace("&", "and"))


class SjrLookup:
    def __init__(self, raw_dir: Path, years: List[int], allow_download: bool, refresh: bool):
        self.by_year: Dict[int, Dict[str, str]] = {}
        self.latest: Dict[str, str] = {}
        self.has_latest = False
        self.year_files: Dict[int, str] = {}
        latest_path = raw_dir / "sjr_latest.csv"
        if fetch_sjr_file(latest_path, None, allow_download, refresh):
            try:
                self.latest = parse_sjr(latest_path)
                self.has_latest = bool(self.latest)
            except Exception as e:
                log("cannot parse %s: %s" % (latest_path, e))
        # year-specific files (cached or downloadable); skip downloads if the host refused us
        for y in years:
            p = raw_dir / ("sjr_%d.csv" % y)
            if fetch_sjr_file(p, y, allow_download, refresh):
                try:
                    d = parse_sjr(p)
                    if d:
                        self.by_year[y] = d
                except Exception as e:
                    log("cannot parse %s: %s" % (p, e))
        if not self.by_year and not self.has_latest:
            note("SJR data unavailable (download failed and no file in data/raw/); "
                 "all SJR tiers are 'No Ranking'. Place a SCImago CSV at data/raw/sjr_latest.csv.")
        elif not self.by_year:
            note("SJR: year-specific files unavailable; the latest SCImago file (sjr_latest.csv) "
                 "was used for all publication years.")
        else:
            missing = [y for y in years if y not in self.by_year]
            note("SJR: year-specific files used for %s%s." % (
                ", ".join(str(y) for y in sorted(self.by_year)),
                "; latest file used for " + ", ".join(str(y) for y in missing) if missing and self.has_latest
                else ("; no data for " + ", ".join(str(y) for y in missing) if missing else "")))

    def tier_by_title(self, title: str, year: int) -> str:
        k = "t:" + norm_title(title)
        return (self.by_year.get(year) or {}).get(k) or self.latest.get(k) or "No Ranking"

    def tier(self, issns: List[str], year: int) -> str:
        table = self.by_year.get(year) or self.latest
        if not table:
            return "No Ranking"
        best = None
        for i in issns:
            q = table.get(norm_issn(i))
            if q and (best is None or q < best):
                best = q
        # fall back to latest file if year-specific file lacks the journal
        if best is None and table is not self.latest:
            for i in issns:
                q = self.latest.get(norm_issn(i))
                if q and (best is None or q < best):
                    best = q
        return best or "No Ranking"


# --------------------------------------------------------------------------- CiteScore

def percentile_to_tier(p: float) -> str:
    if p >= 90:
        return "T1"
    if p >= 75:
        return "Q1"
    if p >= 50:
        return "Q2"
    if p >= 25:
        return "Q3"
    return "Q4"


class CiteScoreLookup:
    def __init__(self, path: Path):
        self.available = False
        self.data: Dict[str, Dict[int, float]] = {}  # issn -> {year (0 = unknown): best percentile}
        if not path.exists():
            return
        try:
            import openpyxl  # type: ignore
        except ImportError:
            note("openpyxl not installed; CiteScore skipped")
            return
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        rows_used = 0
        for ws in wb.worksheets:
            header_idx = None
            cols: Dict[str, Any] = {}
            for n, row in enumerate(ws.iter_rows(values_only=True)):
                if header_idx is None:
                    if n > 15:
                        break
                    names = [str(c).strip().lower() if c is not None else "" for c in row]
                    issn_c = [i for i, h in enumerate(names) if h in ("issn", "print issn", "print-issn")]
                    e_c = [i for i, h in enumerate(names) if h.replace("-", "").replace(" ", "") in ("eissn",)]
                    pct = [i for i, h in enumerate(names) if h == "percentile" or "citescore percentile" in h]
                    yr = [i for i, h in enumerate(names) if h in ("year", "citescore year")]
                    if (issn_c or e_c) and pct:
                        header_idx = n
                        cols = {"issn": issn_c, "e": e_c, "pct": pct[0], "yr": yr[0] if yr else None}
                    continue
                try:
                    p = row[cols["pct"]]
                    if p is None or p == "":
                        continue
                    p = float(p)
                except (ValueError, TypeError, IndexError):
                    continue
                y = 0
                if cols["yr"] is not None:
                    try:
                        y = int(row[cols["yr"]])
                    except (ValueError, TypeError, IndexError):
                        y = 0
                for i in cols["issn"] + cols["e"]:
                    if i >= len(row) or row[i] is None:
                        continue
                    k = norm_issn(str(row[i]))
                    if len(k) == 7:
                        k = "0" + k  # leading zero lost in Excel
                    if len(k) != 8:
                        continue
                    d = self.data.setdefault(k, {})
                    if p > d.get(y, -1):
                        d[y] = p
                rows_used += 1
        wb.close()
        self.available = bool(self.data)
        if self.available:
            note("CiteScore: %d ISSNs loaded from %s (%d rows)." % (len(self.data), path.name, rows_used))
        else:
            note("citescore.xlsx found but no usable ISSN/Percentile columns; CiteScore disabled.")

    def tier(self, issns: List[str], year: int) -> str:
        best = None
        for i in issns:
            d = self.data.get(norm_issn(i))
            if not d:
                continue
            if year in d:
                p = d[year]
            elif 0 in d:
                p = d[0]
            else:
                ys = sorted(d)
                le = [y for y in ys if y <= year]
                p = d[le[-1] if le else ys[0]]
            if best is None or p > best:
                best = p
        return percentile_to_tier(best) if best is not None else "No Ranking"


# --------------------------------------------------------------------------- OpenAlex works

def fetch_works(author_ids: List[str], start: str, end: str) -> List[Dict[str, Any]]:
    works: Dict[str, Dict[str, Any]] = {}
    for i in range(0, len(author_ids), ID_CHUNK):
        chunk = author_ids[i:i + ID_CHUNK]
        flt = "authorships.author.id:%s,from_publication_date:%s,to_publication_date:%s" % (
            "|".join(chunk), start, end)
        # cursor paging past page 1 needs an API key; plain page= paging does not
        use_cursor = bool(os.environ.get("OPENALEX_API_KEY"))
        cursor, page, seen, count = "*", 1, 0, None
        while True:
            params = {"filter": flt, "select": WORK_SELECT, "per-page": "200"}
            if use_cursor:
                params["cursor"] = cursor
            else:
                params["page"] = str(page)
            data = oa_get("/works", params)
            results = data.get("results", [])
            for w in results:
                works[w["id"]] = w
            seen += len(results)
            count = data.get("meta", {}).get("count", count)
            log("  page %d: %d/%s works" % (page, seen, count))
            cursor = data.get("meta", {}).get("next_cursor")
            page += 1
            if not results or (use_cursor and not cursor) or (count is not None and seen >= count):
                break
        log("fetched %d works so far (author chunk %d)" % (len(works), i // ID_CHUNK + 1))
    return list(works.values())


def norm_doi(d: Optional[str]) -> str:
    if not d:
        return ""
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", d.strip().lower())


def work_kind(w: Dict[str, Any]) -> Optional[str]:
    """Return the included publication type or None if excluded."""
    t = w.get("type")
    tc = w.get("type_crossref")
    if w.get("is_retracted"):
        return None
    loc = w.get("primary_location") or {}
    src = loc.get("source") or {}
    if src.get("type") == "repository":
        return None  # preprints/repository copies
    if tc == "proceedings-article" or t == "conference-paper":  # OpenAlex's newer type name
        return "proceedings-article"
    if t in INCLUDED_TYPES:
        if src.get("type") == "conference" and t == "article":
            return "proceedings-article"
        return t
    return None


def collab_class(w: Dict[str, Any], cmu_id: str) -> str:
    countries: Set[str] = set()
    non_cmu_inst = False
    for a in w.get("authorships") or []:
        for inst in a.get("institutions") or []:
            iid = short_id(inst.get("id"))
            lineage = [short_id(x) for x in inst.get("lineage") or []]
            cc = inst.get("country_code")
            if cc:
                countries.add(cc)
            if iid != cmu_id and cmu_id not in lineage:
                non_cmu_inst = True
    if any(c != "TH" for c in countries):
        return "International"
    if non_cmu_inst:
        return "National"
    return "CMU only"


# --------------------------------------------------------------------------- Scopus CSV

SCOPUS_TYPES = {
    "article": "article",
    "review": "review",
    "conference paper": "proceedings-article",
    "book chapter": "book-chapter",
}
DATE_COLUMNS = ("publication date", "cover date", "coverdate", "cover_date", "publicationdate")
DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%d %B %Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y",
                "%B %Y", "%b %Y", "%Y-%m", "%Y %B", "%Y %b")


def norm_name(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    return re.sub(r"[^a-z]", "", s.encode("ascii", "ignore").decode().lower())


def scopus_year_month(raw: str) -> Optional[dt.date]:
    raw = (raw or "").strip()
    if not raw:
        return None
    for f in DATE_FORMATS:
        try:
            return dt.datetime.strptime(raw, f).date()
        except ValueError:
            pass
    return None


def scopus_issns(raw: str) -> List[str]:
    out = []
    for tok in re.split(r"[;,\s]+", raw or ""):
        k = norm_issn(tok)
        if len(k) == 7:
            k = "0" + k  # leading zero lost (Excel)
        if len(k) == 8 and k not in out:
            out.append(k)
    return out


def scopus_collab(affiliations: str) -> Optional[str]:
    affs = [a.strip() for a in (affiliations or "").split(";") if a.strip()]
    if not affs:
        return None
    if any(not re.search(r"thailand\s*\.?$", a, re.I) for a in affs):
        return "International"
    if any(not re.search(r"chiang\s+mai\s+univ", a, re.I) for a in affs):
        return "National"
    return "CMU only"


def split_name_initials(name: str):
    """'Khunlertgit N.' -> ('khunlertgit', 'n'); 'Na Nakorn K.-J.' -> ('nanakorn', 'k')."""
    toks = name.strip().split()
    if len(toks) >= 2 and re.fullmatch(r"(?:[A-Z][a-z]?\.?-?)+", toks[-1]) and toks[-1].endswith("."):
        return norm_name("".join(toks[:-1])), norm_name(toks[-1])[:1]
    return norm_name(name), ""


def cmu_author_keys(authors_with_aff: str):
    """(surname, initial) of authors whose affiliation text mentions Chiang Mai University."""
    keys = set()
    for seg in (authors_with_aff or "").split(";"):
        if re.search(r"chiang\s+mai\s+univ", seg, re.I):
            keys.add(split_name_initials(seg.split(",")[0]))
    return keys


def staff_name_keys(name: str):
    toks = [t for t in re.split(r"\s+", name.strip()) if norm_name(t)]
    if len(toks) < 2:
        return set()
    ini = norm_name(toks[0])[:1]
    return {(norm_name(toks[-1]), ini), (norm_name("".join(toks[1:])), ini)}


def read_scopus_rows(path: Path) -> List[Dict[str, str]]:
    text = path.read_bytes().decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    header = None
    rows: List[Dict[str, str]] = []
    for r in reader:
        if header is None:
            header = [h.strip().lstrip("﻿").strip().lower() for h in r]
            continue
        if not any(c.strip() for c in r):
            continue
        rows.append({h: (r[i].strip() if i < len(r) else "") for i, h in enumerate(header)})
    if not rows or "eid" not in rows[0] or "title" not in rows[0]:
        raise RuntimeError("%s does not look like a Scopus CSV export (need EID and Title columns)" % path)
    return rows


def scopus_pubs(path: Path, staff: List[Dict[str, Any]], start: str, today: dt.date,
                sjr: Any, cs: Any) -> List[Dict[str, Any]]:
    rows = read_scopus_rows(path)
    log("Scopus CSV: %d rows" % len(rows))
    id_to_name = {i: s["name"] for s in staff for i in s["scopus_ids"]}
    fallback = [(s["name"], staff_name_keys(s["name"])) for s in staff if not s["scopus_ids"]]
    if fallback:
        note("Staff without scopus_ids (%d) are matched by surname + given name (initial if that is all Scopus has), only when the "
             "author's affiliation is Chiang Mai University; each fallback match is logged." % len(fallback))
    date_col = next((c for c in DATE_COLUMNS if c in rows[0]), None)
    approx = 0
    no_aff = 0
    pubs: List[Dict[str, Any]] = []
    by_doi: Dict[str, Dict[str, Any]] = {}
    seen_eid: Set[str] = set()
    skipped: Dict[str, int] = {}
    for r in sorted(rows, key=lambda x: x.get("eid", "")):
        eid = r.get("eid", "")
        dtype = r.get("document type", "").lower()
        kind = SCOPUS_TYPES.get(dtype)
        stage = r.get("publication stage", "")
        if not eid or eid in seen_eid:
            continue
        if not kind or (stage and stage.lower() != "final"):
            k = dtype if not kind else "stage " + stage
            skipped[k] = skipped.get(k, 0) + 1
            continue
        seen_eid.add(eid)
        # --- staff
        names: Set[str] = set()
        for i in re.split(r"[;\s]+", r.get("author(s) id", "")):
            if i in id_to_name:
                names.add(id_to_name[i])
        if fallback:
            cmu_keys = cmu_author_keys(r.get("authors with affiliations", ""))
            for full in r.get("author full names", "").split(";"):
                m = re.match(r"^\s*([^,(]+),\s*([^(]*?)\s*(?:\(\d+\))?\s*$", full)
                if not m or not norm_name(m.group(2)):
                    continue
                key = (norm_name(m.group(1)), norm_name(m.group(2))[:1])
                if key not in cmu_keys:
                    continue
                given = norm_name(m.group(2))
                for sname, skeys in fallback:
                    # full given name must agree when Scopus has one (Sakgasit != Sakgasem Ramingwong)
                    first = norm_name(sname.split()[0])
                    if len(given) > 2 and not (first.startswith(given) or given.startswith(first)):
                        continue
                    if key in skeys and sname not in names:
                        names.add(sname)
                        log("FALLBACK name match: %s <- '%s' in %s (%s)" % (
                            sname, full.strip(), eid, r.get("title", "")[:60]))
        if not names:
            continue
        # --- date
        d = scopus_year_month(r.get(date_col, "")) if date_col else None
        is_approx = False
        if d is None:
            try:
                y = int(r.get("year", ""))
            except ValueError:
                continue
            d = dt.date(y, 1, 1)
            is_approx = True
        date = d.isoformat()
        if date < start or date > today.isoformat():
            continue
        doi = norm_doi(r.get("doi", ""))
        if not doi:
            m = re.search(r"doi\.org/(10\..+)$", r.get("link", ""), re.I)
            doi = m.group(1).lower() if m else ""
        if doi and doi in by_doi:
            prev = by_doi[doi]
            prev["staff"] = sorted(set(prev["staff"]) | names)
            continue
        issn = scopus_issns(r.get("issn", ""))
        collab = scopus_collab(r.get("affiliations", ""))
        if collab is None:
            no_aff += 1
            collab = "CMU only"
        p = {
            "id": eid, "title": r.get("title", ""), "year": d.year, "date": date,
            "venue": r.get("source title", ""), "issn": issn,
            "doi": ("https://doi.org/" + doi) if doi else "", "type": kind, "staff": sorted(names),
            "sjr": sjr.tier(issn, d.year) if issn else sjr.tier_by_title(r.get("source title", ""), d.year),
            "citescore": cs.tier(issn, d.year) if cs.available else None,
            "collab": collab,
        }
        pubs.append(p)
        approx += 1 if is_approx else 0
        if doi:
            by_doi[doi] = p
    if skipped:
        log("skipped by type/stage: " + ", ".join("%s=%d" % kv for kv in sorted(skipped.items())))
    if approx:
        note("Scopus export has no exact publication date for %d works; they are placed in "
             "January of their year, so the monthly trend is approximate." % approx)
    if no_aff:
        note("%d Scopus works have no Affiliations value and were classed 'CMU only'." % no_aff)
    pubs.sort(key=lambda p: (p["date"], p["id"]), reverse=True)
    # approximate = any work lacked a real date
    scopus_pubs.approx = approx > 0  # type: ignore[attr-defined]
    return pubs


# --------------------------------------------------------------------------- build

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--staff", default=str(ROOT / "data/staff.json"))
    ap.add_argument("--out", default=str(ROOT / "data/dashboard.json"))
    ap.add_argument("--history", default=str(ROOT / "data/history.json"))
    ap.add_argument("--raw-dir", default=str(ROOT / "data/raw"))
    ap.add_argument("--start", default="1950-01-01",
                    help="earliest publication date; the default covers whole careers (years trimmed to data)")
    ap.add_argument("--cached", action="store_true",
                    help="reuse data/raw/cache/openalex_works.json instead of calling OpenAlex")
    ap.add_argument("--no-download", action="store_true", help="use only cached SJR files")
    ap.add_argument("--refresh-sjr", action="store_true", help="re-download SJR files regardless of age")
    ap.add_argument("--source", choices=["auto", "openalex", "scopus"], default="auto")
    ap.add_argument("--scopus-csv", default=None, help="Scopus Export -> CSV file (default <raw-dir>/scopus.csv)")
    args = ap.parse_args()

    t0 = time.time()
    today = dt.date.today()
    raw_dir = Path(args.raw_dir)
    raw_dir.mkdir(parents=True, exist_ok=True)
    # SCImago starts in 1999; `years` is trimmed to the first year with data after filtering
    years = list(range(max(int(args.start[:4]), 1999), today.year + 1))
    last_month = today.month - 1  # last complete month (0 in January)

    default_csv = raw_dir / "scopus.csv"
    scopus_path = Path(args.scopus_csv) if args.scopus_csv else default_csv
    if args.source == "scopus" or (args.source == "auto" and (args.scopus_csv or default_csv.exists())):
        use_scopus = True
        if not scopus_path.exists():
            log("Scopus CSV not found: %s" % scopus_path)
            return 1
    else:
        use_scopus = False
    source_name = "Scopus" if use_scopus else "OpenAlex"
    log("Source: " + source_name)

    staff_all = json.loads(Path(args.staff).read_text(encoding="utf-8"))
    staff = []
    for s in staff_all:
        ids = [short_id(x) for x in s.get("openalex_ids") or [] if x]
        sids = [re.sub(r"\D", "", str(x)) for x in s.get("scopus_ids") or [] if re.sub(r"\D", "", str(x))]
        if use_scopus:
            staff.append({"name": s["name"], "ids": ids, "scopus_ids": sids})
            continue
        if not ids:
            log("WARNING: no OpenAlex ID for %s, skipped" % s.get("name"))
            continue
        staff.append({"name": s["name"], "ids": ids, "scopus_ids": sids})
    if not staff:
        log("No staff with OpenAlex IDs; nothing to do")
        return 1
    id_to_name = {i: s["name"] for s in staff for i in s["ids"]}

    if use_scopus:
        cmu_id = CMU_INSTITUTION_ID
        raw_works = []  # type: List[Dict[str, Any]]
    else:
        cmu_id = find_cmu_institution(raw_dir)
        log("CMU institution id: " + cmu_id)
        cache = raw_dir / "cache" / "openalex_works.json"
        if args.cached and cache.exists():
            raw_works = json.loads(cache.read_text(encoding="utf-8"))
            log("using cached works from " + str(cache))
        else:
            raw_works = fetch_works(sorted(id_to_name), args.start, today.isoformat())
            cache.parent.mkdir(parents=True, exist_ok=True)
            cache.write_text(json.dumps(raw_works), encoding="utf-8")
        log("raw works: %d" % len(raw_works))

    sjr = SjrLookup(raw_dir, years, not args.no_download, args.refresh_sjr)
    cs = CiteScoreLookup(raw_dir / "citescore.xlsx")
    if not cs.available:
        note("CiteScore file data/raw/citescore.xlsx missing or unusable; CiteScore disabled.")

    # -- dedupe by work id (already unique), then by DOI
    pubs: List[Dict[str, Any]] = []
    by_doi: Dict[str, Dict[str, Any]] = {}
    monthly_is_approximate = False
    if use_scopus:
        pubs = scopus_pubs(scopus_path, staff, args.start, today, sjr, cs)
        monthly_is_approximate = bool(getattr(scopus_pubs, "approx", False))
    for w in ([] if use_scopus else sorted(raw_works, key=lambda x: x["id"])):
        kind = work_kind(w)
        date = w.get("publication_date")
        if not kind or not date or date < args.start or date > today.isoformat():
            continue
        names = sorted({id_to_name[short_id(a["author"]["id"])]
                        for a in w.get("authorships") or []
                        if a.get("author") and short_id(a["author"].get("id")) in id_to_name})
        if not names:
            continue
        doi = norm_doi(w.get("doi"))
        if doi and doi in by_doi:
            prev = by_doi[doi]
            prev["staff"] = sorted(set(prev["staff"]) | set(names))
            continue
        src = (w.get("primary_location") or {}).get("source") or {}
        issn = list(src.get("issn") or [])
        if src.get("issn_l") and src["issn_l"] not in issn:
            issn.append(src["issn_l"])
        year = int(date[:4])
        p = {
            "id": short_id(w["id"]), "title": w.get("title") or "", "year": year, "date": date,
            "venue": src.get("display_name") or "", "issn": issn,
            "doi": ("https://doi.org/" + doi) if doi else "", "type": kind, "staff": names,
            "sjr": sjr.tier(issn, year),
            "citescore": cs.tier(issn, year) if cs.available else None,
            "collab": collab_class(w, cmu_id),
        }
        pubs.append(p)
        if doi:
            by_doi[doi] = p
    if not use_scopus:
        pubs.sort(key=lambda p: (p["date"], p["id"]), reverse=True)
    log("publications after filtering/dedupe: %d" % len(pubs))
    if pubs:
        years = list(range(min(p["year"] for p in pubs), today.year + 1))

    # -- aggregates
    ys = [str(y) for y in years]
    monthly = {y: [0] * 12 for y in ys}
    for p in pubs:
        monthly[str(p["year"])][int(p["date"][5:7]) - 1] += 1
    cumulative: Dict[str, List[Optional[int]]] = {}
    for y in ys:
        run, row = 0, []  # type: int, List[Optional[int]]
        for m in range(12):
            run += monthly[y][m]
            row.append(run)
        if int(y) == today.year:
            row = [v if m < last_month else None for m, v in enumerate(row)]
        cumulative[y] = row
    year_totals = {y: sum(monthly[y]) for y in ys}

    sjr_keys = ["Q1", "Q2", "Q3", "Q4", "No Ranking"]
    cs_keys = ["T1", "Q1", "Q2", "Q3", "Q4", "No Ranking"]
    tiers: Dict[str, Any] = {
        "sjr": {y: {k: 0 for k in sjr_keys} for y in ys},
        "citescore": ({y: {k: 0 for k in cs_keys} for y in ys} if cs.available else {}),
    }
    collab = {y: {"International": 0, "National": 0, "CMU only": 0} for y in ys}
    per_staff = {s["name"]: {"name": s["name"], "openalex_ids": s["ids"], "scopus_ids": s["scopus_ids"], "total": 0, "by_year": {}, "top_tier_by_year": {"sjr": {}, "citescore": {}}}
                 for s in staff}
    for p in pubs:
        y = str(p["year"])
        tiers["sjr"][y][p["sjr"]] += 1
        if cs.available:
            tiers["citescore"][y][p["citescore"]] += 1
        collab[y][p["collab"]] += 1
        for n in p["staff"]:
            ps = per_staff[n]
            ps["total"] += 1
            ps["by_year"][y] = ps["by_year"].get(y, 0) + 1
            if p["sjr"] == "Q1":
                ps["top_tier_by_year"]["sjr"][y] = ps["top_tier_by_year"]["sjr"].get(y, 0) + 1
            if cs.available and p["citescore"] in ("T1", "Q1"):
                ps["top_tier_by_year"]["citescore"][y] = ps["top_tier_by_year"]["citescore"].get(y, 0) + 1

    for y in ys:
        assert sum(tiers["sjr"][y].values()) == year_totals[y], "sjr tier sum mismatch " + y
        if cs.available:
            assert sum(tiers["citescore"][y].values()) == year_totals[y], "citescore tier sum mismatch " + y
        assert sum(collab[y].values()) == year_totals[y], "collaboration sum mismatch " + y

    cy = str(today.year)
    cur_top = {"sjr": tiers["sjr"][cy]["Q1"],
               "citescore": (tiers["citescore"][cy]["T1"] + tiers["citescore"][cy]["Q1"]) if cs.available else None}
    hist_path = Path(args.history)
    try:
        hist = json.loads(hist_path.read_text(encoding="utf-8")) if hist_path.exists() else []
    except ValueError:
        hist = []
    month_key = today.strftime("%Y-%m")
    prev = [h for h in hist if h.get("month", "") < month_key and h.get("current_year") == today.year]
    prev_snap = prev[-1] if prev else None

    def last_month_top(tier_name: str) -> int:
        n = 0
        for p in pubs:
            if p["year"] == today.year and int(p["date"][5:7]) == last_month:
                if (tier_name == "sjr" and p["sjr"] == "Q1") or \
                        (tier_name == "citescore" and p["citescore"] in ("T1", "Q1")):
                    n += 1
        return n

    since_last_month: Dict[str, Optional[int]] = {}
    for t in ("sjr", "citescore"):
        if cur_top[t] is None:
            since_last_month[t] = None
            continue
        pv = ((prev_snap or {}).get("top_tier") or {}).get(t)
        if pv is not None:
            since_last_month[t] = cur_top[t] - pv["current_year"]
        else:
            since_last_month[t] = last_month_top(t)

    out = {
        "generated_at": today.isoformat(),
        "source": source_name,
        "staff_count": len(staff_all),
        "staff_with_openalex_ids": sum(1 for s in staff if s["ids"]),
        "staff_with_scopus_ids": sum(1 for s in staff if s["scopus_ids"]),
        "monthly_is_approximate": monthly_is_approximate,
        "years": years,
        "current_year": today.year,
        "last_month": last_month,
        "has_citescore": cs.available,
        "cmu_institution_id": cmu_id,
        "since_last_month": since_last_month,
        "notes": NOTES,
        "monthly_cumulative": cumulative,
        "year_totals": year_totals,
        "tiers": tiers,
        "collaboration": collab,
        "per_staff": sorted(per_staff.values(), key=lambda s: -sum(s["by_year"].values())),
        "publications": pubs,
    }
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(out, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")

    # -- history snapshot (one per month; re-running in the same month replaces it)
    snap = {
        "month": today.strftime("%Y-%m"),
        "generated_at": today.isoformat(),
        "current_year": today.year,
        "total": len(pubs),
        "current_year_total": year_totals[cy],
        "top_tier": {
            "sjr": {"all": sum(tiers["sjr"][y]["Q1"] for y in ys), "current_year": tiers["sjr"][cy]["Q1"]},
            "citescore": ({"all": sum(tiers["citescore"][y]["T1"] + tiers["citescore"][y]["Q1"] for y in ys),
                           "current_year": tiers["citescore"][cy]["T1"] + tiers["citescore"][cy]["Q1"]}
                          if cs.available else None),
        },
    }
    hist = [h for h in hist if h.get("month") != snap["month"]] + [snap]
    hist.sort(key=lambda h: h["month"])
    hist_path.write_text(json.dumps(hist, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")

    log("wrote %s (%d publications, %d staff) in %.0fs" % (out_path, len(pubs), len(staff), time.time() - t0))
    return 0


if __name__ == "__main__":
    sys.exit(main())
