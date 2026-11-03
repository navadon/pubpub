# CPE CMU Publication Dashboard — Spec

Static dashboard of publications by the Department of Computer Engineering,
Faculty of Engineering, Chiang Mai University. Same idea as the CivilChula
dashboard, but built on free public data. It runs on GitHub Pages and refreshes monthly.

## Decisions (from the grill step)

| Topic | Decision |
|---|---|
| Publication source | OpenAlex (free, no key), or a Scopus "Export -> CSV" file at `data/raw/scopus.csv` (used automatically when present; `--source openalex` overrides). See "Scopus CSV input" below. |
| Tiers | Both, switchable: **SJR** (SCImago Q1–Q4) and **CiteScore** (T1 = top 10% percentile, Q1 = 75–89, Q2 = 50–74, Q3 = 25–49, Q4 = 0–24) |
| CiteScore data | The user downloads the Scopus "CiteScore source title list" `.xlsx` to `data/raw/citescore.xlsx`. If the file is missing, the CiteScore toggle is disabled. |
| SJR data | SCImago journal-rank CSV (free download), cached in `data/raw/` |
| Staff | Everyone listed at https://cpe.eng.cmu.ac.th/lecturer-english.php, plus hand-added entries (`"manual": true`, e.g. Myo Thida) kept on roster refresh. Curated fields: `name_th`, `code`, `scopus_ids`, `scholar_url` |
| Years | Lifetime (whole careers, no affiliation filter); `years` runs from the first year with data. Raw OpenAlex works cached in `data/raw/cache/` (`--cached` reuses them) |
| Front end | One `index.html` file (plain JS, Chart.js from cdnjs) that reads `data/dashboard.json` |
| Update | GitHub Action on the 1st of each month: run pipeline → commit JSON → deploy Pages |
| Python | Must run on 3.9 (local) and 3.12 (CI). Standard library, plus `openpyxl` for the xlsx |

## Layout

```
index.html                 # the dashboard (single file)
data/staff.json            # roster + OpenAlex author IDs (hand-curated after matching)
data/dashboard.json        # pipeline output, consumed by index.html
data/raw/                  # SJR csv, citescore.xlsx (user-supplied), caches
scripts/fetch_staff.py     # scrape roster -> staff.json (keeps existing openalex_ids)
scripts/build_data.py      # OpenAlex + tiers -> dashboard.json
.github/workflows/update.yml
```

## Data contract

### data/staff.json
```json
[{ "name": "Navadon Khunlertgit", "title": "Assoc. Prof. Dr.", "position": "...",
   "profile_url": "...", "openalex_ids": ["A5012345678"], "scopus_ids": ["57200000001"], "match_note": "..." }]
```

`scopus_ids` is optional (strings of Scopus author IDs). `per_staff` entries also carry `scopus_ids`.

### data/dashboard.json
```json
{
  "generated_at": "2026-11-02",
  "source": "OpenAlex",                   // or "Scopus"
  "staff_count": 37,                      // everyone in staff.json, including people with no OpenAlex ID
  "staff_with_openalex_ids": 35,          // staff actually queried
  "staff_with_scopus_ids": 0,             // staff with a non-empty scopus_ids list
  "monthly_is_approximate": false,        // true when some works only have a year (Scopus): they sit in month 1, so the monthly trend is approximate
  "cmu_institution_id": "I48076826",
  "notes": ["..."],                       // list of strings: data caveats (e.g. SJR file/year fallback, CiteScore missing)
  "since_last_month": {"sjr": 2, "citescore": 1},   // citescore null if no CiteScore
  "years": [2017, ..., 2026],
  "current_year": 2026,
  "last_month": 10,                       // last complete month; 0 in January (current-year row is then all null)
  "has_citescore": true,
  "monthly_cumulative": { "2017": [n x12], ... },     // current year padded with null after last_month
  "year_totals": { "2017": 41, ... },
  "tiers": {                              // per year counts
    "sjr":       { "2024": {"Q1":..,"Q2":..,"Q3":..,"Q4":..,"No Ranking":..}, ... },
    "citescore": { "2024": {"T1":..,"Q1":..,"Q2":..,"Q3":..,"Q4":..,"No Ranking":..}, ... }
  },
  "collaboration": { "2026": {"International":..,"National":..,"CMU only":..}, ... },
  "per_staff": [ { "name": "...", "openalex_ids": ["A..."], "total": 12, "by_year": {"2026": 5, ...}, "top_tier_by_year": {"sjr": {...}, "citescore": {...}} } ],
  "publications": [ { "id": "W...", "title": "...", "year": 2026, "date": "2026-03-14",
                      "venue": "...", "issn": ["..."], "doi": "https://doi.org/10....", "type": "article",
                      "staff": ["..."], "sjr": "Q1", "citescore": "T1", "collab": "International" } ]
}
```

## Counting rules
- A work counts once for the department, even if several staff are co-authors (dedupe by OpenAlex work ID, then by DOI).
- Types included: article, review, conference paper (`proceedings-article`), book-chapter. Paratext, errata, retractions and preprints from repositories are excluded.
- Date = `publication_date`. The monthly trend is cumulative by that date.
- Collaboration: **International** if any author institution's country is not TH. **National** if every country is TH but some institution is not CMU. Otherwise **CMU only**.
- Tier lookup is by ISSN/eISSN of the primary source. Where a journal has several subject categories, the best one is used. No match gives "No Ranking". Conference papers usually get "No Ranking" under SJR unless the proceedings series is indexed.
- "Top tier" = Q1 under SJR, or T1+Q1 under CiteScore. The KPI "per professor" = top tier / staff_count.
- "Since last month" KPI = change in top-tier count compared with the previous monthly snapshot (`data/history.json`, appended by the pipeline).
- `since_last_month` = current-year top-tier count minus the `current_year` top-tier count in the most recent earlier-month snapshot of the same year. With no such snapshot it falls back to the number of current-year top-tier works dated in `last_month`.
- Invariant (asserted by the pipeline): for every year, the sums of `tiers.sjr[y]`, `tiers.citescore[y]` (when present) and `collaboration[y]` all equal `year_totals[y]`. Without CiteScore, `tiers.citescore` is `{}` and each publication's `citescore` is null.
- `year_totals`/tiers include the current (incomplete) month; `monthly_cumulative` stops at `last_month`.
- `history.json` entries: `{month, generated_at, current_year, total, current_year_total, top_tier: {sjr|citescore: {all, current_year} | null}}`; one per month (same-month rerun replaces).

## Scopus CSV input

`python3 scripts/build_data.py --scopus-csv data/raw/scopus.csv` (or just put the file there; `--source openalex` forces OpenAlex, `--source scopus` forces Scopus). No OpenAlex requests are made. Everything after the publication records (SJR/CiteScore tiers, collaboration, aggregation, history, invariants) is shared with the OpenAlex path, so `dashboard.json` has the same shape, with `source: "Scopus"`.

- Columns are matched by header name (case-insensitive, BOM-safe, extra columns ignored): Authors, Author full names, Author(s) ID, Title, Year, Source title, DOI, Link, Affiliations, Authors with affiliations, ISSN, Document Type, Publication Stage, EID, and optionally `Publication date` / `Cover date`.
- Types: Article -> `article`, Review -> `review`, Conference paper -> `proceedings-article`, Book chapter -> `book-chapter`; all others (Erratum, Editorial, Letter, Note...) are dropped. Publication Stage must be `Final` or empty.
- `id` = EID. `doi` = `https://doi.org/` + DOI (falls back to a doi.org `Link`). Dedupe by EID, then DOI. ISSN: 8 digits, hyphen optional, a lost leading zero (7 digits) is restored; several ISSNs may be separated by `;`.
- Date: parsed from the date column if present (`YYYY-MM-DD`, `YYYY-MM`, `May 2023`, ...). Otherwise `YYYY-01-01` from `Year`; any such work sets `monthly_is_approximate: true` and adds a note. Months are then approximate. The current-year row is still cut at `last_month`.
- Staff: a work belongs to the staff whose `scopus_ids` appear in `Author(s) ID`. Staff with no `scopus_ids` fall back to normalized surname + first initial against `Author full names`, only when `Authors with affiliations` ties that same author to Chiang Mai University. Every fallback match is logged to stderr.
- Collaboration from `Affiliations` (`;`-separated): **International** if any affiliation does not end with "Thailand"; **National** if all are Thai but some are not Chiang Mai University; otherwise **CMU only**. Empty Affiliations count as CMU only and add a note.
