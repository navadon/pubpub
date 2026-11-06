# PubPub: Publication Dashboard, Public Square

A static dashboard of publications by staff of Department of Computer Engineering,
Faculty of Engineering, Chiang Mai University. It shows yearly and cumulative counts,
journal quartiles (SJR and CiteScore), collaboration type, and per-staff totals.
It is a single `index.html` page that reads `data/dashboard.json`, hosted on GitHub Pages
and refreshed on the 1st of each month by a GitHub Action.

## Disclaimer and privacy

PubPub is an unofficial, independent dashboard, not an official service of Department of
Computer Engineering, Faculty of Engineering, or Chiang Mai University. Its figures are for
general information only and are not a basis for evaluation, promotion, or KPI reporting. The page
collects no visitor data: no accounts, cookies, analytics, or forms. GitHub Pages and cdnjs may log
basic technical data such as IP addresses. Staff names and publication records come from public
sources; to correct or remove a record, open an issue.

## Data sources and caveats

- **Publications: OpenAlex** (free). OpenAlex coverage differs from Scopus and SciVal,
  so counts will not match Scopus numbers exactly. Author matching relies on OpenAlex author IDs.
- **Keywords:** OpenAlex's per-paper `keywords`, shown in the Research keywords section.
- **SJR quartiles:** SCImago journal-rank CSV, cached in `data/raw/`.
- **CiteScore tiers (T1, Q1–Q4):** these need a file you download yourself. Get the
  Scopus "CiteScore source title list" `.xlsx` from scopus.com/sources and save it as
  `data/raw/citescore.xlsx`. If the file is missing, the CiteScore toggle is disabled.
- **Staff roster:** scraped from https://cpe.eng.cmu.ac.th/lecturer-english.php.

## Run locally

Requires Python 3 (3.9 or newer) and `openpyxl` (see `requirements.txt`).

```sh
python3 -m pip install -r requirements.txt
python3 scripts/fetch_staff.py     # refresh data/staff.json from the roster
python3 scripts/build_data.py      # build data/dashboard.json
python3 -m http.server             # then open http://localhost:8000
```

Set `OPENALEX_API_KEY` in the environment if you have one; the build works without it.

### Using a Scopus export instead of OpenAlex

1. In Scopus, search your department's documents, then **Export -> CSV** (include author IDs,
   affiliations, document type, publication stage, ISSN, EID; a date column is optional).
2. Save it as `data/raw/scopus.csv`.
3. Add `"scopus_ids": ["57200000001", ...]` to each person in `data/staff.json`. People without IDs
   are matched by surname + first initial only where Scopus lists Chiang Mai University as their
   affiliation; each such match is printed in the build log, so check it.
4. Run `python3 scripts/build_data.py`. If `data/raw/scopus.csv` exists it is used automatically and
   OpenAlex is not contacted. Use `--scopus-csv path/to/file.csv` for another location, or
   `--source openalex` to ignore the file.

Scopus exports have no exact publication date, so works are placed in January of their year and
`dashboard.json` gets `monthly_is_approximate: true`.

## Fixing an author match

Automatic matching can pick the wrong OpenAlex author or miss one. Edit `data/staff.json`
directly: each person has an `openalex_ids` list, so add or remove IDs there and add a
short note in `match_note`. `fetch_staff.py` keeps existing IDs when it refreshes the roster.
Then rerun `python3 scripts/build_data.py`.

## Enable GitHub Pages

1. In the repository, go to **Settings → Pages**.
2. Under **Build and deployment**, set **Source** to **GitHub Actions**.

After that, the `update.yml` workflow deploys the site on each push to `main` and on the
monthly schedule. You can also run it by hand from the **Actions** tab (workflow_dispatch).
