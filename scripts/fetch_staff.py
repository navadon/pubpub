#!/usr/bin/env python3
"""Scrape the CPE CMU academic staff list into data/staff.json.

Standard library only (Python 3.9+). Keeps existing `openalex_ids` and
`match_note` values from data/staff.json for staff whose name matches.
"""

import json
import os
import re
import sys
import unicodedata
import urllib.request
from html.parser import HTMLParser
from urllib.parse import urljoin

SOURCE_URL = "https://cpe.eng.cmu.ac.th/lecturer-english.php"
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_PATH = os.path.join(PROJECT_ROOT, "data", "staff.json")

# Degree abbreviations that appear before the academic rank, e.g. "Ph.D. Assistant Professor".
DEGREE_RE = re.compile(r"^(?:Ph\.\s*D\.|D\.\s*Eng\.|Dr\.)\s*", re.IGNORECASE)


def collapse(text):
    return re.sub(r"\s+", " ", text or "").strip()


class StaffCardParser(HTMLParser):
    """Collects one record per <div class="single-news"> card.

    Within a card, the first <h2> holds the name (and link), the second <h2>
    holds the role/position in a coloured <font>, and later <h2>s hold room
    and e-mail, which are ignored.
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.cards = []
        self._div_depth = 0
        self._card_depth = None  # div depth at which the current card opened
        self._card = None
        self._in_h2 = False
        self._h2_index = -1  # index of the h2 inside the current card
        self._h2_text = []
        self._h2_href = None
        self._image_href = None  # fallback profile link from the photo anchor

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "div":
            self._div_depth += 1
            classes = (attrs.get("class") or "").split()
            if "single-news" in classes and self._card is None:
                self._card = {"name": None, "position": None, "profile_href": None}
                self._card_depth = self._div_depth
                self._h2_index = -1
                self._image_href = None
        elif self._card is not None:
            if tag == "a" and attrs.get("href") and "view_id=" in attrs["href"]:
                if self._image_href is None:
                    self._image_href = attrs["href"]
                if self._in_h2 and self._h2_href is None:
                    self._h2_href = attrs["href"]
            elif tag == "h2":
                self._h2_index += 1
                self._in_h2 = True
                self._h2_text = []
                self._h2_href = None

    def handle_endtag(self, tag):
        if tag == "h2" and self._card is not None and self._in_h2:
            self._in_h2 = False
            text = collapse("".join(self._h2_text))
            if self._h2_index == 0:
                self._card["name"] = text
                self._card["profile_href"] = self._h2_href
            elif self._h2_index == 1:
                self._card["position"] = text or None
        elif tag == "div":
            if self._card is not None and self._div_depth == self._card_depth:
                if self._card["profile_href"] is None:
                    self._card["profile_href"] = self._image_href
                self.cards.append(self._card)
                self._card = None
                self._card_depth = None
            self._div_depth -= 1

    def handle_data(self, data):
        if self._in_h2:
            self._h2_text.append(data)


def split_name_and_title(raw):
    """'Navadon Khunlertgit, Ph.D.  Assistant Professor' -> (name, rank or None)."""
    if "," in raw:
        name, rest = raw.split(",", 1)
    else:
        name, rest = raw, ""
    rest = collapse(rest)
    # Strip degree abbreviations so only the academic rank remains.
    while True:
        stripped = DEGREE_RE.sub("", rest, count=1)
        if stripped == rest:
            break
        rest = stripped.strip()
    return collapse(name), (rest or None)


def fetch_html(url):
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (cpepubs staff fetcher)"})
    with urllib.request.urlopen(request, timeout=30) as response:
        raw = response.read()
        charset = response.headers.get_content_charset()
    if not charset:
        match = re.search(rb"<meta[^>]+charset=[\"']?([A-Za-z0-9_-]+)", raw[:4096], re.IGNORECASE)
        charset = match.group(1).decode("ascii") if match else "utf-8"
    return raw.decode(charset, errors="replace")


def parse_staff(html):
    parser = StaffCardParser()
    parser.feed(html)
    parser.close()

    staff = []
    seen = set()
    for card in parser.cards:
        raw_name = card["name"] or ""
        if not raw_name:
            continue
        name, title = split_name_and_title(raw_name)
        href = card["profile_href"]
        profile_url = urljoin(SOURCE_URL, href) if href else None
        key = profile_url or name.casefold()
        if key in seen:
            print("warning: duplicate entry skipped: %s" % name, file=sys.stderr)
            continue
        seen.add(key)
        staff.append({
            "name": name,
            "title": title,
            "position": card["position"],
            "profile_url": profile_url,
        })
    return staff


def match_key(name):
    normalized = unicodedata.normalize("NFC", name)
    return collapse(normalized).casefold()


def load_existing(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as handle:
        data = json.load(handle)
    return data if isinstance(data, list) else []


# hand-curated fields that the page does not provide
CURATED = ("name_th", "code", "scopus_ids", "scholar_url", "openalex_ids", "match_note", "manual")


def merge_existing(staff, existing):
    by_name = {match_key(entry.get("name", "")): entry for entry in existing if entry.get("name")}
    matched = set()
    for person in staff:
        old = by_name.get(match_key(person["name"])) or {}
        person.setdefault("openalex_ids", [])
        person.setdefault("match_note", "")
        for field in CURATED:
            if field in old:
                person[field] = old[field]
        if old:
            matched.add(match_key(person["name"]))
    for entry in existing:
        if not entry.get("name") or match_key(entry["name"]) in matched:
            continue
        if entry.get("manual"):
            staff.append(entry)  # added by hand, not on the page
        else:
            print("warning: existing entry not on the page, dropped: %s" % entry["name"], file=sys.stderr)
    return staff


def write_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp_path = path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    os.replace(tmp_path, path)


def main():
    html = fetch_html(SOURCE_URL)
    staff = parse_staff(html)
    if not staff:
        print("error: no staff cards found on %s; leaving %s unchanged" % (SOURCE_URL, OUTPUT_PATH),
              file=sys.stderr)
        return 1
    staff = merge_existing(staff, load_existing(OUTPUT_PATH))
    write_json(OUTPUT_PATH, staff)
    print("wrote %d staff to %s" % (len(staff), OUTPUT_PATH))
    return 0


if __name__ == "__main__":
    sys.exit(main())
