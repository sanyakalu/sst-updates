#!/usr/bin/env python3
"""Pull this month's updates from the Microsoft Update Catalog.

Usage: python pull_ms_updates.py --year YYYY --month M
Writes output to ms_updates/{year}_{month:02d}.txt
"""

import re, sys, os, argparse
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd

parser = argparse.ArgumentParser()
parser.add_argument('--year',  type=int, required=True)
parser.add_argument('--month', type=int, required=True)
args = parser.parse_args()

year  = args.year
month = args.month

print(f"Pulling Microsoft Update Catalog for {year}-{month:02d}")

search_terms = [
    f'{year}-{month:02d}',
    'Edge Stable',
    'Edge Runtime',
    'Microsoft Defender Antivirus antimalware platform',
    'KB890830',
]

exclude_from_title = [
    'Security and Quality Rollup',
    'Security Monthly Quality Rollup',
    'Dynamic Update',
    'Dynamic Cumulative Update',
    'x86',
    'ARM64',
    'Windows 11',
    '.NET 10',
    '.NET 9',
    '24H2',
    '22H2',
]

qualifying_net_framework = '4.8 '

session = requests.Session()
session.headers.update({
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.5',
})


def fetch_page(url):
    try:
        resp = session.get(url, timeout=30)
        resp.raise_for_status()
        return BeautifulSoup(resp.text, 'lxml')
    except Exception as e:
        print(f"  Error fetching {url}: {e}", file=sys.stderr)
        return None


def parse_table_rows(soup):
    """Returns (page_rows, past_month_found, has_any_rows)."""
    table = soup.find('table', id='ctl00_catalogBody_updateMatches')
    if not table:
        return [], False, False
    trs = table.find_all('tr')[1:]
    if not trs:
        return [], False, False

    page_rows = []
    past_month_found = False

    for tr in trs:
        cols = tr.find_all('td')
        if len(cols) < 5:
            continue
        title        = cols[1].get_text(strip=True)
        last_updated = cols[4].get_text(strip=True)
        try:
            lu = datetime.strptime(last_updated, '%m/%d/%Y')
        except ValueError:
            continue
        if (lu.year, lu.month) < (year, month):
            past_month_found = True
            continue
        if (lu.year, lu.month) > (year, month):
            continue
        if any(ex in title for ex in exclude_from_title):
            continue
        if '.NET Framework' in title and qualifying_net_framework not in title:
            continue
        page_rows.append((title, last_updated))

    return page_rows, past_month_found, True


def pull_catalog_table():
    all_rows    = []
    seen_titles = set()

    for term in search_terms:
        encoded  = term.replace(' ', '%20')
        base_url = f"https://www.catalog.update.microsoft.com/Search.aspx?q={encoded}&scol=DateComputed&sdir=desc"
        page = 0

        while True:
            url = base_url if page == 0 else f"{base_url}&p={page}"
            print(f"[{term}] page {page}: {url}")

            soup = fetch_page(url)
            if not soup:
                print("  Failed to fetch, stopping term.")
                break

            page_rows, past_month_found, has_rows = parse_table_rows(soup)

            if not has_rows:
                print(f"  No rows on page {page}, stopping term.")
                break

            new_rows = [(t, lu) for t, lu in page_rows if t not in seen_titles]
            for t, lu in new_rows:
                seen_titles.add(t)
            all_rows.extend(new_rows)
            if new_rows:
                print(f"  -> {len(new_rows)} result(s)")

            if past_month_found:
                print("  Reached past-month entries, stopping term.")
                break

            page += 1

    seen = set()
    return [r for r in all_rows if r[0] not in seen and not seen.add(r[0])]


def extract_products(titles_list):
    found = []
    for title in titles_list:
        for pattern in [r'Windows Server [A-Za-z0-9]{4}', r'Windows 10 Version [A-Za-z0-9]{4}', r'version [A-Za-z0-9]{4}']:
            m = re.search(pattern, title)
            if m:
                found.append(m.group(0))
    seen_p = set()
    return [x for x in found if not (x in seen_p or seen_p.add(x))]


def extract_build(title):
    m = re.search(r'\(Build ([^)]+)\)', title)
    return m.group(1) if m else None


def extract_kb(title):
    m = re.search(r'KB\d+', title)
    return m.group(0) if m else 'N/A'


def build_dataframes(table_results):
    edge_rows  = [(t, lu) for t, lu in table_results if 'Microsoft Edge' in t]
    other_rows = [(t, lu) for t, lu in table_results if 'Microsoft Edge' not in t]

    if edge_rows:
        edge_df = pd.DataFrame(
            [{'date_updated': lu, 'title': t, 'build_number': extract_build(t)} for t, lu in edge_rows]
        )
        edge_df = (
            edge_df.groupby('build_number', sort=False)
            .agg(date_updated=('date_updated', 'first'), titles=('title', list))
            .reset_index()[['date_updated', 'build_number', 'titles']]
        )
        edge_df['products'] = edge_df['titles'].apply(extract_products)
    else:
        edge_df = pd.DataFrame(columns=['date_updated', 'build_number', 'titles', 'products'])

    if other_rows:
        other_updates = pd.DataFrame([
            {'kb_number': extract_kb(t), 'title': t, 'products': extract_products([t])}
            for t, lu in other_rows
        ])
        proposed = (
            other_updates.assign(
                products=other_updates['products'].apply(lambda p: p if p else ['For all clients/servers'])
            )
            .explode('products')
            .groupby('products', sort=False)
            .agg(
                kb_numbers=('kb_number', lambda x: list(dict.fromkeys(x))),
                titles=('title', list)
            )
            .reset_index()
            .rename(columns={'products': 'product'})
        )
    else:
        proposed = pd.DataFrame(columns=['product', 'kb_numbers', 'titles'])

    return edge_df, proposed


def format_edge(df):
    lines = [f"EDGE UPDATES  ({year}-{month:02d})", "=" * 60]
    if df.empty:
        lines.append("\n  No Edge updates found.")
        return "\n".join(lines)
    for _, row in df.iterrows():
        lines.append(f"\nBuild: {row['build_number']}   |   Last Updated: {row['date_updated']}")
        for t in row['titles']:
            lines.append(f"  - {t}")
    return "\n".join(lines)


def format_updates(df):
    lines = [f"UPDATES BY PRODUCT  ({year}-{month:02d})", "=" * 60]
    if df.empty:
        lines.append("\n  No updates found.")
        return "\n".join(lines)
    for _, row in df.iterrows():
        lines.append(f"\n{row['product']}")
        lines.append(f"  KBs   : {', '.join(row['kb_numbers'])}")
        lines.append("  Titles:")
        for t in row['titles']:
            lines.append(f"    - {t}")
    return "\n".join(lines)


# ── Main ──────────────────────────────────────────────────────────────────────

table_results = pull_catalog_table()
print(f"\nTotal results: {len(table_results)}")

edge_df, proposed = build_dataframes(table_results)

content = "\n\n\n".join([format_updates(proposed), format_edge(edge_df)])

os.makedirs("ms_updates", exist_ok=True)
out_path = f"ms_updates/{year}_{month:02d}.txt"
with open(out_path, "w", encoding="utf-8") as f:
    f.write(content)

print(f"Saved to {out_path}")
