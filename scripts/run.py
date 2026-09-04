#!/usr/bin/env python3
"""Unified SST tool runner — MODE env var selects the operation.

Modes:
  ms-updates      -- scrape Microsoft Update Catalog  (needs YEAR, MONTH as integers)
  fetch-dvr-data  -- fetch Azure DevOps test/STICR data (needs MONTH 3-letter, YEAR)
  create-sticrs   -- create STICR work items in ADO (needs PAYLOAD_B64, SST_ADO_PAT)
"""

import base64, json, os, re, sys
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd


# ═══════════════════════════════════════════════════════════════════
#  ms-updates
# ═══════════════════════════════════════════════════════════════════

def run_ms_updates(year: int, month: int):
    print(f"Pulling Microsoft Update Catalog for {year}-{month:02d}")

    search_terms = [
        f'{year}-{month:02d}',
        'Edge Stable',
        'Edge Runtime',
        'Defender Antivirus antimalware platform',
        'KB890830',
    ]

    exclude_from_title = [
        '23H2', 'Server 2008', 'Driver Update', 'Update Preview',
        'Server 2012', 'Security Only Quality Update',
        'Security and Quality Rollup', 'Security Monthly Quality Rollup',
        'Dynamic Update', 'Dynamic Cumulative Update',
        'x86', 'ARM64', 'Windows 11', '.NET 10', '.NET 9', '24H2', '22H2',
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
            if any(ex.lower() in title.lower() for ex in exclude_from_title):
                continue
            if '.net framework' in title.lower() and qualifying_net_framework.lower() not in title.lower():
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

                new_rows = [(t, lu) for t, lu in page_rows if t.lower() not in seen_titles]
                for t, lu in new_rows:
                    seen_titles.add(t.lower())
                all_rows.extend(new_rows)
                if new_rows:
                    print(f"  -> {len(new_rows)} result(s)")

                if past_month_found:
                    print("  Reached past-month entries, stopping term.")
                    break

                page += 1

        seen = set()
        return [r for r in all_rows if r[0].lower() not in seen and not seen.add(r[0].lower())]

    def extract_products(titles_list):
        patterns = [
            r'Windows Server [A-Za-z0-9]{4}',
            r'Windows 10 Version [A-Za-z0-9]{4}',
            r'Microsoft server operating system version [A-Za-z0-9]{4}',
        ]
        found = []
        for title in titles_list:
            for pattern in patterns:
                m = re.search(pattern, title, re.IGNORECASE)
                if m:
                    if m.group(0).lower() == 'microsoft server operating system version 21h2':
                        normalized = 'Windows Server 2022'
                    else:
                        normalized = m.group(0).title()
                    normalized = re.sub(
                        r'Microsoft Server Operating System Version',
                        'Windows Server Version',
                        normalized,
                        flags=re.IGNORECASE,
                    )
                    found.append(normalized)
        seen_p = set()
        return [x for x in found if not (x in seen_p or seen_p.add(x))]

    def extract_build(title):
        m = re.search(r'\(Build ([^)]+)\)', title)
        return m.group(1) if m else None

    def extract_kb(title):
        m = re.search(r'KB\d+', title)
        return m.group(0) if m else 'N/A'

    def build_dataframes(table_results):
        edge_rows  = [(t, lu) for t, lu in table_results if 'microsoft edge' in t.lower()]
        other_rows = [(t, lu) for t, lu in table_results if 'microsoft edge' not in t.lower()]

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

    table_results = pull_catalog_table()
    print(f"\nTotal results: {len(table_results)}")

    edge_df, proposed = build_dataframes(table_results)
    content = "\n\n\n".join([format_updates(proposed), format_edge(edge_df)])

    result_b64 = base64.b64encode(content.encode('utf-8')).decode('ascii')
    print(f"##MS_UPDATES_RESULT##{result_b64}")


# ═══════════════════════════════════════════════════════════════════
#  fetch-dvr-data
# ═══════════════════════════════════════════════════════════════════

def run_fetch_dvr_data(month: str, year: str):
    MONTH_NUM = {
        'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
        'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12,
    }
    MONTH_NAMES_FULL = {
        1:"January",2:"February",3:"March",4:"April",5:"May",6:"June",
        7:"July",8:"August",9:"September",10:"October",11:"November",12:"December",
    }

    pat      = os.environ["SST_ADO_PAT"]
    auth     = base64.b64encode(f":{pat}".encode()).decode()
    hdrs     = {"Authorization": f"Basic {auth}", "Content-Type": "application/json"}
    get_hdrs = {"Authorization": f"Basic {auth}"}

    ORG      = "PhilipsMA"
    PROJECT  = "Philips.PIC"
    AREA     = "Philips.PIC\\SysEng - ESS"
    WIQL_URL = f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/wiql?api-version=7.1"

    month_num   = MONTH_NUM[month]
    prev_month  = month_num - 1 if month_num > 1 else 12
    prev_year   = year if month_num > 1 else str(int(year) - 1)
    month_start = f"{prev_year}-{prev_month:02d}-16"
    month_end   = f"{year}-{month_num:02d}-15"

    KEYWORDS  = [
        "Crowdstrike","VMWare", "Security","SQL Server","Symantec",
        "TrendMicro","Trellix","VMware","Nutanix","Hyper-V","Microsoft Security Updates",
    ]
    kw_clauses = " OR ".join(f"[System.Title] CONTAINS '{k}'" for k in KEYWORDS)

    plan_resp = requests.post(WIQL_URL, headers=hdrs, json={"query":
        f"SELECT [System.Id],[System.Title],[System.CreatedDate] FROM WorkItems "
        f"WHERE [System.WorkItemType]='Test Plan' AND [System.AreaPath] UNDER '{AREA}' "
        f"AND [System.CreatedDate]>='{month_start}' AND [System.CreatedDate]<='{month_end}' "
        f"AND ({kw_clauses}) ORDER BY [System.CreatedDate] DESC"
    })
    plan_ids = [w["id"] for w in plan_resp.json().get("workItems", [])]
    print(f"DVR: found {len(plan_ids)} test plan(s)")

    test_rows  = []
    run_cache  = {}

    for plan_id in plan_ids:
        plan = requests.get(
            f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/testplan/plans/{plan_id}?api-version=7.1",
            headers=get_hdrs,
        ).json()
        plan_name = plan.get("name", "")

        suites = requests.get(
            f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/testplan/plans/{plan_id}/suites?api-version=7.1",
            headers=get_hdrs,
        ).json().get("value", [])

        for suite in suites:
            suite_id   = suite["id"]
            suite_name = suite["name"]

            points = requests.get(
                f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/testplan/plans/{plan_id}"
                f"/suites/{suite_id}/TestPoint?includePointDetails=true&api-version=7.1",
                headers=get_hdrs,
            ).json().get("value", [])

            for point in points:
                last_result    = point.get("results") or {}
                outcome        = last_result.get("outcome", "N/A")
                run_id         = last_result.get("lastTestRunId") or "N/A"
                pipeline_run   = last_result.get("lastRunBuildNumber", "N/A")
                test_case_id   = (point.get("testCaseReference") or {}).get("id", "N/A")
                test_case_name = (point.get("testCaseReference") or {}).get("name", "N/A")

                run_name = "N/A"
                if run_id != "N/A":
                    if run_id not in run_cache:
                        run_data = requests.get(
                            f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/test/runs/{run_id}?api-version=7.1",
                            headers=get_hdrs,
                        ).json()
                        raw_name = run_data.get("name", "N/A")
                        run_cache[run_id] = re.sub(r'\s*\(Manual\)', '', raw_name).strip()
                    run_name = run_cache[run_id]

                # strip PIIC_iX_ prefix
                if isinstance(pipeline_run, str) and pipeline_run.startswith("PIIC_iX_"):
                    pipeline_run = pipeline_run[len("PIIC_iX_"):]
                # numeric fallback: bare build ID → extract version from run name
                if isinstance(pipeline_run, str) and re.match(r'^\d+$', pipeline_run):
                    m = re.search(r'([A-Za-z0-9]+\.[A-Za-z0-9.]+)', run_name)
                    if m:
                        pipeline_run = m.group(1)
                # filter to supported builds
                if not isinstance(pipeline_run, str) or pipeline_run[:1] not in ("4", "C", "B"):
                    continue

                test_rows.append([
                    plan_id, suite_id, plan_name, suite_name,
                    run_id, run_name, outcome,
                    test_case_name, test_case_id, pipeline_run,
                ])

    print(f"DVR: fetched {len(test_rows)} test result(s)")

    sticr_resp = requests.post(WIQL_URL, headers=hdrs, json={"query":
        f"SELECT [System.Id],[System.Title] FROM WorkItems "
        f"WHERE [System.WorkItemType]='STICR' AND [System.AreaPath] UNDER '{AREA}' "
        f"AND [System.TeamProject]='Philips.PIC' "
        f"AND [System.Title] CONTAINS 'Microsoft Security Update' "
        f"AND [System.Title] CONTAINS '{MONTH_NAMES_FULL[month_num]}' "
        f"AND [System.Title] CONTAINS '{year}' "
        f"ORDER BY [System.CreatedDate] DESC"
    })
    sticr_ids = [w["id"] for w in sticr_resp.json().get("workItems", [])]
    print(f"DVR: found {len(sticr_ids)} STICR(s)")

    sticr_items = []
    for sid in sticr_ids:
        item = requests.get(
            f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/workitems/{sid}?$expand=fields&api-version=7.1",
            headers=get_hdrs,
        ).json()
        sticr_items.append({
            "id":          sid,
            "description": item.get("fields", {}).get("System.Description", ""),
        })

    payload    = {"testRows": test_rows, "sticrItems": sticr_items}
    result_b64 = base64.b64encode(json.dumps(payload).encode('utf-8')).decode('ascii')
    print(f"##DVR_DATA_RESULT##{result_b64}")


# ═══════════════════════════════════════════════════════════════════
#  create-sticrs
# ═══════════════════════════════════════════════════════════════════

def run_create_sticrs():
    payload  = json.loads(base64.b64decode(os.environ["PAYLOAD_B64"]).decode("utf-8"))
    pat      = os.environ["SST_ADO_PAT"]
    auth     = base64.b64encode(f":{pat}".encode()).decode()
    hdrs     = {
        "Content-Type":  "application/json-patch+json",
        "Authorization": f"Basic {auth}",
    }

    sticr_data      = payload["sticrs"]
    template_fields = payload.get("templateFields", [])
    project         = payload.get("project", "Sandbox")

    ORG        = "PhilipsMA"
    create_url = f"https://dev.azure.com/{ORG}/{project}/_apis/wit/workitems/$STICR?api-version=7.1"

    created, failed = 0, 0
    created_items   = []

    for item in sticr_data:
        body = [
            {"op": "add", "path": "/fields/System.Title",        "value": item["title"]},
            {"op": "add", "path": "/fields/System.Description",  "value": item["html"]},
            {"op": "add", "path": "/fields/System.AreaPath",      "value": project},
            {"op": "add", "path": "/fields/System.TeamProject",   "value": project},
            {"op": "add", "path": "/fields/System.IterationPath", "value": project + "\\Common"},
            {"op": "add", "path": "/fields/System.AssignedTo",    "value": item.get("userEmail", "")},
            *template_fields,
        ]
        resp = requests.post(create_url, headers=hdrs, json=body)
        if resp.ok:
            wi  = resp.json()
            url = f"https://dev.azure.com/{ORG}/{project}/_workitems/edit/{wi['id']}"
            print(f"Created STICR {wi['id']}: {item['title']}")
            created_items.append({"id": wi["id"], "title": item["title"], "url": url})
            created += 1
        else:
            print(f"FAILED ({resp.status_code}): {item['title']}\n{resp.text}")
            failed += 1

    print(f"\nSummary: {created} created, {failed} failed")
    print(f"##STICR_RESULTS##{json.dumps(created_items)}")

    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary_path and created_items:
        with open(summary_path, "a") as f:
            f.write("## STICRs Created\n\n")
            for item in created_items:
                f.write(f"- [#{item['id']} — {item['title']}]({item['url']})\n")


# ═══════════════════════════════════════════════════════════════════
#  Entry point
# ═══════════════════════════════════════════════════════════════════

mode = os.environ.get("MODE", "").strip()
if mode == "ms-updates":
    run_ms_updates(int(os.environ["YEAR"]), int(os.environ["MONTH"]))
elif mode == "fetch-dvr-data":
    run_fetch_dvr_data(os.environ["MONTH"], os.environ["YEAR"])
elif mode == "create-sticrs":
    run_create_sticrs()
else:
    print(f"Unknown MODE: {mode!r}", file=sys.stderr)
    sys.exit(1)
