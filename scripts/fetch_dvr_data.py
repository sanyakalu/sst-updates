import argparse, base64, json, os, pathlib, requests

MONTH_NUM = {
    'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
    'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12,
}
MONTH_NAMES_FULL = {
    1:"January",2:"February",3:"March",4:"April",5:"May",6:"June",
    7:"July",8:"August",9:"September",10:"October",11:"November",12:"December",
}

parser = argparse.ArgumentParser()
parser.add_argument("--month", required=True, help="3-letter month, e.g. Jan")
parser.add_argument("--year",  required=True, help="4-digit year, e.g. 2026")
args = parser.parse_args()

pat  = os.environ["SST_ADO_PAT"]
auth = base64.b64encode(f":{pat}".encode()).decode()
hdrs = {"Authorization": f"Basic {auth}", "Content-Type": "application/json"}
get_hdrs = {"Authorization": f"Basic {auth}"}

ORG      = "PhilipsMA"
PROJECT  = "Philips.PIC"
AREA     = "Philips.PIC\\SysEng - ESS"
WIQL_URL = f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/wit/wiql?api-version=7.1"

month_num  = MONTH_NUM[args.month]
year       = args.year
prev_month = month_num - 1 if month_num > 1 else 12
prev_year  = year if month_num > 1 else str(int(year) - 1)
month_start = f"{prev_year}-{prev_month:02d}-16"
month_end   = f"{year}-{month_num:02d}-15"

# ── Test plans ────────────────────────────────────────────────────────────────

KEYWORDS  = [
    "Crowdstrike","VMWare","OS Security","SQL Server","Symantec",
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

test_rows = []
for plan_id in plan_ids:
    plan = requests.get(
        f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/testplan/plans/{plan_id}?api-version=7.1",
        headers=get_hdrs,
    ).json()
    runs = requests.get(
        f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/test/runs?planId={plan_id}&api-version=7.1",
        headers=get_hdrs,
    ).json().get("value", [])
    for run in runs:
        results = requests.get(
            f"https://dev.azure.com/{ORG}/{PROJECT}/_apis/test/runs/{run['id']}/results?api-version=7.1",
            headers=get_hdrs,
        ).json().get("value", [])
        for r in results:
            test_rows.append([
                plan_id, plan.get("name", ""), run.get("name", ""), run["id"],
                run.get("state", "N/A"), r.get("outcome", "N/A"),
                r.get("testCaseTitle", "N/A"), (r.get("testCase") or {}).get("id", "N/A"),
            ])
print(f"DVR: fetched {len(test_rows)} test result(s)")

# ── STICRs ────────────────────────────────────────────────────────────────────

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

# ── Write output ──────────────────────────────────────────────────────────────

out_dir  = pathlib.Path("dvr_data")
out_dir.mkdir(exist_ok=True)
out_file = out_dir / f"{args.month}_{year}.json"
out_file.write_text(json.dumps({"testRows": test_rows, "sticrItems": sticr_items}, indent=2))
print(f"Written {out_file}")
