import base64, json, os, requests

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

import pathlib
pathlib.Path("sticr_results.json").write_text(
    json.dumps({"project": project, "items": created_items})
)

summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
if summary_path and created_items:
    with open(summary_path, "a") as f:
        f.write("## STICRs Created\n\n")
        for item in created_items:
            f.write(f"- [#{item['id']} — {item['title']}]({item['url']})\n")
