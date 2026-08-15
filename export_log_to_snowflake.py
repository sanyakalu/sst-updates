import argparse, io, json, os, requests, zipfile
import snowflake.connector

parser = argparse.ArgumentParser()
parser.add_argument("--run-id", required=True)
parser.add_argument("--workflow", required=True)
parser.add_argument("--branch", required=True)
parser.add_argument("--actor", required=True)
parser.add_argument("--status", required=True)
args = parser.parse_args()

# Fetch run metadata from GitHub API
headers = {"Authorization": f"Bearer {os.environ['GH_TOKEN']}"}
repo = os.environ["GITHUB_REPOSITORY"]
run = requests.get(
    f"https://api.github.com/repos/{repo}/actions/runs/{args.run_id}",
    headers=headers
).json()

# Fetch actual log text (zip of per-job .txt files)
logs_response = requests.get(
    f"https://api.github.com/repos/{repo}/actions/runs/{args.run_id}/logs",
    headers=headers
)
log_text = ""
if logs_response.status_code == 200:
    with zipfile.ZipFile(io.BytesIO(logs_response.content)) as zf:
        parts = []
        for name in sorted(zf.namelist()):
            if name.endswith(".txt"):
                parts.append(f"=== {name} ===\n" + zf.read(name).decode("utf-8", errors="replace"))
        log_text = "\n\n".join(parts)

conn = snowflake.connector.connect(
    account="ywa73928.east-us-2.azure",
    user="sanya.kaluarachchi@philips.com",
    authenticator="programmatic_access_token",
    token=os.environ["SNOWFLAKE_PAT"],
    warehouse="SNOWFLAKE_LEARNING_WH",
    database="PHILIPS_APPS",
    schema="IX_TOOLS_HUB",
)

conn.cursor().execute("""
    INSERT INTO DEPLOYMENT_LOGS
        (RUN_ID, WORKFLOW, BRANCH, TRIGGERED_BY, JOB_STATUS,
        STARTED_AT, COMPLETED_AT, LOG_URL, RAW_LOG, LOG_TEXT)
    SELECT
        %(run_id)s, %(workflow)s, %(branch)s, %(actor)s, %(status)s,
        %(started_at)s::TIMESTAMP_TZ, %(completed_at)s::TIMESTAMP_TZ,
        %(log_url)s, TRY_PARSE_JSON(%(raw)s), %(log_text)s
""", {
    "run_id": int(args.run_id),
    "workflow": args.workflow,
    "branch": args.branch,
    "actor": args.actor,
    "status": args.status,
    "started_at": run.get("run_started_at"),
    "completed_at": run.get("updated_at"),
    "log_url": run.get("html_url"),
    "raw": json.dumps(run),
    "log_text": log_text,
})
conn.close()