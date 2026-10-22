"""
Link BOV Test Cases to User Stories in Azure DevOps using the "Tested By" relation.
 
Source: "Rate Config User Stories Tracker.xlsx", sheet "User Stories"
  Column A = USER STORY ID, Column N = Link Status, Column T = BOV TC ID(s)
Only rows with Link Status = "Not Started" and purely numeric TC IDs are processed.
 
Modes:
  --plan-only : Reads Excel only, builds the plan. No ADO connection.
  (default)   : Dry run. Reads ADO to validate IDs and existing links. NO writes.
  --execute   : Same as dry run, then creates the missing "Tested By" links.
 
Examples:
  python link_testcases.py --plan-only
  python link_testcases.py                          # dry run against ADO
  python link_testcases.py --execute --max-stories 5   # pilot
  python link_testcases.py --execute                # full run
 
PAT: taken from env var ADO_PAT, otherwise prompted (hidden). Needs Work Items: Read & Write.
"""
import argparse
import base64
import csv
import getpass
import json
import os
import re
import ssl
import sys
import time
import urllib.error
import urllib.request
from collections import OrderedDict
from datetime import datetime
from pathlib import Path
 
from openpyxl import load_workbook
 
HERE = Path(__file__).resolve().parent
REL_TESTED_BY = "Microsoft.VSTS.Common.TestedBy-Forward"
API_VERSION = "7.1"
TC_TOKEN = re.compile(r"^\d{5,8}$")
# Column T values to disregard (case-insensitive, matched anywhere in the cell)
TC_EXCLUDE = re.compile(r"descope|\bn/?a\b|shared\s*config", re.IGNORECASE)
 
parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
parser.add_argument("--excel", default=str(HERE / "Rate Config User Stories Tracker.xlsx"))
parser.add_argument("--sheet", default="User Stories")
parser.add_argument("--org", default="BHE-IT-01")
parser.add_argument("--project", default="BHE-Technology")
parser.add_argument("--status", default="Not Started", help="Link Status value to process")
parser.add_argument("--remarks", nargs="+", default=["Rate Config", "Rates Matrix"],
                    help="Column P (Remarks) values to include, exact match ignoring case")
parser.add_argument("--max-stories", type=int, default=0, help="Limit number of user stories (pilot)")
parser.add_argument("--story-id", type=int, nargs="+", help="Only process these User Story IDs")
parser.add_argument("--plan-only", action="store_true")
parser.add_argument("--execute", action="store_true")
parser.add_argument("--output", default=str(HERE / "Output"))
args = parser.parse_args()
 
BASE_URL = f"https://dev.azure.com/{args.org}"
STAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
OUT = Path(args.output)
OUT.mkdir(parents=True, exist_ok=True)
LOG_FILE = OUT / f"run_{STAMP}.log"
 
 
def log(msg, level="INFO"):
    line = f"{datetime.now():%H:%M:%S} [{level}] {msg}"
    print(line, flush=True)
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(line + "\n")
 
 
def write_csv(path, rows, fields):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)
 
 
def cell_text(v):
    if v is None:
        return ""
    # Excel stores single IDs as floats; keep huge/corrupted numbers (e.g. 1.08E+20) as-is so they fail validation
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e9:
        return str(int(v))
    return str(v).strip()
 
 
# ---------------------------------------------------------------- 1. Read Excel
log(f"Reading '{args.excel}' sheet '{args.sheet}' (read-only)")
wb = load_workbook(args.excel, read_only=True, data_only=True)
ws = wb[args.sheet]
rows = []
for idx, r in enumerate(ws.iter_rows(min_row=2, max_col=20, values_only=True), start=2):
    us = cell_text(r[0])
    if not us:
        continue
    rows.append({"Row": idx, "UserStory": us, "LinkStatus": cell_text(r[13]), "Remarks": cell_text(r[15]),
                 "TcRaw": cell_text(r[19])})
wb.close()
 
remarks_ok = {s.strip().lower() for s in args.remarks}
candidates = [r for r in rows if r["LinkStatus"] == args.status]
log(f"Rows read: {len(rows)}; '{args.status}': {len(candidates)}; Remarks filter: {args.remarks}")
 
plan, skipped = [], []
for c in candidates:
    def skip(reason):
        skipped.append({"Row": c["Row"], "UserStory": c["UserStory"], "Remarks": c["Remarks"],
                        "TcRaw": c["TcRaw"], "Reason": reason})
 
    if c["Remarks"].lower() not in remarks_ok:
        skip("Remarks not in filter"); continue
    if not c["UserStory"].isdigit():
        skip("Invalid User Story ID"); continue
    if not c["TcRaw"]:
        skip("No TC ID"); continue
    if TC_EXCLUDE.search(c["TcRaw"]):
        skip("Excluded TC value (Descoped / N/A / Shared Config)"); continue
    tokens = [t for t in re.split(r"[,;\s]+", c["TcRaw"]) if t]
    # Any non-numeric token (N/A, Descoped, 1.08E+20, notes) => skip whole row for manual review
    if any(not TC_TOKEN.match(t) for t in tokens):
        skip("Non-numeric / malformed TC value"); continue
    for tc in OrderedDict.fromkeys(tokens):
        plan.append({"Row": c["Row"], "UserStory": int(c["UserStory"]), "TestCase": int(tc), "Action": "ToLink", "Detail": ""})
 
if args.story_id:
    wanted = set(args.story_id)
    plan = [p for p in plan if p["UserStory"] in wanted]
    missing = wanted - {p["UserStory"] for p in plan}
    if missing:
        log(f"Not in plan (wrong status / no valid TC / not in sheet): {sorted(missing)}", "WARN")
story_ids = list(OrderedDict.fromkeys(p["UserStory"] for p in plan))
if args.max_stories > 0:
    story_ids = story_ids[: args.max_stories]
    keep = set(story_ids)
    plan = [p for p in plan if p["UserStory"] in keep]
tc_ids = list(OrderedDict.fromkeys(p["TestCase"] for p in plan))
 
log(f"Plan: {len(story_ids)} user stories, {len(tc_ids)} distinct test cases, {len(plan)} links. Skipped rows: {len(skipped)}")
write_csv(OUT / f"skipped_rows_{STAMP}.csv", skipped, ["Row", "UserStory", "Remarks", "TcRaw", "Reason"])
PLAN_FIELDS = ["Row", "UserStory", "TestCase", "Action", "Detail"]
 
if args.plan_only:
    write_csv(OUT / f"plan_{STAMP}.csv", plan, PLAN_FIELDS)
    log(f"PlanOnly: no ADO calls made. Files in {OUT}", "OK")
    sys.exit(0)
 
# ---------------------------------------------------------------- 2. ADO helpers
pat = (os.environ.get("ADO_PAT") or getpass.getpass("Enter Azure DevOps PAT (Work Items: Read & Write): ")).strip()
log(f"PAT received: {len(pat)} characters")
if not pat:
    log("Empty PAT - paste with right-click (Ctrl+V may not work in hidden prompts)", "ERROR")
    sys.exit(1)
AUTH = "Basic " + base64.b64encode(f":{pat}".encode()).decode()
SSL_CTX = ssl.create_default_context()  # uses Windows cert store (incl. corporate root CAs)
 
 
def ado(method, url, body=None, content_type="application/json"):
    data = json.dumps(body).encode("utf-8") if body is not None else None
    for attempt in range(1, 6):
        req = urllib.request.Request(url, data=data, method=method,
                                     headers={"Authorization": AUTH, "Content-Type": content_type, "Accept": "application/json"})
        try:
            with urllib.request.urlopen(req, context=SSL_CTX, timeout=120) as resp:
                if resp.status == 203:
                    raise RuntimeError("Unauthorized (HTTP 203) - check PAT")
                raw = resp.read()
                return json.loads(raw) if raw else None
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503, 504) and attempt < 5:
                wait = int(e.headers.get("Retry-After") or 2 ** attempt)
                log(f"HTTP {e.code} on {url} - retry in {wait}s", "WARN")
                time.sleep(wait)
                continue
            detail = e.read().decode("utf-8", "replace")[:500]
            raise RuntimeError(f"HTTP {e.code}: {detail}") from None
 
 
def get_work_items(ids, with_relations=False):
    result = {}
    for i in range(0, len(ids), 200):
        body = {"ids": ids[i:i + 200], "errorPolicy": "Omit"}
        if with_relations:
            body["$expand"] = "Relations"
        else:
            body["fields"] = ["System.Id", "System.WorkItemType", "System.Title", "System.State"]
        resp = ado("POST", f"{BASE_URL}/{args.project}/_apis/wit/workitemsbatch?api-version={API_VERSION}", body)
        for wi in resp.get("value", []):
            if wi:
                result[int(wi["id"])] = wi
    return result
 
 
# ---------------------------------------------------------------- 3. Validate against ADO
try:
    me = ado("GET", f"{BASE_URL}/_apis/connectionData")
    user = (me or {}).get("authenticatedUser", {})
    if user.get("descriptor", "").startswith("System:PublicAccess") or not user.get("providerDisplayName"):
        raise RuntimeError("token not accepted (anonymous)")
    ado("GET", f"{BASE_URL}/_apis/projects/{args.project}?api-version={API_VERSION}")
    log(f"Authenticated to {BASE_URL}/{args.project} as {user.get('providerDisplayName')}", "OK")
except Exception as e:
    log(f"Cannot reach ADO / auth failed: {e}", "ERROR")
    log(f"Check: PAT created in org '{args.org}' (or 'All accessible organizations'), not expired, "
        "scope Work Items Read & Write, copied fully.", "ERROR")
    sys.exit(1)
 
log(f"Fetching {len(tc_ids)} test cases...")
tc_map = get_work_items(tc_ids)
log(f"Fetching {len(story_ids)} user stories with relations...")
us_map = get_work_items(story_ids, with_relations=True)
 
for p in plan:
    us, tc = us_map.get(p["UserStory"]), tc_map.get(p["TestCase"])
    if not us:
        p["Action"], p["Detail"] = "Skip", "User Story not found / no access"; continue
    us_type = us["fields"].get("System.WorkItemType")
    if us_type != "User Story":
        p["Action"], p["Detail"] = "Skip", f"US id is a '{us_type}'"; continue
    if not tc:
        p["Action"], p["Detail"] = "Skip", "Test Case not found / no access"; continue
    tc_type = tc["fields"].get("System.WorkItemType")
    if tc_type != "Test Case":
        p["Action"], p["Detail"] = "Skip", f"TC id is a '{tc_type}'"; continue
    existing = {int(r["url"].rsplit("/", 1)[-1]) for r in us.get("relations") or [] if r.get("rel") == REL_TESTED_BY}
    if p["TestCase"] in existing:
        p["Action"] = "AlreadyLinked"
 
plan_file = OUT / f"plan_{STAMP}.csv"
write_csv(plan_file, plan, PLAN_FIELDS)
 
 
def summary():
    counts = {}
    for p in plan:
        key = f"{p['Action']} {p['Detail']}".strip()
        counts[key] = counts.get(key, 0) + 1
    for k, v in sorted(counts.items(), key=lambda kv: -kv[1]):
        log(f"  {v:>6}  {k}")
 
 
summary()
 
# ---------------------------------------------------------------- 4. Execute
to_link = OrderedDict()
for p in plan:
    if p["Action"] == "ToLink":
        to_link.setdefault(p["UserStory"], []).append(p)
 
if not args.execute:
    log(f"DRY RUN complete. {len(to_link)} user stories would be updated. Review {plan_file}, then re-run with --execute.", "OK")
    sys.exit(0)
 
log(f"EXECUTE: updating {len(to_link)} user stories", "WARN")
for n, (us_id, items) in enumerate(to_link.items(), start=1):
    # One PATCH per story with all missing links (atomic per story)
    ops = [{"op": "add", "path": "/relations/-", "value": {
        "rel": REL_TESTED_BY,
        "url": f"{BASE_URL}/_apis/wit/workItems/{p['TestCase']}",
        "attributes": {"comment": "Auto-linked from Rate Config User Stories Tracker"}}} for p in items]
    try:
        ado("PATCH", f"{BASE_URL}/{args.project}/_apis/wit/workitems/{us_id}?api-version={API_VERSION}",
            ops, "application/json-patch+json")
        for p in items:
            p["Action"] = "Linked"
        log(f"[{n}/{len(to_link)}] US {us_id} linked {len(ops)} TC(s)", "OK")
    except Exception as e:
        for p in items:
            p["Action"], p["Detail"] = "Failed", str(e)
        log(f"[{n}/{len(to_link)}] US {us_id} FAILED: {e}", "ERROR")
 
result_file = OUT / f"results_{STAMP}.csv"
write_csv(result_file, plan, PLAN_FIELDS)
summary()
log(f"Done. Results: {result_file}", "OK")