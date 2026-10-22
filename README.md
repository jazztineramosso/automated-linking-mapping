# Link Test Cases to User Stories

This project contains a Python script, `link_testcase.py`, that reads an Excel tracker and links Azure DevOps test cases to matching user stories using the `Tested By` relation.

## What the script does

The script:

- Reads the workbook `Rate Config User Stories Tracker.xlsx`
- Uses the `User Stories` sheet
- Filters rows whose `Link Status` matches a configured value (default: `Not Started`)
- Keeps only rows whose `Remarks` column matches the configured remarks list (default: `Rate Config`, `Rates Matrix`)
- Parses the `TC ID(s)` column and keeps only valid 5- to 8-digit numeric values
- Validates the user stories and test cases in Azure DevOps
- Optionally creates the `Microsoft.VSTS.Common.TestedBy-Forward` relation between the user story and each test case

This is intended to automate the manual linking of test cases to user stories in Azure DevOps.

## Requirements

- Python 3.9+
- `openpyxl`
- Access to an Azure DevOps organization and project
- A PAT (Personal Access Token) with `Work Items: Read & Write` permissions

Install the Python dependency:

```bash
pip install openpyxl
```

## Environment and authentication

The script reads the PAT from the environment variable `ADO_PAT` if it is set. Otherwise, it prompts for it interactively.

Example:

```bash
set ADO_PAT=your_pat_here
python link_testcase.py --plan-only
```

On Linux/macOS:

```bash
export ADO_PAT="your_pat_here"
python link_testcase.py --plan-only
```

## Script usage

### 1. Plan only (read Excel only, no Azure DevOps calls)

```bash
python link_testcase.py --plan-only
```

This creates a plan file under the `Output` folder and does not modify Azure DevOps.

### 2. Dry run (validate IDs and relationships without writing)

```bash
python link_testcase.py
```

This performs Azure DevOps read checks and produces a plan, but does not create links.

### 3. Execute (create missing links)

```bash
python link_testcase.py --execute
```

This validates the user stories and test cases, then creates the missing `Tested By` links.

### Optional filters

```bash
python link_testcase.py --execute --max-stories 5
python link_testcase.py --story-id 12345 67890
```

You can also change the workbook, sheet, project, or status filter with arguments such as:

```bash
python link_testcase.py --excel "C:/path/to/file.xlsx" --sheet "User Stories" --project "MyProject" --status "Not Started"
```

## Important behavior

- Only rows with `Link Status` equal to `Not Started` are processed by default.
- Only numeric test case IDs matching the pattern `^\d{5,8}$` are accepted.
- Values such as `N/A`, `Descoped`, or shared-config entries are skipped for manual review.
- Rows with invalid or malformed TC values are recorded in the skipped rows output rather than being linked automatically.

## Output files

The script writes results into an `Output` directory created beside the script. Typical outputs include:

- `plan_YYYYMMDD_HHMMSS.csv`
- `skipped_rows_YYYYMMDD_HHMMSS.csv`
- `results_YYYYMMDD_HHMMSS.csv`

These files summarize the planned links, skipped rows, and final results.

## Notes

- The default mode is a dry run and is safe to use before executing changes.
- The script logs progress to both the console and a timestamped log file in the `Output` directory.
- The `--execute` option is the only mode that writes data to Azure DevOps.

## Example workflow

```bash
python link_testcase.py --plan-only
python link_testcase.py
python link_testcase.py --execute --max-stories 10
```

This lets you review the generated plan before creating links in Azure DevOps.
