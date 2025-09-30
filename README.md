# Microsoft Graph Excel Helper

Small utilities to interact with Excel files on OneDrive using the Microsoft Graph API.

## Overview

This repository contains two main modules used to authenticate against Microsoft Graph and to add a row of data into a dated worksheet inside an Excel workbook stored in OneDrive.

- `ms_graph.py` — authentication helpers using MSAL and a browser-assisted auth code flow (Selenium). It exposes `get_access_token()` and the `MS_GRAPH_URL` constant.
- `excel.py` — a helper `add_data_to_dated_sheet(file_id, data, access_token=None)` that:
  - optionally creates a dated worksheet (YYYY-MM-DD) if it doesn't exist,
  - finds the first empty row using `usedRange()` and `rowCount`,
  - writes a single row of data (a 2D `values` array) into the correct A..Z range.

## Prerequisites

- Python 3.10+
- An Azure AD app registration with a client ID, client secret and tenant ID configured in environment variables: `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`.
- The app must have permissions for appropriate Graph scopes (delegated or application depending on your flow):
  - `Files.ReadWrite.All`
  - `Sites.ReadWrite.All`

Optional (for automated browser auth):
- Microsoft Edge or Firefox and the corresponding webdriver (the project can download the Edge driver automatically using `webdriver_manager`).

## Installation

1. Create and activate a virtual environment (recommended):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Install dependencies:

```powershell
pip install -r requirements.txt
```

## Usage

1. Ensure environment variables are set (or populated in a `.env` file):

- `CLIENT_ID`
- `CLIENT_SECRET`
- `TENANT_ID`

2. Obtain an access token (the library will prompt/browser-sign you in if needed):

```python
from ms_graph import get_access_token
access_token = get_access_token(scopes=["Files.ReadWrite.All", "Sites.ReadWrite.All"]) 
```

3. Use `OneDrive` helper to get the file id and call `add_data_to_dated_sheet`:

```python
from excel import add_data_to_dated_sheet
from onedrive import OneDrive

onedrive = OneDrive(access_token)
file_id = onedrive.get_item_id("C:/path/to/your.xlsx")
add_data_to_dated_sheet(file_id, ["John Doe", 30, "New York"], access_token)
```

Notes:
- `add_data_to_dated_sheet` accepts a `data` parameter either as a `dict` or a `list`. Internally it converts `dict` to a single-row 2D list of stringified values.
- The module supports writing up to 26 columns (A..Z). If you need more columns, update the column address calculation logic.

## Troubleshooting

- If you see syntax errors when calling `range(address=...)`, ensure the address string is wrapped in single quotes and URL-encoded where used in the path.
- If the sheet is created but no data appears, check the PATCH response body and status code — the helper considers 200 and 201 as success codes.
- If authentication fails, run `ms_graph.get_access_token()` directly and follow the interactive prompt. Check the `.env` file and environment variables.

## Future improvements

- Support column addresses beyond `Z` (AA, AB, ...).
- Add unit tests and CI workflow.
- Add retry/backoff for transient Graph API failures.

## License

MIT
