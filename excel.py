import requests
import urllib.parse
import string
from datetime import datetime
from ms_graph import get_access_token, MS_GRAPH_URL
from onedrive import OneDrive


def add_data_to_dated_sheet(file_id, data, access_token: str | None = None):
    """
    Adds data to an Excel file using Microsoft Graph API, creating a dated sheet if needed
    and appending to the first empty row.

    Args:
        access_token (str): Valid Microsoft Graph API access token
        file_id (str): ID of the Excel file in OneDrive
        data (dict or list): Data to add to the sheet

    Returns:
        bool: True if successful, False otherwise
    """

    if access_token is None:
        access_token = get_access_token(
            scopes=["Files.ReadWrite.All", "Sites.ReadWrite.All"])

    def get_worksheets():
        """Helper function to get list of worksheets"""
        url = f'{MS_GRAPH_URL}/me/drive/items/{file_id}/workbook/worksheets'
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        response = requests.get(url, headers=headers)
        return response.json().get('value', [])

    def create_new_sheet(sheet_name):
        """Helper function to create a new worksheet"""
        url = f'{MS_GRAPH_URL}/me/drive/items/{file_id}/workbook/worksheets/add'
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        body = {
            'name': sheet_name
        }
        response = requests.post(url, headers=headers, json=body)
        return response.status_code in (200, 201)

    def get_last_empty_row(worksheet_id):
        """Get the first empty row index using usedRange()."""
        url = f"{MS_GRAPH_URL}/me/drive/items/{file_id}/workbook/worksheets/{worksheet_id}/usedRange()"
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            raise Exception(f"Failed to read worksheet: {response.text}")

        data = response.json()
        # If the sheet is empty, rowCount may not be present, so default to 0
        last_row = data.get("rowCount", 0)
        # print(last_row)
        return last_row

    try:
        # Get current date for sheet naming
        current_date = datetime.now().strftime('%Y-%m-%d')
        # print(current_date)

        # Get existing worksheets
        worksheets = get_worksheets()

        # Find or create dated sheet
        target_sheet = next(
            (ws for ws in worksheets if ws['name'] == current_date), None)

        if not target_sheet:
            if not create_new_sheet(current_date):
                raise Exception("Failed to create new worksheet")
            target_sheet = next(ws for ws in get_worksheets()
                                if ws['name'] == current_date)

        # Convert data to API format
        if isinstance(data, dict):
            values = [[str(value) for value in data.values()]]
        elif isinstance(data, list):
            values = [data]
        else:
            raise ValueError("Data must be a dictionary or list")

        # Get first empty row
        last_row = get_last_empty_row(target_sheet['name'])

        # Calculate address range based on data length
        num_cols = len(values[0])
        start_col = 'A'
        end_col = string.ascii_uppercase[num_cols - 1]  # Supports up to 26 columns
        row_num = last_row + 1
        address_range = f"{start_col}{row_num}:{end_col}{row_num}"

        # Wrap worksheet name in single quotes and encode
        quoted_sheet_name = f"'{target_sheet['name']}'"
        # print(f"Adding data to {quoted_sheet_name} at {address_range}")
        url = f"{MS_GRAPH_URL}/me/drive/items/{file_id}/workbook/worksheets/{target_sheet['name']}/range(address='{address_range}')"
        # print(f"Writing to range: {address_range} with values: {values}")
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        body = {'values': values}

        response = requests.patch(url, headers=headers, json=body)
        # print(response.status_code, response.text)
        return response.status_code in (200, 201)

    except Exception as e:
        print(f"Error adding data to Excel file: {str(e)}")
        return False


if __name__ == "__main__":
    access_token = get_access_token(
        scopes=["Files.ReadWrite.All", "Sites.ReadWrite.All"])
    onedrive = OneDrive(access_token)

    file_id = onedrive.get_item_id(
        "C:/Users/martinko/OneDrive - MSFT/Documents/spymypc.xlsx")
    # print(file_id)
    add_data_to_dated_sheet(
        file_id, ["John Doe", 30, "New York"], access_token)
