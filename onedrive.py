import requests
import json
from ms_graph import MS_GRAPH_URL
from datetime import datetime, timedelta

class OneDrive:
    """
    Interface to interact with the OneDrive using Microsoft Graph API.
    ### Parameters:
    - access_token (str): The access token to use for authentication.
    """
    def __init__(self, access_token):
        self.access_token = access_token
        
    def get_item_id(self, file_path):
        """
        Returns the item ID for the provided file path.
        ### Parameters:
        - file_path (str): The file path to get the item ID for.
        """
        # Remove C:\Users\{User}\OneDrive - MSFT\ from the file path
        file_path = get_root_path(file_path)
        
        # Get the item ID of the file
        url = f"{MS_GRAPH_URL}/me/drive/root:/{file_path}:/?$select=id"
        headers = {
            "Authorization": f"Bearer {self.access_token}"
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        # Return the item ID
        return response.json()["id"]
        
    def create_sharing_link(self, file_path, type="view", scope="anonymous", expiration=(datetime.utcnow() + timedelta(weeks=1)).isoformat() + "Z", password=None):
        """
        Returns a sharing link for the provided file path.
        ### Parameters:
        - file_path (str): The file path to get the sharing link for
        - type (str): The type of link to create (view or edit)
        - scope (str): The scope of the link to create (anonymous or organization
        - expiration (str): The expiration date and time of the link (optional, defaults to 1 week)
        - password (str): The password for the link (optional)
        """
        # Get the item ID of the file
        item_id = self.get_item_id(file_path)
        
        # Get the sharing link for the file
        url = f"{MS_GRAPH_URL}/me/drive/items/{item_id}/createLink"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        data = {
            "type": type, # view or edit
            "scope": scope, # anonymous or organization
            "expirationDateTime": expiration, # optional
            "password": password, # optional
        }
        response = requests.post(url, headers=headers, data=json.dumps(data))
        response.raise_for_status()
        return response.json()["link"]["webUrl"]
    
def get_root_path(file_path):
    """
    Returns the root path of the OneDrive file.
    ### Parameters:
    - file_path (str): The file path to get the root path for.
    """
    return file_path.replace("C:/Users/martinko/OneDrive - MSFT/", "")