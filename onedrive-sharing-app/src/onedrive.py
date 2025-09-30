import os
import requests
from dotenv import load_dotenv
from requests_oauthlib import OAuth2Session

load_dotenv()

class OneDriveAPI:
    def __init__(self):
        self.client_id = os.getenv('CLIENT_ID')
        self.redirect_uri = os.getenv('REDIRECT_URI')
        self.tenant_id = os.getenv('TENANT_ID')
        self.access_token = None
        self.oauth = OAuth2Session(self.client_id, redirect_uri=self.redirect_uri, scope=["Files.ReadWrite.All"])

    def authenticate(self):
        authorization_base_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/authorize"
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        authorization_url, state = self.oauth.authorization_url(authorization_base_url)
        print(f"Please go to {authorization_url} and authorize access.")

        redirect_response = input("Paste the full redirect URL here: ")
        self.oauth.fetch_token(token_url, client_id=self.client_id, authorization_response=redirect_response)

        self.access_token = self.oauth.token["access_token"]

    def generate_sharing_link(self, item_id):
        if not self.access_token:
            self.authenticate()

        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{item_id}/createLink"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        body = {
            "type": "view",  # or "edit" based on requirement
            "scope": "anonymous"  # or "organization"
        }

        response = requests.post(url, headers=headers, json=body)

        if response.status_code == 200:
            return response.json().get('link').get('webUrl')
        else:
            raise Exception(f"Error generating sharing link: {response.text}")