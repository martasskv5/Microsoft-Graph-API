import os
import webbrowser
from msal import ConfidentialClientApplication
from dotenv import load_dotenv

load_dotenv()

application_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
SCOPES = ["User.Read"]

client = ConfidentialClientApplication(
    client_id=application_id,
    client_credential=client_secret,
    authority="https://login.microsoftonline.com/organizations"
)
auth_url = client.get_authorization_request_url(SCOPES)
webbrowser.open(auth_url)