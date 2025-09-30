"""
Custom module for authenticating with Microsoft Graph using the Microsoft Authentication Library (MSAL) and Selenium.
"""

import msal
import os
import urllib.parse
import pickle
import time
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from win11toast import toast as notification
load_dotenv()

MS_GRAPH_URL = "https://graph.microsoft.com/v1.0"
"""
Microsoft Graph API URL.
"""

def get_access_token(
    client_id: str | None = os.getenv("CLIENT_ID"),
    client_secret: str | None = os.getenv("CLIENT_SECRET"),
    tenant_id: str | None = os.getenv("TENANT_ID"),
    scopes: list[str] = ["User.Read", "Files.ReadWrite.All", "Sites.Manage.All", "Sites.ReadWrite.All"],
    user : str | None = None
):
    """
    Returns an access token for the specified scopes using the provided client ID, client secret, and tenant ID.
    Caches the token in a pickle file to avoid repeated authentication.
    
    Args:
        client_id (str): The client ID of the Azure AD application.
        client_secret (str): The client secret of the Azure AD application.
        tenant_id (str): The tenant ID of the Azure AD application.
        scopes (list[str]): The scopes to request access for.
        user (str): The users email to authenticate as. If None, uses the default user.
    """
    # print(client_id)
    local_app_data = os.getenv('LOCALAPPDATA') or ""
    cache_dir = os.path.join(local_app_data, "martas") if local_app_data else None
    if cache_dir:
        try:
            os.makedirs(cache_dir, exist_ok=True)
        except Exception:
            cache_dir = None

    token_cache_file = os.path.join(cache_dir, f"{user}_msal_cache.bin") if user and cache_dir else "notfound.pickle"

    # MSAL token cache (persisted) to avoid asking user to re-authenticate repeatedly
    token_cache = msal.SerializableTokenCache()
    if os.path.exists(token_cache_file):
        try:
            with open(token_cache_file, "r", encoding="utf-8") as f:
                token_cache.deserialize(f.read())
        except Exception:
            # ignore cache load errors and continue with empty cache
            pass

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=token_cache
    )

    # Try to get a token silently from the cache first
    accounts = app.get_accounts()
    account = None
    if user and accounts:
        # try to find matching account by username
        for a in accounts:
            if a.get('username') == user:
                account = a
                break
    elif accounts:
        account = accounts[0]

    if account:
        result = app.acquire_token_silent(scopes, account=account)
        if result and "access_token" in result:
            return result["access_token"]

    # No token in cache or silent acquire failed — start interactive auth
    flow = app.initiate_auth_code_flow(
        scopes=scopes,
        redirect_uri="http://localhost"
    )

    callback_url = get_redirect_url(flow['auth_uri']) if user is None else input(
        f"Please sign in as {user} and paste the redirect URL here:\n{flow['auth_uri']}\n"
    )
    if callback_url is None:
        notification("Microsoft Graph Authentication Failed", "Please sign in to your Microsoft account in MS Edge.")
        return None
    parsed_url = urllib.parse.urlparse(callback_url)
    auth_response = dict(urllib.parse.parse_qsl(parsed_url.query))

    result = app.acquire_token_by_auth_code_flow(flow, auth_response)

    if "access_token" in result:
        # Persist MSAL token cache to file so acquire_token_silent can use it later
        try:
            if token_cache.has_state_changed:
                with open(token_cache_file, "w", encoding="utf-8") as f:
                    f.write(token_cache.serialize())
        except Exception:
            # non-fatal — continue
            pass
        return result["access_token"]
    raise Exception(f"Authentication failed: {result.get('error')}")

def get_redirect_url(auth_uri: str) -> str | None:
    """
    Returns the redirect URL after authenticating the user using the provided auth URI or None if not signed in. 
    User must be signed in to their Microsoft account in MS Edge.
    ### Parameters:
    - auth_uri (str): The auth URI to use for authentication.
    - user (str): The users email to authenticate as. If None, uses the default user.
    """
    # Create Edge options
    options = EdgeOptions()

    # Set profile directory if provided
    options.add_argument("--profile-directory=Default")

    # Try to initialize Edge driver automatically. If that fails (for example when offline
    # and webdriver_manager cannot download the driver), fall back to a manual flow where the
    # user opens the auth URL and pastes the redirected URL.
    try:
        service = Service(EdgeChromiumDriverManager().install(), log_path=os.path.devnull)
        browser = webdriver.Edge(service=service, options=options)

        try:
            browser.get(auth_uri)
            WebDriverWait(browser, 600).until(
                EC.url_contains("localhost/?code=")
            )

            if (not EC.url_contains("localhost/?code=")):
                notification("Microsoft Graph Authentication Failed", "Please sign in to your Microsoft account in MS Edge.", duration=10)
                return None
            return browser.current_url
        finally:
            if 'browser' in locals():
                browser.quit()
    except Exception:
        # Likely offline or driver download failed. Fall back to manual paste flow.
        print("Could not start Edge WebDriver automatically. If you're offline or the driver can't be downloaded,")
        print("please open the following URL in a browser, sign in, and paste the full redirect URL here when finished:")
        print("\n" + auth_uri + "\n")
        try:
            callback = input("Paste the full redirect URL here: ").strip()
            return callback if callback else None
        except Exception:
            return None
            
# if __name__ == "__main__":
#     print(get_access_token())