"""Authentication helpers for Microsoft Graph using MSAL and optional browser flow."""

import os
import urllib.parse
from typing import Any

import msal
from dotenv import load_dotenv

try:
    from win11toast import toast as notification
except Exception:
    def notification(*args, **kwargs):
        # Notifications are optional and unavailable on non-Windows hosts.
        return None


load_dotenv()

MS_GRAPH_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = [
    "User.Read",
    "Files.ReadWrite.All",
    "Sites.Manage.All",
    "Sites.ReadWrite.All",
]

_AUTH_SESSIONS: dict[str, dict[str, Any]] = {}


def _validate_auth_config(client_id: str | None, client_secret: str | None, tenant_id: str | None) -> tuple[str, str, str]:
    """Validate required Azure app settings and return stripped values."""
    client_id = (client_id or "").strip()
    client_secret = (client_secret or "").strip()
    tenant_id = (tenant_id or "").strip()

    missing = []
    if not client_id:
        missing.append("CLIENT_ID")
    if not client_secret:
        missing.append("CLIENT_SECRET")
    if not tenant_id:
        missing.append("TENANT_ID")

    if missing:
        raise Exception(
            "Missing Azure auth configuration: "
            + ", ".join(missing)
            + ". Set these environment variables in Docker/Portainer."
        )

    return client_id, client_secret, tenant_id


def _get_cache_file(user: str | None) -> str:
    """Build a per-user MSAL cache path that works on Windows, Linux, and macOS."""
    local_app_data = os.getenv("LOCALAPPDATA")
    xdg_cache = os.getenv("XDG_CACHE_HOME")

    if local_app_data:
        base_cache_dir = local_app_data
    elif xdg_cache:
        base_cache_dir = xdg_cache
    else:
        base_cache_dir = os.path.join(os.path.expanduser("~"), ".cache")

    cache_dir = os.path.join(base_cache_dir, "martas")
    os.makedirs(cache_dir, exist_ok=True)

    normalized_user = (user or "default").strip().lower()
    safe_user = normalized_user.replace("@", "_").replace("/", "_").replace("\\", "_")
    return os.path.join(cache_dir, f"{safe_user}_msal_cache.bin")


def get_access_token(
    client_id: str | None = os.getenv("CLIENT_ID"),
    client_secret: str | None = os.getenv("CLIENT_SECRET"),
    tenant_id: str | None = os.getenv("TENANT_ID"),
    scopes: list[str] | None = None,
    user: str | None = None,
    allow_interactive: bool = True,
    redirect_uri: str = "http://localhost",
):
    """
    Returns an access token for the specified scopes using the provided client ID, client secret, and tenant ID.
    Caches the token on disk so repeated calls can renew silently.

    Args:
        client_id (str): The client ID of the Azure AD application.
        client_secret (str): The client secret of the Azure AD application.
        tenant_id (str): The tenant ID of the Azure AD application.
        scopes (list[str]): The scopes to request access for.
        user (str): Optional user principal name (email) to select a specific account in cache.
        allow_interactive (bool): If False, do not start browser sign-in flow.
        redirect_uri (str): Redirect URI configured in app registration.
    """
    client_id, client_secret, tenant_id = _validate_auth_config(client_id, client_secret, tenant_id)
    scopes = scopes or DEFAULT_SCOPES
    token_cache_file = _get_cache_file(user)

    token_cache = msal.SerializableTokenCache()
    if os.path.exists(token_cache_file):
        try:
            with open(token_cache_file, "r", encoding="utf-8") as f:
                token_cache.deserialize(f.read())
        except Exception:
            pass

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=token_cache,
    )

    user = user.strip().lower() if user else None
    accounts = app.get_accounts()
    account = None
    if user and accounts:
        for candidate in accounts:
            candidate_username = (candidate.get("username") or "").strip().lower()
            if candidate_username == user:
                account = candidate
                break
        if account is None:
            # If no exact username match was found, still use an available cached account.
            account = accounts[0]
    elif accounts:
        account = accounts[0]

    if account:
        result = app.acquire_token_silent(scopes, account=account)
        if result and "access_token" in result:
            return result["access_token"]

    if not allow_interactive:
        return None

    flow = app.initiate_auth_code_flow(scopes=scopes, redirect_uri=redirect_uri)
    callback_url = get_redirect_url(flow["auth_uri"]) if user is None else input(
        f"Please sign in as {user} and paste the redirect URL here:\n{flow['auth_uri']}\n"
    )

    if callback_url is None:
        notification(
            "Microsoft Graph Authentication Failed",
            "Please sign in to your Microsoft account in MS Edge.",
        )
        return None

    parsed_url = urllib.parse.urlparse(callback_url)
    auth_response = dict(urllib.parse.parse_qsl(parsed_url.query))
    result = app.acquire_token_by_auth_code_flow(flow, auth_response)

    if "access_token" in result:
        try:
            if token_cache.has_state_changed:
                with open(token_cache_file, "w", encoding="utf-8") as f:
                    f.write(token_cache.serialize())
        except Exception:
            pass
        return result["access_token"]

    raise Exception(f"Authentication failed: {result.get('error')}")


def start_auth_code_flow(
    client_id: str | None = os.getenv("CLIENT_ID"),
    client_secret: str | None = os.getenv("CLIENT_SECRET"),
    tenant_id: str | None = os.getenv("TENANT_ID"),
    scopes: list[str] | None = None,
    user: str | None = None,
    redirect_uri: str = "http://localhost:8787/auth/callback",
) -> dict[str, str]:
    """Start auth code flow and return sign-in URL for browser-based login."""
    client_id, client_secret, tenant_id = _validate_auth_config(client_id, client_secret, tenant_id)
    scopes = scopes or DEFAULT_SCOPES
    token_cache_file = _get_cache_file(user)

    token_cache = msal.SerializableTokenCache()
    if os.path.exists(token_cache_file):
        try:
            with open(token_cache_file, "r", encoding="utf-8") as f:
                token_cache.deserialize(f.read())
        except Exception:
            pass

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=token_cache,
    )

    flow = app.initiate_auth_code_flow(scopes=scopes, redirect_uri=redirect_uri)
    state = flow.get("state")
    if not state:
        raise Exception("Missing OAuth state from auth code flow")

    _AUTH_SESSIONS[state] = {
        "app": app,
        "flow": flow,
        "token_cache": token_cache,
        "token_cache_file": token_cache_file,
    }

    return {"auth_uri": flow["auth_uri"], "state": state}


def complete_auth_code_flow(auth_response: dict[str, str]) -> str:
    """Complete previously-started auth code flow and persist token cache."""
    state = auth_response.get("state")
    if not state:
        raise Exception("Missing state in auth response")

    session = _AUTH_SESSIONS.get(state)
    if not session:
        raise Exception("Auth session not found or expired")

    app = session["app"]
    flow = session["flow"]
    token_cache = session["token_cache"]
    token_cache_file = session["token_cache_file"]

    try:
        result = app.acquire_token_by_auth_code_flow(flow, auth_response)
    finally:
        _AUTH_SESSIONS.pop(state, None)

    if "access_token" not in result:
        raise Exception(f"Authentication failed: {result.get('error')}")

    try:
        if token_cache.has_state_changed:
            with open(token_cache_file, "w", encoding="utf-8") as f:
                f.write(token_cache.serialize())
    except Exception:
        pass

    return result["access_token"]

def get_redirect_url(auth_uri: str) -> str | None:
    """
    Returns the redirect URL after authenticating the user using the provided auth URI or None if not signed in. 
    User must be signed in to their Microsoft account in MS Edge.
    ### Parameters:
    - auth_uri (str): The auth URI to use for authentication.
    - user (str): The users email to authenticate as. If None, uses the default user.
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.edge.options import Options as EdgeOptions
        from selenium.webdriver.edge.service import Service
        from webdriver_manager.microsoft import EdgeChromiumDriverManager
    except Exception:
        webdriver = None

    if webdriver is None:
        print("Selenium/webdriver is not installed. Falling back to manual sign-in flow.")
        print("Open the following URL, sign in, and paste the redirect URL:")
        print("\n" + auth_uri + "\n")
        try:
            callback = input("Paste the full redirect URL here: ").strip()
            return callback if callback else None
        except Exception:
            return None

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