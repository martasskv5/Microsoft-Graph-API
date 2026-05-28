from typing import Any

from fastapi import FastAPI, HTTPException, Request
from pydantic import BaseModel, Field

from emails import send_email
from excel import add_data_to_dated_sheet
from ms_graph import complete_auth_code_flow, get_access_token, start_auth_code_flow
from onedrive import OneDrive

app = FastAPI(title="Local Microsoft Graph API Bridge", version="1.0.0")


class LoginRequest(BaseModel):
    user: str | None = None
    scopes: list[str] = Field(default_factory=lambda: ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All", "Mail.Send"])


class LoginStartRequest(BaseModel):
    user: str | None = None
    scopes: list[str] = Field(default_factory=lambda: ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All", "Mail.Send"])
    redirect_uri: str = "http://localhost:8787/auth/callback"


class SendEmailRequest(BaseModel):
    sender_email: str
    recipient_emails: str | list[str]
    subject: str
    body: str
    user: str | None = None


class SharingLinkRequest(BaseModel):
    file_path: str
    type: str = "view"
    scope: str = "anonymous"
    expiration: str | None = None
    password: str | None = None
    user: str | None = None


class ResolveFileIdRequest(BaseModel):
    file_path: str
    user: str | None = None


class ExcelAppendRequest(BaseModel):
    file_id: str
    data: dict[str, Any] | list[Any]
    user: str | None = None


def _get_cached_token_or_401(scopes: list[str], user: str | None) -> str:
    token = get_access_token(scopes=scopes, user=user, allow_interactive=False)
    if not token:
        raise HTTPException(
            status_code=401,
            detail="No cached login found or token expired. Call POST /auth/login once to sign in and cache tokens.",
        )
    return token


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/auth/login")
def auth_login(payload: LoginRequest) -> dict[str, str]:
    token = get_access_token(scopes=payload.scopes, user=payload.user, allow_interactive=True)
    if not token:
        raise HTTPException(status_code=401, detail="Interactive login failed")
    return {"status": "ok", "message": "Login successful and token cache created."}


@app.post("/auth/login/start")
def auth_login_start(payload: LoginStartRequest) -> dict[str, str]:
    try:
        flow = start_auth_code_flow(
            scopes=payload.scopes,
            user=payload.user,
            redirect_uri=payload.redirect_uri,
        )
        return {
            "status": "ok",
            "auth_uri": flow["auth_uri"],
            "message": "Open auth_uri in a browser to complete login.",
        }
    except Exception as ex:
        raise HTTPException(status_code=500, detail=str(ex)) from ex


@app.get("/auth/callback")
def auth_callback(request: Request) -> dict[str, str]:
    try:
        auth_response = {k: v for k, v in request.query_params.items()}
        complete_auth_code_flow(auth_response)
        return {
            "status": "ok",
            "message": "Authentication completed and token cache saved.",
        }
    except Exception as ex:
        raise HTTPException(status_code=400, detail=str(ex)) from ex


@app.get("/auth/status")
def auth_status(user: str | None = None) -> dict[str, bool]:
    token = get_access_token(scopes=["User.Read"], user=user, allow_interactive=False)
    return {"authenticated": bool(token)}


@app.post("/email/send")
def email_send(payload: SendEmailRequest) -> dict[str, str]:
    required_scopes = ["User.Read", "Mail.Send"]

    # Mail.Send.Shared is only needed when sending as a mailbox different from the signed-in user.
    if payload.user and payload.sender_email.lower() != payload.user.lower():
        required_scopes.append("Mail.Send.Shared")

    token = _get_cached_token_or_401(required_scopes, payload.user)
    sent = send_email(
        sender_email=payload.sender_email,
        recipient_emails=payload.recipient_emails,
        subject=payload.subject,
        body=payload.body,
        access_token=token,
    )
    if not sent:
        raise HTTPException(status_code=502, detail="Microsoft Graph email send failed")
    return {"status": "sent"}


@app.post("/onedrive/file-id")
def onedrive_file_id(payload: ResolveFileIdRequest) -> dict[str, str]:
    token = _get_cached_token_or_401(["Files.Read.All", "Files.ReadWrite.All", "User.Read"], payload.user)
    client = OneDrive(token)
    try:
        file_id = client.get_item_id(payload.file_path)
        return {"file_id": file_id}
    except Exception as ex:
        raise HTTPException(status_code=502, detail=str(ex)) from ex


@app.post("/onedrive/sharing-link")
def onedrive_sharing_link(payload: SharingLinkRequest) -> dict[str, str]:
    token = _get_cached_token_or_401(["Files.ReadWrite.All", "User.Read"], payload.user)
    client = OneDrive(token)

    try:
        link = client.create_sharing_link(
            file_path=payload.file_path,
            type=payload.type,
            scope=payload.scope,
            expiration=payload.expiration,
            password=payload.password,
        )
        return {"link": link}
    except Exception as ex:
        raise HTTPException(status_code=502, detail=str(ex)) from ex


@app.post("/excel/append-dated-row")
def excel_append_dated_row(payload: ExcelAppendRequest) -> dict[str, str]:
    token = _get_cached_token_or_401(["Files.ReadWrite.All", "Sites.ReadWrite.All", "User.Read"], payload.user)
    success = add_data_to_dated_sheet(payload.file_id, payload.data, access_token=token)
    if not success:
        raise HTTPException(status_code=502, detail="Microsoft Graph Excel append failed")
    return {"status": "appended"}
