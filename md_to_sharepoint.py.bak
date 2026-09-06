#!/usr/bin/env python3
"""
md_to_sharepoint.py
Push an Obsidian / Markdown note to SharePoint Online as a modern site page.

Requirements:
    pip install requests msal markdown pyyaml

Usage:
    python md_to_sharepoint.py \
        --file "m28-a02.md" \
        --site "https://contoso.sharepoint.com/sites/lab-notes" \
        --tenant-id "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" \
        --client-id "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"

Obsidian integration (Shell Commands plugin):
    Command: python /path/to/md_to_sharepoint.py --file "{{file_path:absolute}}" --site "..." --tenant-id "..." --client-id "..."
"""

import argparse
import os
import re
import json
import sys

try:
    import requests
    import msal
    import markdown
    import yaml
except ImportError as e:
    print(f"Missing dependency: {e.name}")
    print("Install: pip install requests msal markdown pyyaml")
    sys.exit(1)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def parse_frontmatter(content: str):
    """Extract YAML frontmatter and body from markdown."""
    if content.startswith("---"):
        parts = content.split("---", 2)
        if len(parts) >= 3:
            try:
                fm = yaml.safe_load(parts[1])
                return fm if fm else {}, parts[2].strip()
            except yaml.YAMLError:
                pass
    return {}, content


def md_to_html(md_content: str) -> str:
    """Convert markdown to HTML with fenced code & table support."""
    md = markdown.Markdown(extensions=["fenced_code", "tables"])
    html = md.convert(md_content)
    return html


def get_site_id(site_url: str, token: str) -> str:
    """Get SharePoint site ID from URL via Microsoft Graph."""
    parts = site_url.split("/")
    hostname = parts[2]
    site_path = "/".join(parts[3:])
    url = f"{GRAPH_BASE}/sites/{hostname}:/{site_path}"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    return resp.json()["id"]


def get_access_token(tenant_id: str, client_id: str, cache_path: str = "token_cache.bin") -> str:
    """Authenticate via MSAL device code flow (interactive, no secret needed)."""
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    cache = msal.SerializableTokenCache()
    if os.path.exists(cache_path):
        cache.deserialize(open(cache_path, "r").read())

    app = msal.PublicClientApplication(client_id, authority=authority, token_cache=cache)

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(["Sites.ReadWrite.All"], account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=["Sites.ReadWrite.All"])
    if "user_code" not in flow:
        raise RuntimeError("Failed to create device flow. Check tenant/client-id.")

    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")

    with open(cache_path, "w") as f:
        f.write(cache.serialize())

    return result["access_token"]


def create_page(site_id: str, title: str, html_content: str, token: str, page_name: str = None):
    """Create a SharePoint modern page via Graph API."""
    if not page_name:
        page_name = re.sub(r"[^\w\-]", "-", title).lower() + ".aspx"
    if not page_name.endswith(".aspx"):
        page_name += ".aspx"

    # Strip Obsidian wiki-links [[...]] → plain text
    html_content = re.sub(r"\[\[([^\]]+)\]\]", r"\1", html_content)
    # Remove stray "undefined" artifacts
    html_content = html_content.replace("undefined", "")

    payload = {
        "@odata.type": "#microsoft.graph.sitePage",
        "name": page_name,
        "title": title,
        "pageLayout": "article",
        "showRecommendedPages": False,
        "canvasLayout": {
            "horizontalSections": [
                {
                    "layout": "oneColumn",
                    "id": "1",
                    "emphasis": "none",
                    "columns": [
                        {
                            "id": "1",
                            "width": 12,
                            "webparts": [
                                {
                                    "@odata.type": "#microsoft.graph.textWebPart",
                                    "id": "1",
                                    "innerHtml": html_content,
                                }
                            ],
                        }
                    ],
                }
            ]
        },
    }

    url = f"{GRAPH_BASE}/sites/{site_id}/pages"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.post(url, headers=headers, json=payload)

    if resp.status_code >= 400:
        try:
            err = resp.json()
            print(f"Graph API error response: {json.dumps(err, indent=2)}")
        except Exception:
            print(f"HTTP {resp.status_code}: {resp.text}")
    resp.raise_for_status()
    return resp.json()


def main():
    parser = argparse.ArgumentParser(description="Push Markdown note to SharePoint as a modern page")
    parser.add_argument("--file", "-f", required=True, help="Path to markdown file")
    parser.add_argument("--site", "-s", required=True, help="SharePoint site URL")
    parser.add_argument("--tenant-id", "-t", required=True, help="Azure AD Tenant ID")
    parser.add_argument("--client-id", "-c", required=True, help="Azure AD App Client ID")
    parser.add_argument("--page-name", "-p", help="Target page file name (default: slugified title.aspx)")
    parser.add_argument("--title", help="Page title (default: first H1 or filename)")
    args = parser.parse_args()

    with open(args.file, "r", encoding="utf-8") as f:
        raw = f.read()

    frontmatter, body = parse_frontmatter(raw)

    # Determine title
    title = args.title
    if not title:
        h1_match = re.search(r"^#\s+(.+)$", body, re.MULTILINE)
        if h1_match:
            title = h1_match.group(1).strip()
        else:
            title = os.path.splitext(os.path.basename(args.file))[0]

    # Convert markdown body to HTML
    html_body = md_to_html(body)

    # Prepend tags if frontmatter has them
    tags = frontmatter.get("tags", [])
    if isinstance(tags, str):
        tags = [tags]
    if tags:
        tag_html = (
            '<p style="color:#555;font-size:0.85em;margin-bottom:12px;">'
            + " ".join(
                f'<span style="background:#f0f0f0;padding:2px 8px;border-radius:10px;margin-right:4px;">{t}</span>'
                for t in tags
            )
            + "</p>"
        )
        html_body = tag_html + html_body

    print("🔐 Authenticating with Microsoft Graph (device code flow)...")
    token = get_access_token(args.tenant_id, args.client_id)

    print("📍 Resolving SharePoint site...")
    site_id = get_site_id(args.site, token)

    print(f"📝 Creating page "{title}"...")
    page = create_page(site_id, title, html_body, token, args.page_name)

    web_url = page.get("webUrl", "N/A")
    print(f"✅ Done! Page published at: {web_url}")


if __name__ == "__main__":
    main()
