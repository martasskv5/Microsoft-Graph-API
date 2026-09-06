#!/usr/bin/env python3
"""
obsidian_to_sharepoint.py
Push the current Obsidian note to the correct SharePoint subject site.

Auto-detects:
  - Subject from parent folder name (e.g. "02 - CIT" -> CIT site)
  - School year from vault folder name (e.g. "SPSITknm 25-26" -> 2025-2026)

Usage:
    python obsidian_to_sharepoint.py --file "{{file_path:absolute}}"

Obsidian Shell Commands plugin:
    python C:\\path\\to\\obsidian_to_sharepoint.py --file "{{file_path:absolute}}"

Requirements:
    pip install requests msal markdown pyyaml

Auth overrides:
    --client-id, --client-secret, --tenant-id, --scopes

Example:
    python obsidian_to_sharepoint.py --file "{{file_path:absolute}}" \
      --client-id "..." --client-secret "..." \
      --scopes "User.Read,Files.ReadWrite.All,Sites.ReadWrite.All"
"""

import argparse
import base64
import datetime as dt
import hashlib
import io
import html
import math
import os
import re
import json
import sys
import shutil
import tempfile
from pathlib import Path
from urllib.parse import quote

try:
    import requests
    import markdown
    import yaml
    from PIL import Image, ImageDraw, ImageFont
except ImportError as e:
    print(f"Missing dependency: {e.name}")
    print("Install: pip install requests msal markdown pyyaml")
    sys.exit(1)

from ms_graph import get_access_token

# ============================================================
# CONFIGURATION — adjust to match your SharePoint sites
# ============================================================
TENANT_NAME = "0z6sw"                     # your SharePoint tenant
SCHOOL_ALIAS = "spsitknm"                 # school site alias
DEFAULT_TENANT_ID = os.getenv("TENANT_ID", "").strip()
DEFAULT_CLIENT_ID = os.getenv("O2S_ID", "").strip()
DEFAULT_CLIENT_SECRET = os.getenv("O2S_SECRET", "").strip()

# Map: Vault folder name -> SharePoint subject name & URL alias
# The alias is the last part of /sites/spsitknm-2025-2026-<alias>
SUBJECT_MAP = {
    "00 - ANJ":       {"sp_name": "ANJ",       "alias": "anj"},
    "01 - APM":       {"sp_name": "APM",       "alias": "apm"},
    "02 - Appslab":   {"sp_name": "Appslab",   "alias": "appslab"},
    "03 - CEQ":       {"sp_name": "CEQ",       "alias": "ceq"},
    "04 - CIT":       {"sp_name": "CIT",       "alias": "cit"},
    "05 - CME":       {"sp_name": "CME",       "alias": "cme"},
    "06 - CSX":       {"sp_name": "CSX",       "alias": "csx"},
    "07 - EKO":       {"sp_name": "EKO",       "alias": "eko"},
    "08 - ETS":       {"sp_name": "ETS",       "alias": "ets"},
    "09 - MAT":       {"sp_name": "MAT",       "alias": "mat"},
    "10 - SJL": {"sp_name": "SJL",       "alias": "sjl"},
    "11 - SXT":       {"sp_name": "SXT",       "alias": "sxt"},
    "12 - USP":       {"sp_name": "USP",       "alias": "usp"},
}

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE = os.path.join(Path.home(), ".obsidian_sp_cache.bin")
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff"}
EXCALIDRAW_FOLDER = Path("99 - Meta") / "Excalidraw"
INDEX_PAGE_TITLE = os.getenv("O2S_INDEX_PAGE", "Index")
INDEX_PAGE_DESCRIPTION = "Auto-generated page index"
FONT_CANDIDATES = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
)

# ============================================================
# HELPERS
# ============================================================

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


def default_page_name(title: str) -> str:
    """Build the default SharePoint page file name from a title."""
    page_name = re.sub(r"[^\w\-]", "-", title).lower()
    if not page_name.endswith(".aspx"):
        page_name += ".aspx"
    return page_name


def parse_date_string(value: object) -> dt.datetime | dt.date | None:
    """Parse a date-like frontmatter or string value into a date or datetime."""
    if value is None:
        return None

    if isinstance(value, dt.datetime):
        return value
    if isinstance(value, dt.date):
        return value

    text = str(value).strip()
    if not text:
        return None

    # Try datetime patterns first (date + time, e.g. 2026-06-03 16-14 or 2026-06-03T16:14)
    dt_patterns = [
        r"(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})[ T](?P<hour>\d{2})[-:](?P<minute>\d{2})[-:]?(?P<second>\d{2})?",
        r"(?P<year>\d{4})/(?P<month>\d{2})/(?P<day>\d{2})[ T](?P<hour>\d{2})[-:](?P<minute>\d{2})[-:]?(?P<second>\d{2})?",
    ]
    for pattern in dt_patterns:
        match = re.search(pattern, text)
        if match:
            try:
                year = int(match.group("year"))
                month = int(match.group("month"))
                day = int(match.group("day"))
                hour = int(match.group("hour"))
                minute = int(match.group("minute"))
                second = int(match.group("second") or 0)
                return dt.datetime(year, month, day, hour, minute, second)
            except ValueError:
                continue

    # Fall back to date-only patterns
    for pattern in (
        r"(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})",
        r"(?P<year>\d{4})/(?P<month>\d{2})/(?P<day>\d{2})",
    ):
        match = re.search(pattern, text)
        if match:
            try:
                return dt.date(int(match.group("year")), int(match.group("month")), int(match.group("day")))
            except ValueError:
                return None

    try:
        return dt.date.fromisoformat(text[:10])
    except Exception:
        return None


def extract_note_date(frontmatter: dict, title: str) -> dt.datetime | dt.date | None:
    """Find the best date for a note using filename and frontmatter."""
    title_date = parse_date_string(title)
    if title_date:
        return title_date

    for key in ("date", "published", "publish_date", "publish-date"):
        frontmatter_date = parse_date_string(frontmatter.get(key))
        if frontmatter_date:
            return frontmatter_date

    return None


def extract_page_date(page: dict) -> dt.datetime | dt.date | None:
    """Extract a page date from Graph metadata for sorting."""
    description_date = parse_date_string(page.get("description"))
    if description_date:
        return description_date

    title_date = parse_date_string(page.get("title") or page.get("name"))
    if title_date:
        return title_date

    return None


def page_sort_key(page: dict) -> tuple:
    """Sort undated pages first, then dated pages newest-first."""
    page_date = extract_page_date(page)
    title = (page.get("title") or page.get("name") or "").lower()
    if not page_date:
        return (0, 0, title)
    # Use timestamp for datetime (includes time), toordinal for plain date
    if isinstance(page_date, dt.datetime):
        return (1, -page_date.timestamp(), title)
    return (1, -page_date.toordinal(), title)


def safe_slug(value: str) -> str:
    """Convert an arbitrary string into a stable file-friendly slug."""
    slug = re.sub(r"[^a-zA-Z0-9._-]+", "-", value).strip("-")
    return slug.lower() or "item"


def find_vault_root(file_path: str) -> Path | None:
    """Locate the Obsidian vault root from a note path."""
    path = Path(file_path).resolve()
    for parent in path.parents:
        if (parent / ".obsidian").is_dir():
            return parent
    return None


def normalize_input_file_path(raw_path: str) -> Path:
    """Recover a real filesystem path from a Shell Commands expanded argument."""
    candidate = raw_path.strip().strip('"').strip("'")

    # Undo common shell escaping that shows up in copied command arguments.
    candidate = candidate.replace("\\/", "/")
    candidate = re.sub(r"\\([\\/ ._\-()\[\]{}])", r"\1", candidate)
    candidate = candidate.replace("\\", "")

    # Shell Commands can sometimes inject an escaped absolute path into another path.
    # Prefer the last real /home/... segment when present.
    home_index = candidate.rfind("/home/")
    if home_index != -1:
        candidate = candidate[home_index:]

    return Path(candidate)


def load_font(size: int):
    """Load a readable font for rendering Excalidraw text."""
    for candidate in FONT_CANDIDATES:
        if os.path.exists(candidate):
            return ImageFont.truetype(candidate, size=size)
    return ImageFont.load_default()


def _lz_get_base_value(char: str) -> int:
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    index = alphabet.find(char)
    if index < 0:
        raise ValueError(f"Invalid LZ-string character: {char!r}")
    return index


def lzstring_decompress_from_encoded_uri_component(compressed: str) -> str:
    """Decompress Excalidraw's lz-string encoded payload."""
    if compressed is None:
        return ""
    if compressed == "":
        return ""

    compressed = compressed.replace(" ", "+")
    length = len(compressed)

    def get_next_value(index: int) -> int:
        return _lz_get_base_value(compressed[index])

    dictionary: list[str] = ["", "", ""]
    enlarge_in = 4
    dict_size = 4
    num_bits = 3
    entry = ""
    result: list[str] = []

    data_val = get_next_value(0)
    data_position = 32
    data_index = 1

    def read_bits(n: int) -> int:
        nonlocal data_val, data_position, data_index
        bits = 0
        power = 1
        max_power = 2**n
        while power != max_power:
            resb = data_val & data_position
            data_position >>= 1
            if data_position == 0:
                data_position = 32
                if data_index < length:
                    data_val = get_next_value(data_index)
                else:
                    data_val = 0
                data_index += 1
            if resb > 0:
                bits |= power
            power <<= 1
        return bits

    next_value = read_bits(2)
    if next_value == 0:
        c = chr(read_bits(8))
    elif next_value == 1:
        c = chr(read_bits(16))
    elif next_value == 2:
        return ""
    else:
        raise ValueError("Unsupported Excalidraw payload header")

    dictionary.append(c)
    w = c
    result.append(c)

    while True:
        if data_index > length:
            return ""

        c = read_bits(num_bits)

        if c == 0:
            dictionary.append(chr(read_bits(8)))
            c = dict_size
            dict_size += 1
            enlarge_in -= 1
        elif c == 1:
            dictionary.append(chr(read_bits(16)))
            c = dict_size
            dict_size += 1
            enlarge_in -= 1
        elif c == 2:
            return "".join(result)

        if enlarge_in == 0:
            enlarge_in = 2**num_bits
            num_bits += 1

        if c < len(dictionary):
            entry = dictionary[c]
        elif c == dict_size:
            entry = w + w[0]
        else:
            raise ValueError("Bad Excalidraw compressed payload")

        result.append(entry)
        dictionary.append(w + entry[0])
        dict_size += 1
        enlarge_in -= 1
        w = entry

        if enlarge_in == 0:
            enlarge_in = 2**num_bits
            num_bits += 1


def _extract_excalidraw_json(excalidraw_path: Path) -> dict:
    text = excalidraw_path.read_text(encoding="utf-8")
    match = re.search(r"```compressed-json\s*(.*?)\s*```", text, re.S)
    if not match:
        raise RuntimeError(f"Could not find compressed drawing data in {excalidraw_path}")

    compressed = "".join(match.group(1).split())
    json_text = lzstring_decompress_from_encoded_uri_component(compressed)
    return json.loads(json_text)


def _parse_stroke_color(value: str | None, default: str = "#000000") -> tuple[int, int, int, int]:
    color = value or default
    if not color.startswith("#") or len(color) not in {4, 7, 9}:
        color = default
    if len(color) == 4:
        color = "#" + "".join(ch * 2 for ch in color[1:])
    if len(color) == 7:
        color += "ff"
    red = int(color[1:3], 16)
    green = int(color[3:5], 16)
    blue = int(color[5:7], 16)
    alpha = int(color[7:9], 16)
    return red, green, blue, alpha


def render_excalidraw_to_png(excalidraw_path: Path, output_path: Path) -> Path:
    """Render an Excalidraw file to PNG using the external exporter script.

    This delegates rendering to `excalidraw_to_png.ExcalidrawExporter.export_with_cli`.
    Falls back to raising a RuntimeError on failure so callers can handle it.
    """
    try:
        import excalidraw_to_png as exporter_module
        ExcalidrawExporter = exporter_module.ExcalidrawExporter
        find_vault_root_exp = getattr(exporter_module, "find_vault_root", None)
        has_playwright = getattr(exporter_module, "sync_playwright", None) is not None
    except Exception as exc:
        raise RuntimeError(f"Excalidraw exporter not available: {exc}")

    # Detect vault root properly — not just the Excalidraw file's parent folder
    if find_vault_root_exp:
        vault_root = find_vault_root_exp(excalidraw_path)
    else:
        vault_root = find_vault_root(str(excalidraw_path)) or excalidraw_path.parent

    exporter = ExcalidrawExporter(
        vault_path=str(vault_root),
        output_dir=str(output_path.parent),
    )

    def try_export(path: Path) -> Path | None:
        # Try CLI first
        try:
            res = exporter.export_with_cli(path, output_path)
            if res:
                return res
        except Exception:
            pass

        # Then try Playwright if available
        if has_playwright:
            try:
                res = exporter.export_with_playwright(path, output_path)
                if res:
                    return res
            except Exception:
                pass
        return None

    # 1) Try original source file
    result = try_export(excalidraw_path)
    if result:
        return result

    # 2) Fallback: extract JSON locally and retry with a temp .excalidraw
    try:
        scene_data = _extract_excalidraw_json(excalidraw_path)
    except Exception as exc:
        raise RuntimeError(f"Excalidraw export failed and JSON extraction failed: {exc}")

    with tempfile.NamedTemporaryFile(mode="w", suffix=".excalidraw", delete=False) as tf:
        json.dump(scene_data, tf)
        temp_json_path = Path(tf.name)

    try:
        retry = try_export(temp_json_path)
        if retry:
            return retry
        raise RuntimeError(
            f"Excalidraw exporter failed for original and generated JSON: {excalidraw_path}. "
            "Ensure `npx excalidraw-brute-export-cli` is available or install Playwright (pip) and run `playwright install firefox`."
        )
    finally:
        try:
            temp_json_path.unlink()
        except Exception:
            pass


def find_obsidian_asset_path(note_path: str, target: str) -> Path | None:
    """Resolve an Obsidian embed target to a local file path."""
    note_file = Path(note_path).resolve()
    note_dir = note_file.parent
    vault_root = find_vault_root(note_path)

    target = target.strip()
    if not target:
        return None

    direct = Path(target)
    candidates: list[Path] = []

    if direct.is_absolute():
        candidates.append(direct)
    else:
        candidates.append(note_dir / direct)
        candidates.append(note_dir / "_img" / direct.name)
        if vault_root:
            candidates.append(vault_root / direct)
            candidates.append(vault_root / "_img" / direct.name)
            candidates.append(vault_root / "99 - Meta" / "02 - Attachments" / direct.name)
            candidates.append(vault_root / EXCALIDRAW_FOLDER / f"{direct.name}.md")
            candidates.append(vault_root / EXCALIDRAW_FOLDER / f"{direct.name}.excalidraw.md")

    for candidate in candidates:
        if candidate.exists():
            return candidate

    if vault_root:
        basename = direct.name
        for candidate in vault_root.rglob(basename):
            if candidate.is_file():
                return candidate

    return None


def build_remote_asset_name(page_name: str, source_key: str, local_path: Path, source_name: str | None = None) -> str:
    """Build a deterministic remote file name for an uploaded asset."""
    page_slug = safe_slug(Path(page_name).stem if page_name else "page")
    source_slug = safe_slug(source_name or local_path.stem)
    digest = hashlib.sha1(source_key.encode("utf-8")).hexdigest()[:10]
    suffix = local_path.suffix.lower() or ".png"
    return f"{page_slug}-{source_slug}-{digest}{suffix}"


def build_local_asset_name(page_name: str, source_key: str, local_path: Path, source_name: str | None = None) -> str:
    """Build a deterministic file name for a locally exported asset."""
    return build_remote_asset_name(page_name, source_key, local_path, source_name)


def upload_file_to_site(site_id: str, local_path: Path, remote_name: str, token: str) -> dict:
    """Upload a binary file into the site's default document library."""
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{quote(remote_name)}:/content"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"}
    with open(local_path, "rb") as file_handle:
        resp = requests.put(url, headers=headers, data=file_handle)

    if resp.status_code in (401, 403):
        raise RuntimeError(
            "Microsoft Graph rejected the file upload with an authorization error. "
            "Check Files.ReadWrite.All and Sites.ReadWrite.All permissions and tenant admin consent."
        )
    resp.raise_for_status()
    return resp.json()


def save_file_to_local_assets(local_path: Path, output_assets_dir: Path, asset_name: str) -> Path:
    """Copy a binary file into a local export assets folder."""
    output_assets_dir.mkdir(parents=True, exist_ok=True)
    destination = output_assets_dir / asset_name
    shutil.copy2(local_path, destination)
    return destination


def is_excalidraw_source(path: Path) -> bool:
    """Check whether a resolved asset path is an Excalidraw source file."""
    return path.name.endswith(".excalidraw.md")


def render_and_upload_obsidian_asset(
    site_id: str,
    token: str,
    note_path: str,
    page_name: str,
    target: str,
    upload_cache: dict[str, dict],
) -> dict | None:
    """Resolve, render if needed, upload, and cache an Obsidian embedded asset."""
    source_path = find_obsidian_asset_path(note_path, target)
    if not source_path:
        return None

    source_key = str(source_path.resolve())
    if source_key in upload_cache:
        return upload_cache[source_key]

    upload_path = source_path
    temp_path: Path | None = None

    try:
        if is_excalidraw_source(source_path):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_handle:
                temp_path = Path(temp_handle.name)
            render_excalidraw_to_png(source_path, temp_path)
            upload_path = temp_path

        remote_name = build_remote_asset_name(page_name, source_key, upload_path, source_path.stem)
        uploaded = upload_file_to_site(site_id, upload_path, remote_name, token)
        upload_cache[source_key] = uploaded
        return uploaded
    finally:
        if temp_path and temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass


def render_and_store_obsidian_asset_locally(
    note_path: str,
    page_name: str,
    target: str,
    output_assets_dir: Path,
    export_cache: dict[str, str],
) -> str | None:
    """Resolve, render if needed, and copy an embedded asset into the local dry-run folder."""
    source_path = find_obsidian_asset_path(note_path, target)
    if not source_path:
        return None

    source_key = str(source_path.resolve())
    if source_key in export_cache:
        return export_cache[source_key]

    asset_path = source_path
    temp_path: Path | None = None

    try:
        if is_excalidraw_source(source_path):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_handle:
                temp_path = Path(temp_handle.name)
            render_excalidraw_to_png(source_path, temp_path)
            asset_path = temp_path

        asset_name = build_local_asset_name(page_name, source_key, asset_path, source_path.stem)
        copied = save_file_to_local_assets(asset_path, output_assets_dir, asset_name)
        relative_path = f"assets/{copied.name}"
        export_cache[source_key] = relative_path
        return relative_path
    finally:
        if temp_path and temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass


def replace_obsidian_embeds(
    md_content: str,
    site_id: str,
    token: str,
    note_path: str,
    page_name: str,
) -> str:
    """Replace Obsidian image embeds with SharePoint-hosted image tags."""
    upload_cache: dict[str, dict] = {}

    def replace_embed(match: re.Match[str]) -> str:
        inner = match.group(1).strip()
        target, _, size_or_caption = inner.partition("|")
        target = target.strip()
        size_or_caption = size_or_caption.strip()

        uploaded = render_and_upload_obsidian_asset(site_id, token, note_path, page_name, target, upload_cache)
        if not uploaded:
            return match.group(0)

        web_url = uploaded.get("webUrl") or uploaded.get("@microsoft.graph.downloadUrl")
        if not web_url:
            return match.group(0)

        width_style = "max-width:100%;height:auto;"
        if size_or_caption.isdigit():
            width_style = f"max-width:{size_or_caption}px;width:100%;height:auto;"

        alt_text = html.escape(Path(target).stem or "image", quote=True)
        return f'<img src="{html.escape(web_url, quote=True)}" alt="{alt_text}" style="{width_style}" />'

    return re.sub(r"!\[\[([^\]]+)\]\]", replace_embed, md_content)


def replace_obsidian_embeds_local(
    md_content: str,
    note_path: str,
    page_name: str,
    output_assets_dir: Path,
) -> str:
    """Replace Obsidian embeds with local file references for dry-run export."""
    export_cache: dict[str, str] = {}

    def replace_embed(match: re.Match[str]) -> str:
        inner = match.group(1).strip()
        target, _, size_or_caption = inner.partition("|")
        target = target.strip()
        size_or_caption = size_or_caption.strip()

        local_asset = render_and_store_obsidian_asset_locally(note_path, page_name, target, output_assets_dir, export_cache)
        if not local_asset:
            return match.group(0)

        width_style = "max-width:100%;height:auto;"
        if size_or_caption.isdigit():
            width_style = f"max-width:{size_or_caption}px;width:100%;height:auto;"

        alt_text = html.escape(Path(target).stem or "image", quote=True)
        return f'<img src="{html.escape(local_asset, quote=True)}" alt="{alt_text}" style="{width_style}" />'

    return re.sub(r"!\[\[([^\]]+)\]\]", replace_embed, md_content)


def resolve_page_body(md_content: str, site_id: str, token: str, note_path: str, page_name: str) -> str:
    """Prepare markdown body for SharePoint publishing, including media embeds."""
    md_content = replace_obsidian_embeds(md_content, site_id, token, note_path, page_name)
    return md_content


def resolve_local_page_body(md_content: str, note_path: str, page_name: str, output_assets_dir: Path) -> str:
    """Prepare markdown body for local dry-run export, including media embeds."""
    md_content = replace_obsidian_embeds_local(md_content, note_path, page_name, output_assets_dir)
    return md_content


def export_dry_run_page(
    note_path: str,
    title: str,
    page_name: str,
    html_body: str,
) -> Path:
    """Write a local HTML preview of the page into a dry-run folder."""
    note_file = Path(note_path).resolve()
    output_root = note_file.parent / "dry-run"
    output_assets_dir = output_root / "assets"
    output_root.mkdir(parents=True, exist_ok=True)
    output_assets_dir.mkdir(parents=True, exist_ok=True)

    page_file_name = f"{Path(page_name).stem}.html"
    output_file = output_root / page_file_name
    full_html = (
        "<!doctype html>\n"
        "<html lang=\"sk\">\n"
        "<head>\n"
        f"<meta charset=\"utf-8\">\n<title>{html.escape(title)}</title>\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
        "<style>body{font-family:system-ui,sans-serif;max-width:980px;margin:40px auto;padding:0 20px;line-height:1.6}img{display:block;margin:12px 0;max-width:100%}pre,code{white-space:pre-wrap}</style>\n"
        "</head>\n"
        "<body>\n"
        f"<h1>{html.escape(title)}</h1>\n"
        f"{html_body}\n"
        "</body>\n"
        "</html>\n"
    )
    output_file.write_text(full_html, encoding="utf-8")
    return output_file


def clean_obsidian_footer(md_content: str) -> str:
    """Remove Obsidian footer artifacts that should not reach SharePoint."""
    nav_line_pattern = re.compile(r"^(Predchádzajúca|Nasledujúca|Ďalšia)\s*:\s*\[\[[^\]]+\]\]\s*$")
    lines = md_content.splitlines()

    # Remove trailing footer blocks from the end of the note.
    while lines and not lines[-1].strip():
        lines.pop()

    removed_footer_content = False
    while lines:
        stripped = lines[-1].strip()

        if stripped == "undefined":
            lines.pop()
            removed_footer_content = True
            continue

        if nav_line_pattern.match(stripped):
            lines.pop()
            removed_footer_content = True
            continue

        if stripped == "---" and removed_footer_content:
            lines.pop()
            removed_footer_content = True
            continue

        break

    cleaned = "\n".join(lines).strip()
    cleaned = re.sub(r"(?:\n\s*){3,}$", "\n\n", cleaned)
    return cleaned


def parse_scopes(value: str | None) -> list[str] | None:
    """Parse a comma-separated scopes string into a clean list."""
    if not value:
        return None

    scopes = [scope.strip() for scope in value.split(",")]
    scopes = [scope for scope in scopes if scope]
    return scopes or None


def detect_subject(file_path: str) -> tuple:
    """
    Walk up the path and find the first folder that matches SUBJECT_MAP.
    Returns (vault_folder_name, mapping_dict) or (None, None).
    """
    path = Path(file_path).resolve()
    for part in path.parts:
        if part in SUBJECT_MAP:
            return part, SUBJECT_MAP[part]
    return None, None


def detect_year_from_vault(file_path: str) -> str | None:
    """
    Walk up to the vault root (folder containing .obsidian)
    and parse year from its name, e.g. "SPSITknm 25-26" -> "2025-2026".
    """
    path = Path(file_path).resolve()
    for parent in path.parents:
        if (parent / ".obsidian").is_dir():
            vault_name = parent.name
            match = re.search(r"(\d{2})-(\d{2})", vault_name)
            if match:
                start, end = match.groups()
                return f"20{start}-20{end}"
            else:
                print(f"⚠️  Vault folder '{vault_name}' does not contain a year pattern like '25-26'.")
                return None
    print("⚠️  Could not find vault root (no .obsidian folder in any parent).")
    return None


def get_site_id(site_url: str, token: str) -> str:
    """Get SharePoint site ID from URL via Microsoft Graph."""
    parts = site_url.split("/")
    hostname = parts[2]
    site_path = "/".join(parts[3:])
    url = f"{GRAPH_BASE}/sites/{hostname}:/{site_path}"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    if resp.status_code == 404:
        raise RuntimeError(f"SharePoint site not found: {site_url}")
    if resp.status_code in (401, 403):
        raise RuntimeError(
            "Microsoft Graph rejected the site lookup with an authorization error. "
            "Make sure this app has SharePoint/Graph permissions for Sites.ReadWrite.All "
            "and that admin consent has been granted for the tenant."
        )
    resp.raise_for_status()
    return resp.json()["id"]


def create_page(
    site_id: str,
    title: str,
    html_content: str,
    token: str,
    page_name: str | None = None,
    description: str | None = None,
):
    """Create a SharePoint modern page via Graph API."""
    if not page_name:
        page_name = default_page_name(title)

    # Clean up artifacts
    html_content = re.sub(r"\[\[([^\]]+)\]\]", r"\1", html_content)   # Obsidian links
    html_content = html_content.replace("undefined", "")

    payload = {
        "@odata.type": "#microsoft.graph.sitePage",
        "name": page_name,
        "title": title,
        "pageLayout": "article",
        "showRecommendedPages": False,
    }

    if description is not None:
        payload["description"] = description

    payload["canvasLayout"] = {
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
    }

    url = f"{GRAPH_BASE}/sites/{site_id}/pages"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.post(url, headers=headers, json=payload)

    if resp.status_code == 409:
        print(f"⚠️  Page '{page_name}' already exists. Updating instead...")
        existing_page = get_page_by_name(site_id, page_name, token)
        if existing_page:
            return update_page(site_id, existing_page["id"], title, html_content, token, description=description)
        raise RuntimeError(f"A page named '{page_name}' already exists but could not be found for update.")
    if resp.status_code in (401, 403):
        raise RuntimeError(
            "Microsoft Graph rejected the page creation request with an authorization error. "
            "Check Sites.ReadWrite.All permissions and tenant admin consent for this app."
        )
    if resp.status_code >= 400:
        try:
            err = resp.json()
            print(f"Graph API error: {json.dumps(err, indent=2)}")
        except Exception:
            print(f"HTTP {resp.status_code}: {resp.text}")
    resp.raise_for_status()
    return resp.json()


def get_page_by_name(site_id: str, page_name: str, token: str):
    """Look up a SharePoint page by file name."""
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json;odata.metadata=none"}

    # Try 1: Filter query on the pages endpoint
    url = f"{GRAPH_BASE}/sites/{site_id}/pages/microsoft.graph.sitePage"
    escaped_page_name = page_name.replace("'", "''")
    params = {"$filter": f"name eq '{escaped_page_name}'", "$top": 1}
    resp = requests.get(url, headers=headers, params=params)

    if resp.status_code in (401, 403):
        raise RuntimeError(
            "Microsoft Graph rejected the page lookup with an authorization error. "
            "Check Sites.ReadWrite.All permissions and tenant admin consent for this app."
        )
    resp.raise_for_status()

    results = resp.json().get("value", [])
    if results:
        return results[0]

    # Try 2: Fallback — list all pages and match manually (filter sometimes misses existing pages)
    all_pages = list_site_pages(site_id, token)
    page_name_lower = page_name.lower()
    for page in all_pages:
        if (page.get("name") or "").lower() == page_name_lower:
            return page
        # Also match by title-derived name in case the stored name differs
        title_name = default_page_name(page.get("title") or "").lower()
        if title_name == page_name_lower:
            return page

    return None


def list_site_pages(site_id: str, token: str) -> list[dict]:
    """List all SharePoint pages in a site."""
    url = f"{GRAPH_BASE}/sites/{site_id}/pages/microsoft.graph.sitePage"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json;odata.metadata=none"}
    params = {"$select": "id,name,title,description,webUrl", "$top": 100}
    pages: list[dict] = []

    while url:
        resp = requests.get(url, headers=headers, params=params)
        if resp.status_code in (401, 403):
            raise RuntimeError(
                "Microsoft Graph rejected the page listing request with an authorization error. "
                "Check Sites.ReadWrite.All permissions and tenant admin consent for this app."
            )
        resp.raise_for_status()

        data = resp.json()
        pages.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None

    return pages


def build_index_page_html(pages: list[dict]) -> str:
    """Build the HTML for the shared navigation index page."""
    filtered_pages = [page for page in pages if (page.get("title") or "").strip().lower() != INDEX_PAGE_TITLE.lower()]
    filtered_pages.sort(key=page_sort_key)

    undated = [page for page in filtered_pages if not extract_page_date(page)]
    dated = [page for page in filtered_pages if extract_page_date(page)]

    sections: list[str] = []

    if undated:
        sections.append("<h2>Pages without date</h2>")
        sections.append("<ul>")
        for page in undated:
            title = html.escape(page.get("title") or page.get("name") or "Untitled")
            url = html.escape(page.get("webUrl") or "#", quote=True)
            sections.append(f'<li><a href="{url}">{title}</a></li>')
        sections.append("</ul>")

    if dated:
        sections.append("<h2>Dated pages</h2>")
        sections.append("<ul>")
        for page in dated:
            page_date = extract_page_date(page)
            if isinstance(page_date, dt.datetime):
                date_label = page_date.strftime("%Y-%m-%d %H:%M")
            elif isinstance(page_date, dt.date):
                date_label = page_date.isoformat()
            else:
                date_label = ""
            title = html.escape(page.get("title") or page.get("name") or "Untitled")
            url = html.escape(page.get("webUrl") or "#", quote=True)
            prefix = f"<strong>{date_label}</strong> — " if date_label else ""
            sections.append(f'<li>{prefix}<a href="{url}">{title}</a></li>')
        sections.append("</ul>")

    body = "\n".join(sections) if sections else "<p>No pages yet.</p>"
    return f"<div><p>Auto-generated index of published notes.</p>{body}</div>"


def refresh_navigation_index(site_id: str, token: str, current_page: dict | None = None):
    """Create or update the shared navigation index page."""
    pages = list_site_pages(site_id, token)
    if current_page and not any(page.get("id") == current_page.get("id") for page in pages):
        pages.append(current_page)

    index_title = INDEX_PAGE_TITLE
    index_page_name = default_page_name(index_title)
    index_html = build_index_page_html(pages)
    index_page = get_page_by_name(site_id, index_page_name, token)

    if index_page:
        return update_page(site_id, index_page["id"], index_title, index_html, token, description=INDEX_PAGE_DESCRIPTION)
    return create_page(site_id, index_title, index_html, token, index_page_name, description=INDEX_PAGE_DESCRIPTION)


def update_page(
    site_id: str,
    page_id: str,
    title: str,
    html_content: str,
    token: str,
    description: str | None = None,
):
    """Update an existing SharePoint modern page via Graph API."""
    payload = {
        "@odata.type": "#microsoft.graph.sitePage",
        "title": title,
        "showRecommendedPages": False,
        "description": description,
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

    if description is not None:
        payload["description"] = description

    url = f"{GRAPH_BASE}/sites/{site_id}/pages/{page_id}/microsoft.graph.sitePage"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.patch(url, headers=headers, json=payload)

    if resp.status_code in (401, 403):
        raise RuntimeError(
            "Microsoft Graph rejected the page update request with an authorization error. "
            "Check Sites.ReadWrite.All permissions and tenant admin consent for this app."
        )
    if resp.status_code >= 400:
        try:
            err = resp.json()
            print(f"Graph API error: {json.dumps(err, indent=2)}")
        except Exception:
            print(f"HTTP {resp.status_code}: {resp.text}")
    resp.raise_for_status()
    return resp.json()


def publish_page(site_id: str, page_id: str, token: str):
    """Publish a SharePoint modern page via Graph API."""
    url = f"{GRAPH_BASE}/sites/{site_id}/pages/{page_id}/microsoft.graph.sitePage/publish"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.post(url, headers=headers)

    if resp.status_code in (401, 403):
        raise RuntimeError(
            "Microsoft Graph rejected the page publish request with an authorization error. "
            "Check Sites.ReadWrite.All permissions and tenant admin consent for this app."
        )
    if resp.status_code >= 400:
        try:
            err = resp.json()
            print(f"Graph API error: {json.dumps(err, indent=2)}")
        except Exception:
            print(f"HTTP {resp.status_code}: {resp.text}")
    resp.raise_for_status()


# ============================================================
# MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="Push Obsidian note to SharePoint subject site")
    parser.add_argument("--file", "-f", required=True, help="Absolute path to the markdown file")
    parser.add_argument("--year", help="School year override, e.g. 2025-2026 (auto-detected by default)")
    parser.add_argument("--title", help="Page title override (default: filename without extension)")
    parser.add_argument("--page-name", help="Target .aspx file name (default: slugified title)")
    parser.add_argument("--tenant-id", default=DEFAULT_TENANT_ID, help="Azure tenant ID override")
    parser.add_argument("--client-id", default=DEFAULT_CLIENT_ID, help="Azure app client ID override")
    parser.add_argument("--client-secret", default=DEFAULT_CLIENT_SECRET, help="Azure app client secret override")
    parser.add_argument(
        "--scopes",
        help="Comma-separated Graph scopes override, e.g. User.Read,Files.ReadWrite.All,Sites.ReadWrite.All",
    )
    parser.add_argument("--dry-run", action="store_true", help="Create or edit a draft page with a dry-run prefix without publishing")
    args = parser.parse_args()

    file_path = normalize_input_file_path(args.file).resolve()
    if not file_path.exists():
        print(f"❌ File not found: {file_path}")
        sys.exit(1)

    # Detect subject from folder path
    vault_folder, subject_info = detect_subject(str(file_path))
    if not subject_info:
        print("❌ Could not detect subject from file path.")
        print(f"   File: {file_path}")
        print("   Make sure the file is inside one of these subject folders:")
        for k in SUBJECT_MAP:
            print(f"      • {k}")
        sys.exit(1)

    sp_name = subject_info["sp_name"]
    sp_alias = subject_info["alias"]
    print(f"📁 Vault folder:  {vault_folder}")
    print(f"📚 Subject:       {sp_name} (alias: {sp_alias})")

    # Detect or use provided year
    year = args.year or detect_year_from_vault(str(file_path))
    if not year:
        print("❌ Could not auto-detect school year from vault name.")
        print("   Pass it manually: --year 2025-2026")
        sys.exit(1)
    print(f"📅 School year:   {year}")

    # Build target SharePoint site URL
    site_url = f"https://{TENANT_NAME}.sharepoint.com/sites/{SCHOOL_ALIAS}-{year}-{sp_alias}"
    print(f"🌐 Target site:   {site_url}")

    # Read and parse markdown
    with open(file_path, "r", encoding="utf-8") as f:
        raw = f.read()

    frontmatter, body = parse_frontmatter(raw)

    # Determine title: filename stem by default, --title overrides
    title = args.title if args.title else file_path.stem
    print(f"📝 Page title:    {title}")

    requested_scopes = parse_scopes(args.scopes)
    if requested_scopes:
        print(f"🔑 Scopes:        {', '.join(requested_scopes)}")
    print(f"🆔 Client ID:     {args.client_id}")
    print(f"🏛️  Tenant ID:    {args.tenant_id}")

    page_name = args.page_name if args.page_name else default_page_name(title)
    body = clean_obsidian_footer(body)

    note_date = extract_note_date(frontmatter, title)
    note_description = note_date.isoformat() if note_date else ""

    if args.dry_run:
        dry_run_root = Path(file_path).resolve().parent / "dry-run"
        dry_run_assets_dir = dry_run_root / "assets"
        body = resolve_local_page_body(body, str(file_path), page_name, dry_run_assets_dir)
        html_body = md_to_html(body)

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

        output_file = export_dry_run_page(str(file_path), title, page_name, html_body)
        print(f"🏁 Dry-run mode — created local preview at {output_file}")
        sys.exit(0)

    # Auth and publish
    print("\n🔐 Authenticating with Microsoft Graph...")
    token = get_access_token(
        client_id=args.client_id,
        client_secret=args.client_secret,
        tenant_id=args.tenant_id,
        scopes=requested_scopes,
    )
    if not token:
        raise RuntimeError("Authentication failed: no access token was returned.")

    print("📍 Resolving SharePoint site...")
    site_id = get_site_id(site_url, token)

    body = resolve_page_body(body, site_id, token, str(file_path), page_name)
    html_body = md_to_html(body)

    # Prepend tags if present in frontmatter
    tags = frontmatter.get("tags", [])
    if isinstance(tags, str):
        tags = [tags]
    if tags:
        tag_html = (
            '<p style="color:#555;font-size:0.85em;margin-bottom:12px;">'
            + " " .join(
                f'<span style="background:#f0f0f0;padding:2px 8px;border-radius:10px;margin-right:4px;">{t}</span>'
                for t in tags
            )
            + "</p>"
        )
        html_body = tag_html + html_body

    existing_page = get_page_by_name(site_id, page_name, token)

    if existing_page:
        print(f"✏️  Editing existing page: {page_name}")
        page = update_page(site_id, existing_page["id"], title, html_body, token, description=note_description)
        action = "edited"
    else:
        print(f"🆕 Creating new page: {page_name}")
        page = create_page(site_id, title, html_body, token, page_name, description=note_description)
        action = "created"

    page_id = page.get("id")
    if not page_id:
        raise RuntimeError("Graph did not return a page id, so the page cannot be published.")

    if args.dry_run:
        print(f"🏁 Dry-run mode — draft page created or updated as {page_name}.")
        print(f"   {page.get('webUrl', 'N/A')}")
        sys.exit(0)

    print(f"📤 Publishing page after {action}...")
    publish_page(site_id, page_id, token)

    print("🧭 Refreshing Notebook index...")
    refresh_navigation_index(site_id, token, current_page=page)

    web_url = page.get("webUrl", "N/A")
    print(f"\n✅ Published!")
    print(f"   {web_url}")


if __name__ == "__main__":
    main()
