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
    System: latex + dvipng on PATH (Arch: texlive-basic texlive-latex
            texlive-latexrecommended texlive-latexextra texlive-bin)

Auth overrides:
    --client-id, --client-secret, --tenant-id, --scopes
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
import subprocess
import tempfile
import textwrap
import time
from pathlib import Path
from urllib.parse import quote, unquote

try:
    import requests
    import markdown
    import yaml
    from PIL import Image, ImageDraw, ImageFont
except ImportError as e:
    print(f"Missing dependency: {e.name}")
    print("Install: pip install requests msal markdown pyyaml pillow")
    sys.exit(1)

try:
    from playwright.sync_api import sync_playwright
    _HAS_PLAYWRIGHT = True
except ImportError:
    _HAS_PLAYWRIGHT = False

from ms_graph import get_access_token


# ============================================================
# CONFIGURATION
# ============================================================
TENANT_NAME = "0z6sw"
SCHOOL_ALIAS = "spsitknm"
DEFAULT_TENANT_ID = os.getenv("TENANT_ID", "").strip()
DEFAULT_CLIENT_ID = os.getenv("O2S_ID", "").strip()
DEFAULT_CLIENT_SECRET = os.getenv("O2S_SECRET", "").strip()

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
    "10 - SJL":       {"sp_name": "SJL",       "alias": "sjl"},
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
# LaTeX TOOLCHAIN DETECTION
# ============================================================

def _detect_latex_toolchain() -> tuple:
    """Return (available, reason). Checks latex and dvipng on PATH."""
    for tool in ("latex", "dvipng"):
        try:
            subprocess.run(
                [tool, "--version"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=True,
            )
        except (FileNotFoundError, subprocess.CalledProcessError):
            return False, f"{tool} not found on PATH"
    return True, "ok"


_HAS_LATEX, _LATEX_REASON = _detect_latex_toolchain()


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
    """Convert markdown to HTML with fenced code, table, and newline support."""
    md = markdown.Markdown(extensions=["fenced_code", "tables", "nl2br", "md_in_html"])
    return md.convert(md_content)


def default_page_name(title: str) -> str:
    """Build the default SharePoint page file name from a title."""
    page_name = re.sub(r"[^\w\-]", "-", title).lower()
    if not page_name.endswith(".aspx"):
        page_name += ".aspx"
    return page_name


def parse_date_string(value: object):
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

    dt_patterns = [
        r"(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})[ T](?P<hour>\d{2})[-:](?P<minute>\d{2})[-:]?(?P<second>\d{2})?",
        r"(?P<year>\d{4})/(?P<month>\d{2})/(?P<day>\d{2})[ T](?P<hour>\d{2})[-:](?P<minute>\d{2})[-:]?(?P<second>\d{2})?",
    ]
    for pattern in dt_patterns:
        match = re.search(pattern, text)
        if match:
            try:
                return dt.datetime(
                    int(match.group("year")), int(match.group("month")), int(match.group("day")),
                    int(match.group("hour")), int(match.group("minute")), int(match.group("second") or 0)
                )
            except ValueError:
                continue

    for pattern in (r"(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})", r"(?P<year>\d{4})/(?P<month>\d{2})/(?P<day>\d{2})"):
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


def extract_note_date(frontmatter: dict, title: str):
    title_date = parse_date_string(title)
    if title_date:
        return title_date
    for key in ("date", "published", "publish_date", "publish-date"):
        d = parse_date_string(frontmatter.get(key))
        if d:
            return d
    return None


def extract_page_date(page: dict):
    d = parse_date_string(page.get("description"))
    if d:
        return d
    d = parse_date_string(page.get("title") or page.get("name"))
    if d:
        return d
    return None


def page_sort_key(page: dict) -> tuple:
    page_date = extract_page_date(page)
    title = (page.get("title") or page.get("name") or "").lower()
    if not page_date:
        return (0, 0, title)
    if isinstance(page_date, dt.datetime):
        return (1, -page_date.timestamp(), title)
    return (1, -page_date.toordinal(), title)


def safe_slug(value: str) -> str:
    slug = re.sub(r"[^a-zA-Z0-9._-]+", "-", value).strip("-")
    return slug.lower() or "item"


def find_vault_root(file_path: str):
    path = Path(file_path).resolve()
    for parent in path.parents:
        if (parent / ".obsidian").is_dir():
            return parent
    return None


def normalize_input_file_path(raw_path: str) -> Path:
    candidate = raw_path.strip().strip('"').strip("'")
    candidate = candidate.replace("\\/", "/")
    candidate = re.sub(r"\\([\\/ ._\-()\[\]{}])", r"\1", candidate)
    candidate = candidate.replace("\\", "")
    home_index = candidate.rfind("/home/")
    if home_index != -1:
        candidate = candidate[home_index:]
    return Path(candidate)


def load_font(size: int):
    for candidate in FONT_CANDIDATES:
        if os.path.exists(candidate):
            return ImageFont.truetype(candidate, size=size)
    return ImageFont.load_default()


# ------------------------------------------------------------
# Excalidraw / Obsidian asset resolution (unchanged)
# ------------------------------------------------------------

def _lz_get_base_value(char: str) -> int:
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    index = alphabet.find(char)
    if index < 0:
        raise ValueError(f"Invalid LZ-string character: {char!r}")
    return index


def lzstring_decompress_from_encoded_uri_component(compressed: str) -> str:
    if compressed is None or compressed == "":
        return ""
    compressed = compressed.replace(" ", "+")
    length = len(compressed)

    def get_next_value(index: int) -> int:
        return _lz_get_base_value(compressed[index])

    dictionary = ["", "", ""]
    enlarge_in = 4
    dict_size = 4
    num_bits = 3
    entry = ""
    result = []

    data_val = get_next_value(0)
    data_position = 32
    data_index = 1

    def read_bits(n: int) -> int:
        nonlocal data_val, data_position, data_index
        bits = 0
        power = 1
        max_power = 2 ** n
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
            enlarge_in = 2 ** num_bits
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
            enlarge_in = 2 ** num_bits
            num_bits += 1


def _extract_excalidraw_json(excalidraw_path: Path) -> dict:
    text = excalidraw_path.read_text(encoding="utf-8")
    match = re.search(r"```compressed-json\s*(.*?)\s*```", text, re.S)
    if not match:
        raise RuntimeError(f"Could not find compressed drawing data in {excalidraw_path}")
    compressed = "".join(match.group(1).split())
    json_text = lzstring_decompress_from_encoded_uri_component(compressed)
    return json.loads(json_text)


def render_excalidraw_to_png(excalidraw_path: Path, output_path: Path) -> Path:
    try:
        import excalidraw_to_png as exporter_module
        ExcalidrawExporter = exporter_module.ExcalidrawExporter
        find_vault_root_exp = getattr(exporter_module, "find_vault_root", None)
        has_playwright = getattr(exporter_module, "sync_playwright", None) is not None
    except Exception as exc:
        raise RuntimeError(f"Excalidraw exporter not available: {exc}")

    if find_vault_root_exp:
        vault_root = find_vault_root_exp(excalidraw_path)
    else:
        vault_root = find_vault_root(str(excalidraw_path)) or excalidraw_path.parent

    exporter = ExcalidrawExporter(
        vault_path=str(vault_root),
        output_dir=str(output_path.parent),
    )

    def try_export(path: Path):
        try:
            res = exporter.export_with_cli(path, output_path)
            if res:
                return res
        except Exception:
            pass
        if has_playwright:
            try:
                res = exporter.export_with_playwright(path, output_path)
                if res:
                    return res
            except Exception:
                pass
        return None

    result = try_export(excalidraw_path)
    if result:
        return result

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


def find_obsidian_asset_path(note_path: str, target: str):
    note_file = Path(note_path).resolve()
    note_dir = note_file.parent
    vault_root = find_vault_root(note_path)
    target = target.strip()
    if not target:
        return None
    direct = Path(target)
    candidates = []
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


def build_remote_asset_name(page_name: str, source_key: str, local_path: Path, source_name=None) -> str:
    page_slug = safe_slug(Path(page_name).stem if page_name else "page")
    source_slug = safe_slug(source_name or local_path.stem)
    digest = hashlib.sha1(source_key.encode("utf-8")).hexdigest()[:10]
    suffix = local_path.suffix.lower() or ".png"
    return f"{page_slug}-{source_slug}-{digest}{suffix}"


def build_local_asset_name(page_name: str, source_key: str, local_path: Path, source_name=None) -> str:
    return build_remote_asset_name(page_name, source_key, local_path, source_name)


def upload_file_to_site(site_id: str, local_path: Path, remote_name: str, token: str) -> dict:
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
    output_assets_dir.mkdir(parents=True, exist_ok=True)
    destination = output_assets_dir / asset_name
    shutil.copy2(local_path, destination)
    return destination


def is_excalidraw_source(path: Path) -> bool:
    return path.name.endswith(".excalidraw.md")


def render_and_upload_obsidian_asset(
    site_id: str,
    token: str,
    note_path: str,
    page_name: str,
    target: str,
    upload_cache: dict,
):
    source_path = find_obsidian_asset_path(note_path, target)
    if not source_path:
        return None
    source_key = str(source_path.resolve())
    if source_key in upload_cache:
        return upload_cache[source_key]
    upload_path = source_path
    temp_path = None
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
    export_cache: dict,
):
    source_path = find_obsidian_asset_path(note_path, target)
    if not source_path:
        return None
    source_key = str(source_path.resolve())
    if source_key in export_cache:
        return export_cache[source_key]
    asset_path = source_path
    temp_path = None
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
    upload_cache: dict = {}

    def replace_embed(match: re.Match) -> str:
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
    export_cache: dict = {}

    def replace_embed(match: re.Match) -> str:
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


# ============================================================
# LaTeX rendering via latex + dvipng -> PNG
# ============================================================

_LATEX_DOCUMENT_TEMPLATE = r"""\documentclass[border=2pt]{standalone}
\usepackage{amsmath}
\usepackage{amssymb}
\usepackage{amsfonts}
\usepackage{varwidth}
\begin{document}
%(body)s
\end{document}
"""


def _run_latex_toolchain(
    latex_source: str,
    workdir: Path,
    timeout: int = 30,
    dpi: int = 300,
) -> Path:
    """
    Compile a LaTeX snippet with latex + dvipng.
    Returns the path to the produced .png file.
    Raises RuntimeError on any failure.
    """
    tex_file = workdir / "formula.tex"
    dvi_file = workdir / "formula.dvi"
    png_file = workdir / "formula.png"

    tex_file.write_text(latex_source, encoding="utf-8")

    try:
        latex_proc = subprocess.run(
            [
                "latex",
                "-interaction=nonstopmode",
                "-halt-on-error",
                "-output-directory", str(workdir),
                str(tex_file),
            ],
            cwd=str(workdir),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("latex timed out")

    if latex_proc.returncode != 0 or not dvi_file.exists():
        log_tail = latex_proc.stdout.decode("utf-8", errors="replace")[-1500:]
        raise RuntimeError(f"latex failed:\n{log_tail}")

    try:
        dvipng_proc = subprocess.run(
            [
                "dvipng",
                "-D", str(dpi),
                "-T", "tight",
                "-bg", "Transparent",
                "-o", str(png_file),
                str(dvi_file),
            ],
            cwd=str(workdir),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("dvipng timed out")

    if dvipng_proc.returncode != 0 or not png_file.exists():
        log_tail = dvipng_proc.stdout.decode("utf-8", errors="replace")[-1500:]
        raise RuntimeError(f"dvipng failed:\n{log_tail}")

    return png_file


def _render_latex_to_png_file(
    latex_expr: str,
    display_mode: bool,
    output_path: Path,
) -> bool:
    """
    Render a LaTeX expression to a PNG file at output_path.
    Returns True on success, False on failure.
    """
    if not _HAS_LATEX:
        return False

    expr = latex_expr.strip()
    if not expr:
        return False

    expr = re.sub(r"^\$\$|\$\$$", "", expr).strip()

    if display_mode:
        body = f"$\\displaystyle {expr}$"
        dpi = 300
    else:
        body = f"${expr}$"
        dpi = 240  # smaller raster so inline math matches body text height

    source = _LATEX_DOCUMENT_TEMPLATE % {"body": body}

    with tempfile.TemporaryDirectory(prefix="o2s_latex_") as tmp:
        workdir = Path(tmp)
        try:
            produced_png = _run_latex_toolchain(source, workdir, dpi=dpi)
        except RuntimeError as exc:
            print(f"      \u26a0\ufe0f  LaTeX render failed for: {expr[:60]}... ({exc})")
            return False

        shutil.copy2(produced_png, output_path)

    return True


def replace_latex_expressions(
    md_content: str,
    page_name: str,
    store_func,
    cache: dict,
) -> str:
    """
    Replace $...$ and $$...$$ with rendered PNG images.
    store_func(tmp_path, asset_name) -> public URL string
    """
    if not _HAS_LATEX:
        print(f"\u26a0\ufe0f  LaTeX toolchain unavailable ({_LATEX_REASON}); "
              f"leaving raw $...$ math in place.")
        return md_content

    def render(latex_expr: str, display_mode: bool):
        expr_clean = latex_expr.strip()
        if not expr_clean:
            return None

        expr_key = f"latex:{'D' if display_mode else 'I'}:{expr_clean}"
        if expr_key in cache:
            return cache[expr_key]

        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            tmp_path = Path(tmp.name)

        try:
            ok = _render_latex_to_png_file(expr_clean, display_mode, tmp_path)
            if not ok:
                return None
            asset_name = build_remote_asset_name(page_name, expr_key, tmp_path, "latex")
            url = store_func(tmp_path, asset_name)
            cache[expr_key] = url
            return url
        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except Exception:
                    pass

    # Protect code blocks
    code_pattern = r"(```.*?```|~~~.*?~~~|`[^`]*`)"
    parts = re.split(code_pattern, md_content, flags=re.DOTALL)

    processed = []
    for i, part in enumerate(parts):
        if i % 2 == 1:
            processed.append(part)
            continue

        # Display math $$...$$
        def replace_display(m: re.Match) -> str:
            url = render(m.group(1), True)
            if url:
                # Removed \n\n so it doesn't break out of lists
                return f'<img src="{html.escape(url, quote=True)}" alt="math" style="display:block;margin:16px auto;max-width:100%;" />'
            return m.group(0)

        part = re.sub(r"\$\$(.+?)\$\$", replace_display, part, flags=re.DOTALL)

        # Inline math $...$
        def replace_inline(m: re.Match) -> str:
            expr = m.group(1).strip()
            if not expr or re.fullmatch(r"[\d\s,.]+", expr):
                return m.group(0)
            url = render(expr, False)
            if url:
                return (
                    f'<img src="{html.escape(url, quote=True)}" alt="math" '
                    'style="display:inline-block;vertical-align:-0.25em;'
                    'height:1em;width:auto;" />'
                )
            return m.group(0)

        part = re.sub(r"(?<!\$)\$(?!\$)([^$\n]+?)\$(?!\$)", replace_inline, part)
        processed.append(part)

    return "".join(processed)


# ============================================================
# Desmos handling
# ============================================================

def _screenshot_desmos(url: str, output_path: Path, timeout_ms: int = 15000) -> bool:
    if not _HAS_PLAYWRIGHT:
        return False
    try:
        with sync_playwright() as p:
            browser = p.firefox.launch(headless=True)
            page = browser.new_page(viewport={"width": 800, "height": 600})
            page.goto(url, wait_until="networkidle", timeout=timeout_ms)
            page.wait_for_selector(".dcg-graph-paper", timeout=timeout_ms)
            time.sleep(1.5)
            page.screenshot(path=str(output_path), full_page=False)
            browser.close()
            return True
    except Exception as exc:
        print(f"      \u26a0\ufe0f  Desmos screenshot failed: {exc}")
        return False


def replace_desmos_embeds(
    md_content: str,
    page_name=None,
    store_func=None,
    cache=None,
) -> str:
    DESMOS_URL_RE = r"https://(?:www\.)?desmos\.com/calculator/[a-zA-Z0-9]+"
    DESMOS_HREF_RE = r"https://(?:www\.)?desmos\.com/calculator/[^)\"]+"

    def desmos_block_repl(m: re.Match) -> str:
        content = html.escape(m.group(1).strip())
        return (
            '\n\n<div style="border:1px solid #ddd;border-radius:6px;margin:12px 0;overflow:hidden;">'
            '<div style="background:#2f72a8;color:#fff;padding:6px 12px;font-family:system-ui,sans-serif;'
            'font-size:0.9em;font-weight:600;">Desmos Graph</div>'
            f'<pre style="margin:0;padding:12px;background:#f8f9fa;overflow:auto;"><code>{content}</code></pre>'
            '<div style="padding:8px 12px;background:#f8f9fa;border-top:1px solid #eee;">'
            '<a href="https://www.desmos.com/calculator" target="_blank" rel="noopener" '
            'style="color:#2f72a8;text-decoration:none;font-weight:500;">Open Desmos Calculator \u2192</a>'
            "</div></div>\n\n"
        )

    md_content = re.sub(
        r"```desmos(?:-graph)?\s*(.*?)```",
        desmos_block_repl,
        md_content,
        flags=re.DOTALL | re.IGNORECASE,
    )

    code_pattern = r"(```.*?```|~~~.*?~~~|<code>.*?</code>)"
    parts = re.split(code_pattern, md_content, flags=re.DOTALL | re.IGNORECASE)

    processed = []
    for i, part in enumerate(parts):
        if i % 2 == 1:
            processed.append(part)
            continue

        def process_desmos_url(url: str, alt_text: str = "Desmos Graph") -> str:
            img_url = None
            if store_func and cache is not None and page_name:
                cache_key = f"desmos:{url}"
                if cache_key in cache:
                    img_url = cache[cache_key]
                else:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                        tmp_path = Path(tmp.name)
                    try:
                        ok = _screenshot_desmos(url, tmp_path)
                        if ok:
                            asset_name = build_remote_asset_name(page_name, cache_key, tmp_path, "desmos")
                            img_url = store_func(tmp_path, asset_name)
                            cache[cache_key] = img_url
                    finally:
                        if tmp_path.exists():
                            try:
                                tmp_path.unlink()
                            except Exception:
                                pass

            if img_url:
                return (
                    f'\n\n<div style="margin:12px 0;text-align:center;">'
                    f'<a href="{html.escape(url, quote=True)}" target="_blank" rel="noopener">'
                    f'<img src="{html.escape(img_url, quote=True)}" alt="{html.escape(alt_text)}" '
                    'style="max-width:100%;border:1px solid #ddd;border-radius:6px;" /></a>'
                    f'<div style="margin-top:4px;"><a href="{html.escape(url, quote=True)}" target="_blank" '
                    'rel="noopener" style="font-size:0.85em;color:#2f72a8;">Open interactive graph \u2192</a></div>'
                    '</div>\n\n'
                )

            return (
                f'\n\n<a href="{html.escape(url, quote=True)}" target="_blank" rel="noopener" '
                'style="display:inline-block;padding:8px 16px;background:#2f72a8;color:#fff;'
                'text-decoration:none;border-radius:4px;font-family:system-ui,sans-serif;">'
                f"\U0001F537 {html.escape(alt_text)}</a>\n\n"
            )

        def iframe_repl(m: re.Match) -> str:
            return process_desmos_url(m.group(1), "Desmos Graph")

        part = re.sub(
            rf'<iframe[^>]+src="({DESMOS_HREF_RE})"[^>]*>.*?</iframe>',
            iframe_repl,
            part,
            flags=re.DOTALL | re.IGNORECASE,
        )

        def md_img_repl(m: re.Match) -> str:
            alt_text = m.group(1) or "Desmos Graph"
            url = m.group(2)
            return process_desmos_url(url, alt_text)

        part = re.sub(
            rf"!\[([^\]]*)\]\(({DESMOS_HREF_RE})\)",
            md_img_repl,
            part,
        )

        def md_link_repl(m: re.Match) -> str:
            text = m.group(1) or "Desmos Graph"
            url = m.group(2)
            return process_desmos_url(url, text)

        part = re.sub(
            rf"(?<!!)\[([^\]]+)\]\(({DESMOS_HREF_RE})\)",
            md_link_repl,
            part,
        )

        def bare_url_repl(m: re.Match) -> str:
            return process_desmos_url(m.group(0), "Desmos Graph")

        part = re.sub(
            rf"(?<![\\(\"=]){DESMOS_URL_RE}(?![\\w\"])",
            bare_url_repl,
            part,
        )

        processed.append(part)

    return "".join(processed)


# ============================================================
# Obsidian callouts -> styled <div> blocks
# ============================================================

CALLOUT_STYLES = {
    "note":      ("\U0001F4DD", "#2f72a8", "#eaf3fb"),
    "info":      ("\u2139\ufe0f", "#2f72a8", "#eaf3fb"),
    "tip":       ("\U0001F4A1", "#4c9a3f", "#eef8eb"),
    "hint":      ("\U0001F4A1", "#4c9a3f", "#eef8eb"),
    "important": ("\u2757", "#2f72a8", "#eaf3fb"),
    "success":   ("\u2705", "#4c9a3f", "#eef8eb"),
    "check":     ("\u2705", "#4c9a3f", "#eef8eb"),
    "done":      ("\u2705", "#4c9a3f", "#eef8eb"),
    "question":  ("\u2753", "#b07d2a", "#fbf4e8"),
    "help":      ("\u2753", "#b07d2a", "#fbf4e8"),
    "faq":       ("\u2753", "#b07d2a", "#fbf4e8"),
    "warning":   ("\u26a0\ufe0f", "#c98a1a", "#fdf5e2"),
    "caution":   ("\u26a0\ufe0f", "#c98a1a", "#fdf5e2"),
    "attention": ("\u26a0\ufe0f", "#c98a1a", "#fdf5e2"),
    "failure":   ("\u274c", "#b83a3a", "#fbecec"),
    "fail":      ("\u274c", "#b83a3a", "#fbecec"),
    "missing":   ("\u274c", "#b83a3a", "#fbecec"),
    "danger":    ("\U0001F6A8", "#b83a3a", "#fbecec"),
    "error":     ("\U0001F6A8", "#b83a3a", "#fbecec"),
    "bug":       ("\U0001F41E", "#b83a3a", "#fbecec"),
    "example":   ("\U0001F9EA", "#6f42c1", "#f3eefa"),
    "quote":     ("\u275D", "#666666", "#f5f5f5"),
    "cite":      ("\u275D", "#666666", "#f5f5f5"),
}

_CALLOUT_RE = re.compile(
    r"^>\s*\[!(?P<kind>[A-Za-z]+)\](?P<fold>[+-]?)\s*(?P<title>.*?)\n"
    r"(?P<body>(?:^>.*$\n?)*)",
    re.MULTILINE,
)


def convert_obsidian_callouts(md_content: str) -> str:
    """
    Convert Obsidian callouts of the form:
        > [!note] Title
        > body line 1
        > body line 2
    into styled <div> blocks. The body is rendered as Markdown first so
    lists, bold, links, and inline code inside callouts work correctly.
    """
    def repl(match: re.Match) -> str:
        kind = match.group("kind").lower()
        title = match.group("title").strip()
        body_raw = match.group("body") or ""

        # Strip the leading '>' from every body line.
        body_lines = []
        for line in body_raw.splitlines():
            if line.startswith(">"):
                body_lines.append(line[1:].lstrip())
            else:
                body_lines.append(line)
        while body_lines and not body_lines[0].strip():
            body_lines.pop(0)
        while body_lines and not body_lines[-1].strip():
            body_lines.pop()
        body_md = "\n".join(body_lines)

        icon, accent, background = CALLOUT_STYLES.get(
            kind, ("\U0001F4CC", "#555555", "#f5f5f5")
        )

        # Ensure the first line after the title is treated as the start
        # of a block, so lists become <ol>/<ul> instead of plain text.
        if body_md.strip():
            body_md = "\n" + body_md
        
        # Recursively convert nested callouts first.
        body_md = convert_obsidian_callouts(body_md)

        # NEW: Fix lists/code blocks inside the callout
        body_md = fix_markdown_formatting(body_md)

        # Render the body as HTML *now*, so lists/bold/code work.
        body_html = md_to_html(body_md) if body_md.strip() else ""

        title_html = (
            f'<div style="font-weight:600;margin-bottom:6px;color:{accent};">'
            f'{icon} {html.escape(title)}</div>'
            if title else
            f'<div style="font-weight:600;margin-bottom:6px;color:{accent};">'
            f'{icon} {kind.capitalize()}</div>'
        )

        return (
            f'\n\n<div style="border-left:4px solid {accent};'
            f'background:{background};padding:12px 16px;margin:16px 0;'
            f'border-radius:4px;">'
            f'{title_html}'
            f'<div>{body_html}</div>'
            f'</div>\n\n'
        )

    previous = None
    while previous != md_content:
        previous = md_content
        md_content = _CALLOUT_RE.sub(repl, md_content)
    return md_content


# ============================================================
# Body resolution
# ============================================================

def resolve_page_body(md_content: str, site_id: str, token: str, note_path: str, page_name: str) -> str:
    latex_cache: dict = {}
    desmos_cache: dict = {}

    def upload_asset(local_path: Path, name: str) -> str:
        uploaded = upload_file_to_site(site_id, local_path, name, token)
        return uploaded.get("webUrl") or uploaded.get("@microsoft.graph.downloadUrl")

    md_content = replace_latex_expressions(md_content, page_name, upload_asset, latex_cache)
    md_content = replace_desmos_embeds(md_content, page_name, upload_asset, desmos_cache)
    md_content = replace_obsidian_embeds(md_content, site_id, token, note_path, page_name)
    return md_content


def resolve_local_page_body(md_content: str, note_path: str, page_name: str, output_assets_dir: Path) -> str:
    latex_cache: dict = {}
    desmos_cache: dict = {}

    def save_asset(local_path: Path, name: str) -> str:
        copied = save_file_to_local_assets(local_path, output_assets_dir, name)
        return f"assets/{copied.name}"

    md_content = replace_latex_expressions(md_content, page_name, save_asset, latex_cache)
    md_content = replace_desmos_embeds(md_content, page_name, save_asset, desmos_cache)
    md_content = replace_obsidian_embeds_local(md_content, note_path, page_name, output_assets_dir)
    return md_content


def export_dry_run_page(
    note_path: str,
    title: str,
    page_name: str,
    html_body: str,
) -> Path:
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
        "<style>body{font-family:system-ui,sans-serif;max-width:980px;margin:40px auto;padding:0 20px;line-height:1.6}"
        "img{display:block;margin:12px 0;max-width:100%}pre,code{white-space:pre-wrap}</style>\n"
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
    nav_line_pattern = re.compile(r"^(Predchádzajúca|Nasledujúca|Ďalšia)\s*:\s*\[\[[^\]]+\]\]\s*$")
    lines = md_content.splitlines()
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


def fix_markdown_formatting(md_content: str) -> str:
    """
    Fixes Markdown formatting issues where lists and code blocks are 
    not separated by blank lines (which Python Markdown requires).
    """
    lines = md_content.split('\n')
    out_lines = []
    in_code_block = False
    
    for line in lines:
        stripped = line.strip()
        
        # Handle code blocks (```)
        if stripped.startswith('```') or stripped.startswith('~~~'):
            in_code_block = not in_code_block
            if not in_code_block: # closing tag
                out_lines.append(line)
                out_lines.append('') # ensure blank line after code block
            else: # opening tag
                if out_lines and out_lines[-1].strip() != '':
                    out_lines.append('')
                out_lines.append(line)
            continue
            
        if in_code_block:
            out_lines.append(line)
            continue
            
        # Handle lists (e.g., "- ", "* ", "1. ")
        if re.match(r'^[ \t]*(?:[-*+]|\d+\.)\s+', line):
            if out_lines and out_lines[-1].strip() != '':
                if not re.match(r'^[ \t]*(?:[-*+]|\d+\.)\s+', out_lines[-1]):
                    if not re.match(r'^[ \t]+', out_lines[-1]):
                        out_lines.append('')
                        
        out_lines.append(line)
        
    # Clean up multiple consecutive blank lines
    cleaned = []
    for line in out_lines:
        if line.strip() == '' and cleaned and cleaned[-1].strip() == '':
            continue
        cleaned.append(line)
        
    return '\n'.join(cleaned)


def parse_scopes(value):
    if not value:
        return None
    scopes = [scope.strip() for scope in value.split(",")]
    scopes = [scope for scope in scopes if scope]
    return scopes or None


def detect_subject(file_path: str):
    path = Path(file_path).resolve()
    for part in path.parts:
        if part in SUBJECT_MAP:
            return part, SUBJECT_MAP[part]
    return None, None


def detect_year_from_vault(file_path: str):
    path = Path(file_path).resolve()
    for parent in path.parents:
        if (parent / ".obsidian").is_dir():
            vault_name = parent.name
            match = re.search(r"(\d{2})-(\d{2})", vault_name)
            if match:
                start, end = match.groups()
                return f"20{start}-20{end}"
            else:
                print(f"\u26a0\ufe0f  Vault folder '{vault_name}' does not contain a year pattern like '25-26'.")
                return None
    print("\u26a0\ufe0f  Could not find vault root (no .obsidian folder in any parent).")
    return None


def get_site_id(site_url: str, token: str) -> str:
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
    page_name=None,
    description=None,
):
    if not page_name:
        page_name = default_page_name(title)

    html_content = re.sub(r"\[\[([^\]]+)\]\]", r"\1", html_content)
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
        print(f"\u26a0\ufe0f  Page '{page_name}' already exists. Updating instead...")
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
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json;odata.metadata=none"}
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

    all_pages = list_site_pages(site_id, token)
    page_name_lower = page_name.lower()
    for page in all_pages:
        if (page.get("name") or "").lower() == page_name_lower:
            return page
        title_name = default_page_name(page.get("title") or "").lower()
        if title_name == page_name_lower:
            return page
    return None


def list_site_pages(site_id: str, token: str):
    url = f"{GRAPH_BASE}/sites/{site_id}/pages/microsoft.graph.sitePage"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json;odata.metadata=none"}
    params = {"$select": "id,name,title,description,webUrl", "$top": 100}
    pages = []

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


def build_index_page_html(pages) -> str:
    filtered_pages = [page for page in pages if (page.get("title") or "").strip().lower() != INDEX_PAGE_TITLE.lower()]
    filtered_pages.sort(key=page_sort_key)

    undated = [page for page in filtered_pages if not extract_page_date(page)]
    dated = [page for page in filtered_pages if extract_page_date(page)]

    sections = []
    if undated:
        sections.append("<h2>Pages without date</h2><ul>")
        for page in undated:
            title = html.escape(page.get("title") or page.get("name") or "Untitled")
            url = html.escape(page.get("webUrl") or "#", quote=True)
            sections.append(f'<li><a href="{url}">{title}</a></li>')
        sections.append("</ul>")
    if dated:
        sections.append("<h2>Dated pages</h2><ul>")
        for page in dated:
            page_date = extract_page_date(page)
            date_label = page_date.strftime("%Y-%m-%d %H:%M") if isinstance(page_date, dt.datetime) else (
                page_date.isoformat() if isinstance(page_date, dt.date) else ""
            )
            title = html.escape(page.get("title") or page.get("name") or "Untitled")
            url = html.escape(page.get("webUrl") or "#", quote=True)
            prefix = f"<strong>{date_label}</strong> \u2014 " if date_label else ""
            sections.append(f'<li>{prefix}<a href="{url}">{title}</a></li>')
        sections.append("</ul>")

    body = "\n".join(sections) if sections else "<p>No pages yet.</p>"
    return f"<div><p>Auto-generated index of published notes.</p>{body}</div>"


def refresh_navigation_index(site_id: str, token: str, current_page=None):
    pages = list_site_pages(site_id, token)
    if current_page and not any(page.get("id") == current_page.get("id") for page in pages):
        pages.append(current_page)

    index_title = INDEX_PAGE_TITLE
    index_page_name = default_page_name(index_title)
    index_html = build_index_page_html(pages)
    index_page = get_page_by_name(site_id, index_page_name, token)

    if index_page:
        updated_page = update_page(site_id, index_page["id"], index_title, index_html, token, description=INDEX_PAGE_DESCRIPTION)
    else:
        updated_page = create_page(site_id, index_title, index_html, token, index_page_name, description=INDEX_PAGE_DESCRIPTION)

    # NEW: Publish the index page so it becomes visible to readers
    index_page_id = updated_page.get("id")
    if index_page_id:
        publish_page(site_id, index_page_id, token)
        
    return updated_page


def update_page(
    site_id: str,
    page_id: str,
    title: str,
    html_content: str,
    token: str,
    description=None,
):
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
    parser.add_argument("--dry-run", action="store_true", help="Create local preview without publishing")
    args = parser.parse_args()

    file_path = normalize_input_file_path(args.file).resolve()
    if not file_path.exists():
        print(f"\u274c File not found: {file_path}")
        sys.exit(1)

    vault_folder, subject_info = detect_subject(str(file_path))
    if not subject_info:
        print("\u274c Could not detect subject from file path.")
        print(f"   File: {file_path}")
        print("   Make sure the file is inside one of these subject folders:")
        for k in SUBJECT_MAP:
            print(f"      \u2022 {k}")
        sys.exit(1)

    sp_name = subject_info["sp_name"]
    sp_alias = subject_info["alias"]
    print(f"\U0001F4C1 Vault folder:  {vault_folder}")
    print(f"\U0001F4DA Subject:       {sp_name} (alias: {sp_alias})")

    year = args.year or detect_year_from_vault(str(file_path))
    if not year:
        print("\u274c Could not auto-detect school year from vault name.")
        print("   Pass it manually: --year 2025-2026")
        sys.exit(1)
    print(f"\U0001F4C5 School year:   {year}")

    site_url = f"https://{TENANT_NAME}.sharepoint.com/sites/{SCHOOL_ALIAS}-{year}-{sp_alias}"
    print(f"\U0001F310 Target site:   {site_url}")

    with open(file_path, "r", encoding="utf-8") as f:
        raw = f.read()

    frontmatter, body = parse_frontmatter(raw)
    title = args.title if args.title else file_path.stem
    print(f"\U0001F4DD Page title:    {title}")

    requested_scopes = parse_scopes(args.scopes)
    if requested_scopes:
        print(f"\U0001F511 Scopes:        {', '.join(requested_scopes)}")
    print(f"\U0001F194 Client ID:     {args.client_id}")
    print(f"\U0001F3DB\ufe0f  Tenant ID:    {args.tenant_id}")

    page_name = args.page_name if args.page_name else default_page_name(title)
    body = clean_obsidian_footer(body)
    
    # Insert blank lines before lists/code blocks
    body = fix_markdown_formatting(body)
    
    body = convert_obsidian_callouts(body)

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
        print(f"\U0001F3C1 Dry-run mode \u2014 created local preview at {output_file}")
        sys.exit(0)

    print("\n\U0001F510 Authenticating with Microsoft Graph...")
    token = get_access_token(
        client_id=args.client_id,
        client_secret=args.client_secret,
        tenant_id=args.tenant_id,
        scopes=requested_scopes,
    )
    if not token:
        raise RuntimeError("Authentication failed: no access token was returned.")

    print("\U0001F4CD Resolving SharePoint site...")
    site_id = get_site_id(site_url, token)

    body = resolve_page_body(body, site_id, token, str(file_path), page_name)
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

    existing_page = get_page_by_name(site_id, page_name, token)

    if existing_page:
        print(f"\u270f\ufe0f  Editing existing page: {page_name}")
        page = update_page(site_id, existing_page["id"], title, html_body, token, description=note_description)
        action = "edited"
    else:
        print(f"\U0001F195 Creating new page: {page_name}")
        page = create_page(site_id, title, html_body, token, page_name, description=note_description)
        action = "created"

    page_id = page.get("id")
    if not page_id:
        raise RuntimeError("Graph did not return a page id, so the page cannot be published.")

    print(f"\U0001F4E4 Publishing page after {action}...")
    publish_page(site_id, page_id, token)

    print("\U0001F9ED Refreshing Notebook index...")
    refresh_navigation_index(site_id, token, current_page=page)

    web_url = page.get("webUrl", "N/A")
    print(f"\n\u2705 Published!")
    print(f"   {web_url}")


if __name__ == "__main__":
    main()