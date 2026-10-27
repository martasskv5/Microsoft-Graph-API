#!/usr/bin/env python3
"""
publish_note.py
Publishes multiple Obsidian notes to a Zensical static site (Cloudflare)
and pushes a backup to SharePoint.
"""

import argparse
import datetime as dt
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import yaml

try:
    import obsidian_to_sharepoint as sp_backup
except ImportError:
    print("⚠️  Warning: obsidian_to_sharepoint.py not found. Skipping SharePoint backup.")
    sp_backup = None

# --- CONFIGURATION ---
CONTENT_REPO_DIR = Path("/home/martinko/Projects/spsknm-notes").resolve()

ZENSICAL_CONFIG = """\
[project]
site_name = "SPSITKNM Notes"
use_directory_urls = false
extra_javascript = [
    "javascripts/mathjax.js",
    "https://unpkg.com/mathjax@3/es5/tex-mml-chtml.js"
]
markdown_extensions = [
    "tables",
    "admonition",
    "pymdownx.superfences",
    { "pymdownx.arithmatex" = { generic = true } }
]

[project.theme]
name = "material"
features = [
    "navigation.tabs",
    "navigation.sections",
    "content.code.copy"
]
"""

MATHJAX_JS = r"""window.MathJax = {
  tex: {
    inlineMath: [["\\(", "\\)"]],
    displayMath: [["\\[", "\\]"]],
    processEscapes: true,
    processEnvironments: true
  },
  options: {
    ignoreHtmlClass: ".*|",
    processHtmlClass: "arithmatex"
  }
};
document$.subscribe(() => {
  MathJax.startup.output.clearCache();
  MathJax.typesetClear();
  MathJax.texReset();
  MathJax.typesetPromise();
});
"""

def find_vault_root(file_path: Path):
    for parent in file_path.parents:
        if (parent / ".obsidian").is_dir():
            return parent
    return None

def normalize_path(raw_path: str) -> Path:
    candidate = raw_path.strip().strip('"').strip("'")
    candidate = candidate.replace("\\/", "/")
    candidate = re.sub(r"\\([\\/ ._\-()\[\]{}])", r"\1", candidate)
    candidate = candidate.replace("\\", "")
    home_index = candidate.rfind("/home/")
    if home_index != -1:
        candidate = candidate[home_index:]
    return Path(candidate).resolve()

def get_safe_filename(stem: str) -> str:
    """Handles hierarchical numbering (e.g., 1., 1.1, 1.2) and pads them."""
    match = re.match(r'^([\d\.]+)\s*', stem)
    if match:
        num_str = match.group(1).strip('.')
        parts = [p for p in num_str.split('.') if p]
        
        # Pad each part to 2 digits (e.g., 1 -> 01)
        padded = [f"{int(p):02d}" for p in parts]
        
        # If it's a single number (e.g., "1"), pad with 00 to sort before its children
        if len(padded) == 1:
            padded.append("00")
            
        prefix = "-".join(padded)
        
        title_part = stem[match.end():]
        safe_title = re.sub(r"[^\w\-]", "-", title_part).strip("-").lower()
        return f"{prefix}-{safe_title}" if safe_title else prefix
    else:
        return re.sub(r"[^\w\-]", "-", stem).strip("-").lower()

def extract_attachments(md_content: str):
    return re.findall(r"!\[\[([^\]|]+)", md_content)

def find_asset_path(vault_root: Path, note_dir: Path, target: str):
    candidates = [
        note_dir / target,
        note_dir / "_img" / Path(target).name,
        vault_root / target,
        vault_root / "_img" / Path(target).name,
        vault_root / "99 - Meta" / "02 - Attachments" / Path(target).name,
    ]
    for c in candidates:
        if c.exists() and c.is_file():
            return c
    for c in vault_root.rglob(Path(target).name):
        if c.is_file():
            return c
    return None

def preprocess_math(md):
    code_blocks = []
    def save_code(m):
        code_blocks.append(m.group(0))
        return f"@@CODEBLOCK{len(code_blocks)-1}@@"
    md = re.sub(r'(```.*?```|`[^`]*`)', save_code, md, flags=re.DOTALL)

    def fix_math_blocks(match):
        content = match.group(1).strip()
        content = re.sub(r'\n\s*\n', '\n', content)
        return f"\n\n$$\n{content}\n$$\n\n"

    md = re.sub(r'\$\$(.*?)\$\$', fix_math_blocks, md, flags=re.DOTALL)

    for i, block in enumerate(code_blocks):
        md = md.replace(f"@@CODEBLOCK{i}@@", block)

    return md

def extract_metadata(file_path: Path):
    try:
        content = file_path.read_text(encoding="utf-8")
        fm = {}
        body = content
        if content.startswith("---"):
            parts = content.split("---", 2)
            if len(parts) >= 3:
                try:
                    fm = yaml.safe_load(parts[1]) or {}
                    body = parts[2]
                except Exception:
                    pass
        
        title = file_path.stem
        if isinstance(fm, dict) and "title" in fm:
            title = str(fm["title"])
        else:
            m = re.search(r'^#\s+(.+)$', body, re.MULTILINE)
            if m:
                title = m.group(1).strip()

        date_str = None
        if isinstance(fm, dict):
            for key in ("date", "published", "publish_date", "publish-date"):
                if key in fm and fm[key]:
                    date_str = str(fm[key])
                    break
        
        parsed_date = None
        if date_str:
            try:
                parsed_date = dt.datetime.strptime(date_str.strip()[:10], "%Y-%m-%d").date()
            except ValueError:
                pass
                
        if not parsed_date:
            m = re.search(r'(\d{4}-\d{2}-\d{2})', file_path.stem)
            if m:
                try:
                    parsed_date = dt.datetime.strptime(m.group(1), "%Y-%m-%d").date()
                except ValueError:
                    pass
                    
        if not parsed_date:
            parsed_date = dt.date.fromtimestamp(file_path.stat().st_mtime)

        return title, parsed_date
    except Exception:
        return file_path.stem, dt.date.today()

def generate_subject_index(subject_dir: Path, subject_alias: str):
    notes = []
    for p in subject_dir.glob("*.md"):
        if p.name == "index.md":
            continue
        title, date = extract_metadata(p)
        notes.append((date, title, p.name))
        
    # Sort primarily by date (newest first), then by filename (A-Z) if dates match
    notes.sort(key=lambda x: x[2]) # First alphabetical by filename
    notes.sort(key=lambda x: x[0], reverse=True) # Then newest by date
    
    lines = [f"# {subject_alias}\n", "Zoznam poznámok zoradený podľa dátumu:\n"]
    
    if not notes:
        lines.append("*Zatiaľ žiadne poznámky.*")
    else:
        for date, title, filename in notes:
            lines.append(f"- **{date.strftime('%Y-%m-%d')}** &mdash; [{title}]({filename})")
            
    (subject_dir / "index.md").write_text("\n".join(lines), encoding="utf-8")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", "-f", required=True, nargs="+", help="Paths to markdown files or directories")
    parser.add_argument("--preview", "-p", action="store_true", help="Start a local preview server without pushing")
    parser.add_argument("--no-sp", action="store_true", help="Skip SharePoint backup")
    args = parser.parse_args()

    files_to_process = []
    for f in args.file:
        p = normalize_path(f)
        if not p.exists():
            print(f"⚠️ Path not found, skipping: {p}")
            continue
            
        if p.is_dir():
            print(f"📁 Scanning directory: {p.name}")
            files_to_process.extend(list(p.rglob("*.md")))
        else:
            files_to_process.append(p)

    if not files_to_process:
        print("❌ No markdown files found to process.")
        sys.exit(1)

    docs_dir = CONTENT_REPO_DIR / "docs"
    docs_dir.mkdir(parents=True, exist_ok=True)

    # --- CLEANUP GHOST FILES ---
    subjects = ["00 - ANJ", "01 - APM", "02 - Appslab", "03 - CEQ", "04 - CIT", "05 - CME", "06 - CSX", "07 - EKO", "08 - ETS", "09 - MAT", "10 - SJL", "11 - SXT", "12 - USP"]
    valid_capitalized_folders = [s.split(" - ")[-1] for s in subjects]
    
    for item in docs_dir.iterdir():
        if item.is_file() and item.name != "index.md":
            item.unlink()
        elif item.is_dir() and item.name not in ("javascripts", *valid_capitalized_folders):
            if item.name.lower() in [f.lower() for f in valid_capitalized_folders]:
                shutil.rmtree(item)
    # ---------------------------

    config_path = CONTENT_REPO_DIR / "zensical.toml"
    config_path.write_text(ZENSICAL_CONFIG, encoding="utf-8")

    js_dir = docs_dir / "javascripts"
    js_dir.mkdir(exist_ok=True)
    (js_dir / "mathjax.js").write_text(MATHJAX_JS, encoding="utf-8")

    index_md = docs_dir / "index.md"
    if not index_md.exists():
        index_md.write_text("# SPSITKNM Notes\n\nVitajte v poznámkach. Vyberte si predmet z horného menu.", encoding="utf-8")

    affected_subjects = set()

    # --- PROCESS FILES ---
    for file_path in files_to_process:
        vault_root = find_vault_root(file_path)
        if not vault_root:
            print(f"❌ Could not find vault root for {file_path.name}. Skipping.")
            continue

        print(f"📄 Staging note: {file_path.name}")
        md_content = file_path.read_text(encoding="utf-8")
        md_content = preprocess_math(md_content)
        
        subject_folder_name = next((p.name for p in file_path.parents if p.name in subjects), "misc")
        subject_alias = subject_folder_name.split(" - ")[-1]
        
        safe_title = get_safe_filename(file_path.stem)
        
        subject_dir = docs_dir / subject_alias
        subject_dir.mkdir(parents=True, exist_ok=True)

        staged_path = subject_dir / f"{safe_title}.md"
        staged_path.write_text(md_content, encoding="utf-8")

        attachments = extract_attachments(md_content)
        if attachments:
            print(f"🖼️  Copying {len(attachments)} attachments for {file_path.name}...")
            for att in attachments:
                src = find_asset_path(vault_root, file_path.parent, att.strip())
                if src:
                    shutil.copy2(src, subject_dir / src.name)
                else:
                    print(f"   ⚠️  Attachment not found: {att}")
                    
        affected_subjects.add((subject_dir, subject_alias))

    # --- GENERATE INDEXES ---
    for subject_dir, subject_alias in affected_subjects:
        generate_subject_index(subject_dir, subject_alias)

    print("🔨 Building static site...")
    try:
        with tempfile.TemporaryDirectory(prefix="zensical_build_") as temp_dir:
            temp_path = Path(temp_dir)
            
            shutil.copy2(config_path, temp_path / "zensical.toml")
            shutil.copytree(docs_dir, temp_path / "docs")
            
            subprocess.run(["zensical", "build"], cwd=str(temp_path), check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
            
            built_site = temp_path / "site"
            target_site = CONTENT_REPO_DIR / "site"
            if target_site.exists():
                shutil.rmtree(target_site)
            shutil.copytree(built_site, target_site)

            if args.preview:
                print("\n👀 Starting local preview server...")
                print("   Press Ctrl+C to stop the preview and exit.")
                if files_to_process:
                    first_file = files_to_process[0]
                    first_subject = next((p.name for p in first_file.parents if p.name in subjects), "misc").split(" - ")[-1]
                    first_title = get_safe_filename(first_file.stem)
                    print(f"   Example URL: http://127.0.0.1:8000/{first_subject}/{first_title}.html\n")
                try:
                    subprocess.run(["zensical", "serve"], cwd=str(temp_path))
                except KeyboardInterrupt:
                    print("\n🛑 Preview stopped.")
                sys.exit(0)

    except FileNotFoundError:
        print("❌ Zensical is not installed. Run: pip install zensical")
        sys.exit(1)
    except subprocess.CalledProcessError as e:
        print(f"❌ Zensical build failed:\n{e.stdout.decode()}")
        sys.exit(1)

    print("🚀 Pushing to Content Repo Git...")
    try:
        subprocess.run(["git", "rev-parse", "--is-inside-work-tree"], cwd=str(CONTENT_REPO_DIR), check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        subprocess.run(["git", "add", "."], cwd=str(CONTENT_REPO_DIR), check=True)
        
        status = subprocess.run(["git", "status", "--porcelain"], cwd=str(CONTENT_REPO_DIR), capture_output=True, text=True)
        if status.stdout.strip():
            commit_msg = f"Publish: {len(files_to_process)} note(s)" if len(files_to_process) > 1 else f"Publish: {files_to_process[0].name}"
            subprocess.run(["git", "commit", "-m", commit_msg], cwd=str(CONTENT_REPO_DIR), check=True)
            subprocess.run(["git", "push"], cwd=str(CONTENT_REPO_DIR), check=True)
            print("   ✅ Git push successful.")
        else:
            print("   ℹ️ No changes to push.")
    except subprocess.CalledProcessError:
        print(f"   ⚠️ Git push failed. Is {CONTENT_REPO_DIR} a valid Git repository?")

    if not args.no_sp and sp_backup:
        print("💾 Running SharePoint backup...")
        for file_path in files_to_process:
            print(f"   Backing up: {file_path.name}")
            original_args = sys.argv.copy()
            sys.argv = ["obsidian_to_sharepoint.py", "--file", str(file_path), "--scopes", "Sites.ReadWrite.All"]
            try:
                sp_backup.main()
            except SystemExit:
                pass
            finally:
                sys.argv = original_args

    print("\n✅ All done!")

if __name__ == "__main__":
    main()