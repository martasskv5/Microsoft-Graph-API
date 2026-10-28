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
import os
from pathlib import Path

import yaml

try:
    import obsidian_to_sharepoint as sp_backup
except ImportError:
    print("⚠️  Warning: obsidian_to_sharepoint.py not found. Skipping SharePoint backup.")
    sp_backup = None

# --- CONFIGURATION ---
CONTENT_REPO_DIR = Path("/home/martinko/Projects/spsknm-notes").resolve()

# Moved # NAV_PLACEHOLDER ABOVE [project.extra] to ensure TOML parses it correctly
BASE_ZENSICAL_CONFIG = """\
[project]
site_name = "notes.michalec.it"
site_url = "https://notes.michalec.it/"
copyright = "Copyright &copy; 2026 Martin Michalec"
use_directory_urls = false
extra_javascript = [
    "javascripts/mathjax.js",
    "https://unpkg.com/mathjax@3/es5/tex-mml-chtml.js",
    "https://unpkg.com/mermaid@10/dist/mermaid.min.js"
]
markdown_extensions = [
    "tables",
    "admonition",
    "nl2br",
    "md_in_html",
    "attr_list",
    { "pymdownx.arithmatex" = { generic = true } },
    { "pymdownx.superfences" = { custom_fences = [
        { name = "mermaid", class = "mermaid", format = "pymdownx.superfences.fence_code_format" }
    ] } },
    { "toc" = { toc_depth = "2-6" } }
]

# NAV_PLACEHOLDER

[project.extra]
generator = false

[project.theme]
name = "material"
language = "sk"
features = [
    "navigation.footer",
    "navigation.tabs",
    "navigation.sections",
    "content.code.copy",
    "navigation.instant",
    "navigation.instant.preview"
]

[[project.theme.palette]]
media = "(prefers-color-scheme)"
toggle.icon = "lucide/sun-moon"
toggle.name = "Switch to light mode"

[[project.theme.palette]]
media = "(prefers-color-scheme: light)"
scheme = "default"
toggle.icon = "lucide/sun"
toggle.name = "Switch to dark mode"

[[project.theme.palette]]
media = "(prefers-color-scheme: dark)"
scheme = "slate"
toggle.icon = "lucide/moon"
toggle.name = "Switch to system preference"

[[project.extra.social]]
icon = "fontawesome/brands/github"
link = "https://github.com/martasskv5"
"""

DEFAULT_MATHJAX_JS = r"""window.MathJax = {
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

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'(\d+)', s)]

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

def convert_obsidian_links(md_content: str, current_subject: str, all_files_map: dict) -> str:
    code_blocks = []
    def save_code(m):
        code_blocks.append(m.group(0))
        return f"@@CODEBLOCK{len(code_blocks)-1}@@"
    md = re.sub(r'(```.*?```|`[^`]*`)', save_code, md_content, flags=re.DOTALL)

    def repl(match):
        inner = match.group(1).strip()
        if "|" in inner:
            target_part, display_text = inner.split("|", 1)
        else:
            target_part = inner
            display_text = inner.split("#")[0].strip()
        
        target_part = target_part.strip()
        display_text = display_text.strip()
        
        anchor = ""
        if "#" in target_part:
            target_part, anchor = target_part.split("#", 1)
            anchor = "#" + anchor.strip()
        
        target_part = target_part.strip()
        target_lower = target_part.lower()
        
        resolved = None
        if "/" in target_lower:
            subj, name = target_lower.split("/", 1)
            subj = subj.strip()
            name = name.strip()
            if (subj, name) in all_files_map:
                resolved = all_files_map[(subj, name)]
                if subj != current_subject:
                    resolved = f"../{subj}/{resolved}"
        else:
            if (current_subject, target_lower) in all_files_map:
                resolved = all_files_map[(current_subject, target_lower)]
            else:
                for (subj, name), path in all_files_map.items():
                    if name == target_lower:
                        resolved = f"../{subj}/{path}"
                        break
                        
        if resolved:
            return f"[{display_text}]({resolved}{anchor}){{ data-preview }}"
        else:
            safe_target = re.sub(r"[^\w\-]", "-", target_part).strip("-").lower()
            return f"[{display_text}]({safe_target}.md{anchor}){{ data-preview }}"

    md = re.sub(r"\[\[([^\]]+)\]\]", repl, md)

    for i, block in enumerate(code_blocks):
        md = md.replace(f"@@CODEBLOCK{i}@@", block)

    return md

def fix_header_hierarchy(md_content: str) -> str:
    lines = md_content.split('\n')
    h1_count = 0
    out_lines = []
    
    for line in lines:
        if line.startswith('# ') and not line.startswith('## '):
            h1_count += 1
            if h1_count > 1:
                line = '## ' + line[2:]
        out_lines.append(line)
        
    return '\n'.join(out_lines)

def convert_obsidian_callouts(md_content: str) -> str:
    callout_re = re.compile(r"^>\s*\[!([^\]]+)\][+-]?\s*(.*?)\n((?:^>.*(?:\n|$))*)", re.MULTILINE)

    def repl(match):
        kind = match.group(1).lower()
        title = match.group(2).strip()
        body = match.group(3) or ""

        type_map = {
            "hint": "tip", "check": "success", "done": "success",
            "help": "question", "faq": "question", "caution": "warning",
            "attention": "warning", "fail": "failure", "missing": "failure",
            "error": "danger", "cite": "quote"
        }
        md_kind = type_map.get(kind, kind)
        title_part = f' "{title}"' if title else ""

        body_lines = []
        for line in body.splitlines():
            if line.startswith(">"):
                line = line[1:]
                if line.startswith(" "):
                    line = line[1:]
            body_lines.append("    " + line)

        body_str = "\n".join(body_lines)
        return f"!!! {md_kind}{title_part}\n{body_str}\n\n"

    return callout_re.sub(repl, md_content)

def fix_markdown_formatting(md_content: str) -> str:
    lines = md_content.split('\n')
    out_lines = []
    in_code_block = False

    for line in lines:
        stripped = line.strip()

        if stripped.startswith('```') or stripped.startswith('~~~'):
            in_code_block = not in_code_block
            if not in_code_block:
                out_lines.append(line)
                out_lines.append('')
            else:
                if out_lines and out_lines[-1].strip() != '':
                    out_lines.append('')
                out_lines.append(line)
            continue

        if in_code_block:
            out_lines.append(line)
            continue

        if re.match(r'^[ \t]*(?:[-*+]|\d+\.)\s+', line):
            if out_lines and out_lines[-1].strip() != '':
                if not re.match(r'^[ \t]*(?:[-*+]|\d+\.)\s+', out_lines[-1]):
                    if not re.match(r'^[ \t]+', out_lines[-1]):
                        out_lines.append('')

        out_lines.append(line)

    cleaned = []
    for line in out_lines:
        if line.strip() == '' and cleaned and cleaned[-1].strip() == '':
            continue
        cleaned.append(line)

    return '\n'.join(cleaned)

def preprocess_math(md):
    md = re.sub(r'\\(xrightarrow|xleftarrow)\{([^{}]*\\\\[^{}]*)\}', r'\\\1{\\substack{\2}}', md)

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
        
    notes.sort(key=lambda x: natural_sort_key(x[1]))
    
    lines = [f"# {subject_alias}\n", "Zoznam poznámok:\n"]
    
    if not notes:
        lines.append("*Zatiaľ žiadne poznámky.*")
    else:
        for date, title, filename in notes:
            lines.append(f"- **{date.strftime('%Y-%m-%d')}** &mdash; [{title}]({filename}){{ data-preview }}")
            
    (subject_dir / "index.md").write_text("\n".join(lines), encoding="utf-8")

def generate_main_index(docs_dir: Path):
    subjects = ["00 - ANJ", "01 - APM", "02 - Appslab", "03 - CEQ", "04 - CIT", "05 - CME", "06 - CSX", "07 - EKO", "08 - ETS", "09 - MAT", "10 - SJL", "11 - SXT", "12 - USP"]
    alias_to_full = {s.split(" - ")[-1]: s for s in subjects}
    
    existing_aliases = []
    for d in docs_dir.iterdir():
        if d.is_dir() and d.name in alias_to_full:
            existing_aliases.append(d.name)
            
    existing_aliases.sort()
    
    icons = {
        "ANJ": "🇬🇧", "APM": "📐", "Appslab": "📱", "CEQ": "🧪", "CIT": "💻",
        "CME": "⚙️", "CSX": "🖥️", "EKO": "📈", "ETS": "🔌", "MAT": "🧮",
        "SJL": "📖", "SXT": "🌐", "USP": "🏛️"
    }
    
    html_out = [
        "# SPSITKNM Notes\n",
        "Vitajte v poznámkach. Vyberte si predmet:\n\n",
        '<div class="grid" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 1.5rem; margin-top: 2rem;">\n'
    ]
    
    for alias in existing_aliases:
        icon = icons.get(alias, "📁")
        link = f"{alias}/index.html"
        
        html_out.append(f'''\
    <a href="{link}" style="display: flex; flex-direction: column; padding: 1.5rem; border: 1px solid var(--md-default-fg-color--lightest); border-radius: 8px; text-decoration: none; background-color: var(--md-code-bg-color); transition: transform 0.2s, box-shadow 0.2s, border-color 0.2s; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
        <h2 style="margin-top: 0; margin-bottom: 0.5rem; font-size: 1.5rem; color: var(--md-primary-fg-color); display: flex; align-items: center; gap: 0.5rem;">
            <span>{icon}</span> {alias}
        </h2>
        <p style="margin: 0; color: var(--md-default-fg-color--light); font-size: 0.9rem;">Poznámky pre predmet</p>
    </a>
''')
    
    html_out.append('</div>\n')
    
    (docs_dir / "index.md").write_text("".join(html_out), encoding="utf-8")

def regenerate_config(config_path: Path, docs_dir: Path, run_order: dict):
    nav_items = []
    for subject_dir in sorted(docs_dir.iterdir(), key=lambda p: p.name):
        if subject_dir.is_dir() and subject_dir.name not in ("javascripts",):
            subject_alias = subject_dir.name
            files = list(subject_dir.glob("*.md"))
            
            ordered_files = []
            for fname in run_order.get(subject_alias, []):
                fpath = subject_dir / fname
                if fpath in files:
                    ordered_files.append(fpath)
            
            rest_files = [f for f in files if f not in ordered_files]
            rest_files.sort(key=lambda p: natural_sort_key(extract_metadata(p)[0]))
            
            final_files = ordered_files + rest_files
            
            sub_nav = []
            for p in final_files:
                if p.name == "index.md":
                    continue
                title, _ = extract_metadata(p)
                title_clean = title.replace('"', '\\"')
                sub_nav.append(f'{{ "{title_clean}" = "{subject_alias}/{p.name}" }}')
                
            sub_nav.insert(0, f'{{ "Index" = "{subject_alias}/index.md" }}')
                
            if sub_nav:
                files_str = ",\n                ".join(sub_nav)
                nav_items.append(f'    {{ "{subject_alias}" = [\n                {files_str}\n            ] }}')

    nav_str = ",\n".join(nav_items)
    nav_block = f"nav = [\n{nav_str}\n]" if nav_items else ""

    full_content = BASE_ZENSICAL_CONFIG.replace("# NAV_PLACEHOLDER", nav_block)
    config_path.write_text(full_content, encoding="utf-8")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", "-f", required=True, nargs="+", action="extend", help="Paths to markdown files or directories")
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

    config_path = CONTENT_REPO_DIR / "zensical.toml"

    js_dir = docs_dir / "javascripts"
    js_dir.mkdir(exist_ok=True)
    
    mathjax_path = js_dir / "mathjax.js"
    if not mathjax_path.exists():
        mathjax_path.write_text(DEFAULT_MATHJAX_JS, encoding="utf-8")

    affected_subjects = set()
    subjects = ["00 - ANJ", "01 - APM", "02 - Appslab", "03 - CEQ", "04 - CIT", "05 - CME", "06 - CSX", "07 - EKO", "08 - ETS", "09 - MAT", "10 - SJL", "11 - SXT", "12 - USP"]
    
    run_order = {}

    # --- PASS 1: Stage raw files ---
    for file_path in files_to_process:
        vault_root = find_vault_root(file_path)
        if not vault_root:
            print(f"❌ Could not find vault root for {file_path.name}. Skipping.")
            continue

        print(f"📄 Staging note: {file_path.name}")
        md_content = file_path.read_text(encoding="utf-8")
        
        subject_folder_name = next((p.name for p in file_path.parents if p.name in subjects), "misc")
        subject_alias = subject_folder_name.split(" - ")[-1]
        
        safe_title = re.sub(r"[^\w\-]", "-", file_path.stem).strip("-").lower()
        staged_name = f"{safe_title}.md"
        
        run_order.setdefault(subject_alias, []).append(staged_name)
        
        subject_dir = docs_dir / subject_alias
        subject_dir.mkdir(parents=True, exist_ok=True)

        staged_path = subject_dir / staged_name
        staged_path.write_text(md_content, encoding="utf-8")
        
        stat = file_path.stat()
        os.utime(staged_path, (stat.st_atime, stat.st_mtime))

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

    # --- PASS 2: Post-process and convert links ---
    print("🔗 Resolving Obsidian wikilinks...")
    all_files_map = {}
    for subj_dir in docs_dir.iterdir():
        if subj_dir.is_dir() and subj_dir.name not in ("javascripts", "assets"):
            for md_file in subj_dir.glob("*.md"):
                if md_file.name == "index.md":
                    continue
                title, _ = extract_metadata(md_file)
                all_files_map[(subj_dir.name, title.lower())] = md_file.name
                safe_name = re.sub(r"[^\w\-]", "-", title).strip("-").lower()
                all_files_map[(subj_dir.name, safe_name)] = md_file.name
                orig_stem = md_file.stem.lower()
                all_files_map[(subj_dir.name, orig_stem)] = md_file.name

    for subj_dir in docs_dir.iterdir():
        if subj_dir.is_dir() and subj_dir.name not in ("javascripts", "assets"):
            for md_file in subj_dir.glob("*.md"):
                if md_file.name == "index.md":
                    continue
                content = md_file.read_text(encoding="utf-8")
                
                content = fix_header_hierarchy(content)
                content = convert_obsidian_links(content, subj_dir.name, all_files_map)
                content = convert_obsidian_callouts(content)
                content = fix_markdown_formatting(content)
                content = preprocess_math(content)
                
                has_h1 = re.search(r'^#\s+', content, re.MULTILINE) is not None
                if not has_h1:
                    if content.startswith("---"):
                        parts = content.split("---", 2)
                        if len(parts) >= 3:
                            if "title:" not in parts[1]:
                                content = f"---\ntitle: \"{md_file.stem}\"\n{parts[1]}---{parts[2]}"
                    else:
                        content = f"---\ntitle: \"{md_file.stem}\"\n---\n\n{content}"
                        
                md_file.write_text(content, encoding="utf-8")

    for subject_dir, subject_alias in affected_subjects:
        generate_subject_index(subject_dir, subject_alias)

    generate_main_index(docs_dir)
    regenerate_config(config_path, docs_dir, run_order)

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
                print(f"   URL: http://127.0.0.1:8000/index.html\n")
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