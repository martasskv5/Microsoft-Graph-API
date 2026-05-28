#!/usr/bin/env python3
"""
Excalidraw to PNG Exporter for Obsidian

Supports:
- Legacy .excalidraw files (plain JSON)
- Modern .excalidraw.md files (with compressed or uncompressed JSON)
- Batch processing entire vaults

Requirements:
    pip install playwright pyyaml lzstring
    playwright install firefox
    npm install -g excalidraw-brute-export-cli

Usage:
    python excalidraw_to_png.py /path/to/vault --output /path/to/output
    python excalidraw_to_png.py /path/to/drawing.excalidraw.md --output drawing.png
"""

import argparse
import base64
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Dict, List, Optional

try:
    import yaml
except ImportError:
    yaml = None

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sync_playwright = None


def extract_excalidraw_json(file_path: Path) -> Optional[Dict]:
    """Extract Excalidraw scene data from any supported file format."""
    content = file_path.read_text(encoding="utf-8", errors="ignore")

    # Legacy .excalidraw file (pure JSON)
    if file_path.suffix == ".excalidraw" or content.strip().startswith("{"):
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            pass

    # Modern .md format with frontmatter
    if yaml and content.startswith("---"):
        parts = content.split("---", 2)
        if len(parts) >= 3:
            try:
                frontmatter = yaml.safe_load(parts[1])
                if frontmatter and "excalidraw-plugin" in frontmatter:
                    return _extract_json_from_md_body(content)
            except yaml.YAMLError:
                pass

    return _extract_json_from_md_body(content)


def _extract_json_from_md_body(content: str) -> Optional[Dict]:
    """Extract Excalidraw JSON from markdown body."""
    # Compressed JSON block
    compressed_match = re.search(
        r"```compressed-json\n(.*?)\n```",
        content,
        re.DOTALL
    )
    if compressed_match:
        compressed_data = compressed_match.group(1).strip().replace("\n", "").replace("\r", "")
        try:
            import lzstring
            decompressed = lzstring.LZString.decompressFromBase64(compressed_data)
            return json.loads(decompressed)
        except Exception:
            pass

    # Uncompressed JSON block
    json_match = re.search(
        r"```json\n(.*?)\n```",
        content,
        re.DOTALL
    )
    if json_match:
        try:
            return json.loads(json_match.group(1).strip())
        except json.JSONDecodeError:
            pass

    # excalidraw-json code block
    excalidraw_match = re.search(
        r"```excalidraw-json\n(.*?)\n```",
        content,
        re.DOTALL
    )
    if excalidraw_match:
        try:
            return json.loads(excalidraw_match.group(1).strip())
        except json.JSONDecodeError:
            pass

    return None


class ExcalidrawExporter:
    """Export Excalidraw drawings to PNG."""

    def __init__(
        self,
        vault_path: str,
        output_dir: Optional[str] = None,
        scale: float = 2.0,
        padding: int = 10,
        dark_mode: bool = False,
        transparent: bool = False
    ):
        self.vault_path = Path(vault_path).resolve()
        self.output_dir = (
            Path(output_dir).resolve()
            if output_dir
            else self.vault_path / "exports"
        )
        self.scale = scale
        self.padding = padding
        self.dark_mode = dark_mode
        self.transparent = transparent
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def find_excalidraw_files(self) -> List[Path]:
        """Find all Excalidraw files in the vault."""
        files = []
        files.extend(self.vault_path.rglob("*.excalidraw"))
        files.extend(self.vault_path.rglob("*.excalidraw.md"))

        # Also check regular .md files
        for md_file in self.vault_path.rglob("*.md"):
            if md_file not in files:
                content = md_file.read_text(encoding="utf-8", errors="ignore")
                if "excalidraw-plugin" in content:
                    files.append(md_file)

        return sorted(set(files))

    def export_with_cli(
        self,
        file_path: Path,
        output_path: Optional[Path] = None
    ) -> Optional[Path]:
        """Export using excalidraw-brute-export-cli (most reliable)."""
        scene_data = extract_excalidraw_json(file_path)
        if not scene_data:
            print(f"  Warning: Could not parse {file_path}")
            return None

        if output_path is None:
            output_path = self.output_dir / f"{file_path.stem}.png"

        # Write temp JSON
        with tempfile.NamedTemporaryFile(
            mode="w",
            suffix=".excalidraw",
            delete=False
        ) as f:
            json.dump(scene_data, f)
            temp_json = Path(f.name)

        try:
            cmd = [
                "npx", "excalidraw-brute-export-cli",
                "-i", str(temp_json),
                "--scale", str(self.scale),
                "--format", "png",
                "-o", str(output_path),
            ]

            if self.dark_mode:
                cmd.extend(["--dark-mode", "1"])
            else:
                cmd.extend(["--dark-mode", "0"])

            if self.transparent:
                cmd.extend(["--background", "0"])
            else:
                cmd.extend(["--background", "1"])

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120
            )

            if result.returncode != 0:
                print(f"  CLI Error: {result.stderr}")
                return None

            print(f"  Exported: {output_path}")
            return output_path

        finally:
            temp_json.unlink()

    def export_with_playwright(
        self,
        file_path: Path,
        output_path: Optional[Path] = None
    ) -> Optional[Path]:
        """Export using Playwright headless browser."""
        if sync_playwright is None:
            print("  Playwright not installed. Use: pip install playwright")
            return None

        scene_data = extract_excalidraw_json(file_path)
        if not scene_data:
            print(f"  Warning: Could not parse {file_path}")
            return None

        if output_path is None:
            output_path = self.output_dir / f"{file_path.stem}.png"

        scene_json = json.dumps(scene_data)
        theme = "dark" if self.dark_mode else "light"
        bg_color = "transparent" if self.transparent else ("#1e1e1e" if self.dark_mode else "#ffffff")

        # Build HTML
        html_lines = [
            "<!DOCTYPE html>",
            "<html>",
            "<head>",
            '    <meta charset="UTF-8">',
            "    <style>",
            f"        body {{ margin: 0; padding: 0; background: {bg_color}; }}",
            "        #export-container { width: 100%; height: 100vh; }",
            "    </style>",
            "</head>",
            "<body>",
            '    <div id="export-container"></div>',
            '    <script src="https://unpkg.com/react@18/umd/react.production.min.js"></script>',
            '    <script src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>',
            '    <script src="https://unpkg.com/@excalidraw/excalidraw@0.17.0/dist/excalidraw.production.min.js"></script>',
            "    <script>",
            f"        const sceneData = {scene_json};",
            "        const App = () => {",
            "            return React.createElement(",
            "                window.ExcalidrawLib.Excalidraw,",
            "                {",
            "                    initialData: sceneData,",
            "                    UIOptions: { canvasActions: { loadScene: false } },",
            f'                    theme: "{theme}"',
            "                }",
            "            );",
            "        };",
            '        const root = ReactDOM.createRoot(document.getElementById("export-container"));',
            "        root.render(React.createElement(App));",
            "    </script>",
            "</body>",
            "</html>",
        ]

        html_content = "\n".join(html_lines)
        temp_html = self.output_dir / "_temp_export.html"
        temp_html.write_text(html_content, encoding="utf-8")

        with sync_playwright() as p:
            browser = p.firefox.launch(headless=True)
            page = browser.new_page()

            try:
                page.goto(f"file://{temp_html}")
                page.wait_for_timeout(3000)

                canvas = page.query_selector("canvas")
                if canvas:
                    canvas.screenshot(path=str(output_path))
                    print(f"  Exported: {output_path}")
                    return output_path
                else:
                    print("  No canvas found")
                    return None

            finally:
                browser.close()
                if temp_html.exists():
                    temp_html.unlink()

    def batch_export(self, use_cli: bool = True) -> List[Path]:
        """Export all Excalidraw files in the vault."""
        files = self.find_excalidraw_files()
        exported = []

        print(f"Found {len(files)} Excalidraw files")

        export_method = (
            self.export_with_cli if use_cli else self.export_with_playwright
        )

        for i, file_path in enumerate(files, 1):
            print(f"[{i}/{len(files)}] Processing: {file_path.name}")
            try:
                output_path = export_method(file_path)
                if output_path:
                    exported.append(output_path)
            except Exception as e:
                print(f"  Error: {e}")

        return exported


def main():
    parser = argparse.ArgumentParser(
        description="Export Excalidraw drawings from Obsidian to PNG"
    )
    parser.add_argument(
        "input",
        help="Path to Obsidian vault or specific Excalidraw file"
    )
    parser.add_argument(
        "--output", "-o",
        help="Output directory or file path"
    )
    parser.add_argument(
        "--scale", "-s",
        type=float,
        default=2.0,
        help="Export scale factor (default: 2.0)"
    )
    parser.add_argument(
        "--padding", "-p",
        type=int,
        default=10,
        help="Export padding in pixels (default: 10)"
    )
    parser.add_argument(
        "--dark-mode", "-d",
        action="store_true",
        help="Export in dark mode"
    )
    parser.add_argument(
        "--transparent", "-t",
        action="store_true",
        help="Transparent background"
    )
    parser.add_argument(
        "--playwright",
        action="store_true",
        help="Use Playwright instead of CLI tool"
    )

    args = parser.parse_args()

    input_path = Path(args.input).resolve()

    if not input_path.exists():
        print(f"Error: Path not found: {input_path}")
        sys.exit(1)

    try:
        exporter = ExcalidrawExporter(
            vault_path=str(
                input_path if input_path.is_dir() else input_path.parent
            ),
            output_dir=args.output,
            scale=args.scale,
            padding=args.padding,
            dark_mode=args.dark_mode,
            transparent=args.transparent
        )

        if input_path.is_file():
            if args.playwright:
                output = exporter.export_with_playwright(input_path)
            else:
                output = exporter.export_with_cli(input_path)

            if output:
                print(f"\nSuccessfully exported to: {output}")
            else:
                print("\nExport failed")
                sys.exit(1)
        else:
            exported = exporter.batch_export(use_cli=not args.playwright)
            print(f"\nExported {len(exported)} files to {exporter.output_dir}")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
