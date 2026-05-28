# Excalidraw to PNG Export Guide for Obsidian

## Overview

There are several approaches to export Excalidraw drawings from Obsidian to PNG using Python:

## Method 1: Using excalidraw-brute-export-cli (Recommended)

This is the most reliable method as it uses a real headless browser with the actual Excalidraw rendering engine.

### Installation

```bash
# Install Node.js tool globally
npm install -g excalidraw-brute-export-cli

# Install browser dependencies
npx playwright install-deps
npx playwright install firefox
```

### Usage

```bash
# Export a single file
npx excalidraw-brute-export-cli \
  -i ./drawing.excalidraw \
  --background 1 \
  --embed-scene 0 \
  --dark-mode 0 \
  --scale 2 \
  --format png \
  -o ./drawing.png
```

### Python Wrapper Script

```python
#!/usr/bin/env python3
import subprocess
import json
import re
from pathlib import Path
import argparse


def extract_excalidraw_json(md_file: Path) -> dict:
    content = md_file.read_text(encoding='utf-8')
    
    # Try compressed JSON first
    compressed_match = re.search(r'```compressed-json\n(.*?)\n```', content, re.DOTALL)
    if compressed_match:
        try:
            import lzstring
            compressed = compressed_match.group(1).strip().replace('\n', '')
            decompressed = lzstring.LZString.decompressFromBase64(compressed)
            return json.loads(decompressed)
        except ImportError:
            print("Warning: lzstring not installed.")
            print("  pip install lzstring")
    
    # Try uncompressed JSON
    json_match = re.search(r'```json\n(.*?)\n```', content, re.DOTALL)
    if json_match:
        return json.loads(json_match.group(1).strip())
    
    raise ValueError(f"Could not extract Excalidraw data from {md_file}")


def export_file(input_file: Path, output_file: Path, scale: float = 2.0, dark_mode: bool = False):
    # If markdown file, extract JSON to temp file
    if input_file.suffix == '.md':
        scene_data = extract_excalidraw_json(input_file)
        temp_json = output_file.parent / f"_{input_file.stem}.excalidraw"
        temp_json.write_text(json.dumps(scene_data), encoding='utf-8')
        input_file = temp_json
    
    cmd = [
        'npx', 'excalidraw-brute-export-cli',
        '-i', str(input_file),
        '--scale', str(scale),
        '--format', 'png',
        '-o', str(output_file),
    ]
    
    if dark_mode:
        cmd.extend(['--dark-mode', '1'])
    else:
        cmd.extend(['--dark-mode', '0'])
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"Error: {result.stderr}")
        return False
    
    if input_file.suffix == '.excalidraw' and input_file.name.startswith('_'):
        input_file.unlink()
    
    print(f"Exported: {output_file}")
    return True


def batch_export(vault_path: Path, output_dir: Path):
    files = list(vault_path.rglob('*.excalidraw')) + list(vault_path.rglob('*.excalidraw.md'))
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    for file in files:
        output = output_dir / f"{file.stem}.png"
        export_file(file, output)
    
    print(f"\nExported {len(files)} files to {output_dir}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('input', help='Input file or vault directory')
    parser.add_argument('--output', '-o', help='Output file or directory')
    parser.add_argument('--scale', '-s', type=float, default=2.0)
    parser.add_argument('--dark-mode', '-d', action='store_true')
    args = parser.parse_args()
    
    input_path = Path(args.input)
    
    if input_path.is_file():
        output = Path(args.output) if args.output else input_path.with_suffix('.png')
        export_file(input_path, output, args.scale, args.dark_mode)
    else:
        output = Path(args.output) if args.output else input_path / 'exports'
        batch_export(input_path, output)
```

## Method 2: Using Playwright Directly

```python
from playwright.sync_api import sync_playwright
import json


def export_with_playwright(scene_json: dict, output_path: str, scale: float = 2.0):
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <script src="https://unpkg.com/react@18/umd/react.production.min.js"></script>
        <script src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>
        <script src="https://unpkg.com/@excalidraw/excalidraw@0.17.0/dist/excalidraw.production.min.js"></script>
    </head>
    <body>
        <div id="root"></div>
        <script>
            const sceneData = {json.dumps(scene_json)};
            
            const App = () => React.createElement(
                window.ExcalidrawLib.Excalidraw,
                {{ initialData: sceneData }}
            );
            
            ReactDOM.createRoot(document.getElementById('root'))
                .render(React.createElement(App));
        </script>
    </body>
    </html>
    """
    
    with sync_playwright() as p:
        browser = p.firefox.launch()
        page = browser.new_page()
        
        temp_html = '/tmp/excalidraw_export.html'
        with open(temp_html, 'w') as f:
            f.write(html_content)
        
        page.goto(f'file://{temp_html}')
        page.wait_for_timeout(3000)
        
        canvas = page.query_selector('canvas')
        if canvas:
            canvas.screenshot(path=output_path, scale='css')
        
        browser.close()
```

## Method 3: Obsidian Plugin Auto-Export (No Python Needed)

The Obsidian Excalidraw plugin has built-in auto-export functionality:

1. Open Obsidian Settings -> Community Plugins -> Excalidraw
2. Go to **Auto-export Settings**
3. Enable **Auto-export PNG**
4. Enable **Keep the .PNG in sync**
5. Optionally set **PNG scale factor**

This will automatically create .png files alongside your .excalidraw.md files every time you save.

You can also set per-file overrides using frontmatter:
```yaml
---
excalidraw-autoexport: png
excalidraw-export-pngscale: 2
excalidraw-export-transparent: false
---
```

## File Format Notes

### Obsidian Excalidraw File Structure

Modern Obsidian Excalidraw files (.excalidraw.md) are markdown files with:

1. YAML Frontmatter with `excalidraw-plugin: parsed` or `excalidraw-plugin: raw`
2. Text Elements section (human-readable text)
3. Excalidraw Data section containing either:
   - ` ```json ` - Uncompressed Excalidraw scene JSON
   - ` ```compressed-json ` - LZ-String compressed JSON (if compression enabled in settings)

### Extracting Scene Data

```python
import json
import re


def parse_obsidian_excalidraw(filepath):
    with open(filepath, 'r') as f:
        content = f.read()
    
    # Check for compressed data
    compressed = re.search(r'```compressed-json\n(.*?)\n```', content, re.DOTALL)
    if compressed:
        import lzstring
        data = compressed.group(1).replace('\n', '')
        json_str = lzstring.LZString.decompressFromBase64(data)
        return json.loads(json_str)
    
    # Check for uncompressed JSON
    json_block = re.search(r'```json\n(.*?)\n```', content, re.DOTALL)
    if json_block:
        return json.loads(json_block.group(1))
    
    return None
```

## Dependencies

```bash
# For Method 1 (CLI wrapper)
pip install lzstring
npm install -g excalidraw-brute-export-cli

# For Method 2 (Playwright)
pip install playwright
playwright install firefox
```

## Tips

1. **Scale Factor**: Use 2.0 or higher for crisp images suitable for presentations
2. **Dark Mode**: Set --dark-mode 1 for dark-themed exports
3. **Transparency**: Use --background 0 with the CLI for transparent PNGs
4. **Batch Processing**: Use the Python wrapper to process entire vaults
5. **Embedded Images**: The exported PNG will include any images embedded in the drawing