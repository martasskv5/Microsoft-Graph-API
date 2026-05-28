from pathlib import Path
from excalidraw_to_png import ExcalidrawExporter
from PIL import Image

input_file = Path("/home/martinko/OneDrive - MSFT/Skola/SPSITknm 25-26/99 - Meta/Excalidraw/Drawing 2025-10-07 17.43.56.excalidraw.md").resolve()
output_dir = Path("./test_outputs")
output_dir.mkdir(parents=True, exist_ok=True)

exporter = ExcalidrawExporter(vault_path=str(input_file.parent), output_dir=str(output_dir), scale=2.0)

print("Input:", input_file)

# Try CLI
print("Trying CLI export...")
cli_out = None
try:
    cli_out = exporter.export_with_cli(input_file, output_dir / (input_file.stem + "-cli.png"))
    print("CLI result:", cli_out)
    if cli_out:
        p = Path(cli_out)
        print("Exists:", p.exists(), "Size:", p.stat().st_size)
        img = Image.open(p)
        print("Dimensions:", img.size, "Mode:", img.mode)
except Exception as e:
    print("CLI exception:", e)

# Try Playwright
print("Trying Playwright export...")
pw_out = None
try:
    pw_out = exporter.export_with_playwright(input_file, output_dir / (input_file.stem + "-pw.png"))
    print("Playwright result:", pw_out)
    if pw_out:
        p = Path(pw_out)
        print("Exists:", p.exists(), "Size:", p.stat().st_size)
        img = Image.open(p)
        print("Dimensions:", img.size, "Mode:", img.mode)
except Exception as e:
    print("Playwright exception:", e)

print("Done")
