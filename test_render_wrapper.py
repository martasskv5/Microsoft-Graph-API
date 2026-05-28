from pathlib import Path
from obsidian_to_sharepoint import render_excalidraw_to_png

input_file = Path("/home/martinko/OneDrive - MSFT/Skola/SPSITknm 25-26/99 - Meta/Excalidraw/Drawing 2025-10-07 17.43.56.excalidraw.md")
output = Path("./test_outputs")
output.mkdir(parents=True, exist_ok=True)

out_path = output / (input_file.stem + "-wrapper.png")

try:
    res = render_excalidraw_to_png(input_file, out_path)
    print("Rendered to:", res)
except Exception as e:
    print("Render failed:", e)
    import traceback
    traceback.print_exc()
