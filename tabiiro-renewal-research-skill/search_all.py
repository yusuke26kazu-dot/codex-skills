import sys, io, os
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

downloads = Path.home() / "Downloads"
xlsxs = list(downloads.glob("*.xlsx"))

cmd = f"python scripts/facebook_history.py --terms ふみ --all-sheets --workbooks " + " ".join([f'"{p}"' for p in xlsxs])
print("Executing:", cmd)
os.system(cmd)
