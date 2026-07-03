import sys
from pathlib import Path
from pptx import Presentation

base_dir = Path(r"D:\0. Lập trình\1.DEANTRANS\outputs")
target_file = None
for f in base_dir.glob("*.pptx"):
    if f.name.startswith("~$"): continue
    if "38" in f.name and "QC" in f.name and not f.name.endswith("_En.pptx") and not f.name.endswith("_Vi.pptx"):
        target_file = f
        break

if not target_file:
    print("File not found")
    sys.exit(1)

# skip print
prs = Presentation(str(target_file))

out_path = base_dir / "test_save_only_no_changes.pptx"
prs.save(str(out_path))
