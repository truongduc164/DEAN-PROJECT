import zipfile
from pathlib import Path

base_dir = Path(r"D:\0. Lập trình\1.DEANTRANS\outputs")
test_dir = base_dir.parent / "DeanTran" / "test_ppt_extract"
test_dir.mkdir(exist_ok=True)

orig_file = None
saved_file = base_dir / "test_save_only_no_changes.pptx"

for f in base_dir.glob("*.pptx"):
    if f.name.startswith("~$"): continue
    if "38" in f.name and "QC" in f.name and not f.name.endswith("_En.pptx") and not f.name.endswith("_Vi.pptx"):
        orig_file = f
        break

if orig_file and saved_file.exists():
    with zipfile.ZipFile(str(orig_file), 'r') as zip_ref:
        zip_ref.extractall(str(test_dir / "original"))
    with zipfile.ZipFile(str(saved_file), 'r') as zip_ref:
        zip_ref.extractall(str(test_dir / "saved"))
    print("Files extracted successfully")
else:
    print("Files not found")
