from pathlib import Path
from datetime import datetime
import shutil

PROJECT = Path.cwd()
BACKUP_ROOT = PROJECT / "code_backups"
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
TARGET = BACKUP_ROOT / f"code_backup_{timestamp}"

TARGET.mkdir(parents=True, exist_ok=True)

patterns = [
    "*.py",
    "*.bat",
    "*.db",
]

for pattern in patterns:
    for path in PROJECT.glob(pattern):
        if path.is_file():
            shutil.copy2(path, TARGET / path.name)

for folder_name in ["app", "database", "tools"]:
    src = PROJECT / folder_name
    if src.exists():
        shutil.copytree(src, TARGET / folder_name, dirs_exist_ok=True)

print("Code backup created:")
print(TARGET)