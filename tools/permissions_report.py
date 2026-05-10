from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from app.permissions import ROLE_PERMISSIONS

print("PharmaGuard Role Permissions Report")
print("=" * 50)

for role, permissions in ROLE_PERMISSIONS.items():
    print()
    print("ROLE:", role)
    print("-" * 50)

    for permission in sorted(permissions):
        print("  -", permission)

print()
print("Report finished.")