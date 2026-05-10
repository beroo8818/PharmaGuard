from pathlib import Path

root = Path.cwd()
count = 0

for path in root.rglob("*.py"):
    try:
        text = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        text = path.read_text(encoding="utf-8", errors="replace")

    cleaned = text.replace("\ufeff", "")

    if cleaned != text:
        count += 1
        print("Cleaned:", path)

    path.write_text(cleaned, encoding="utf-8")

print("Done. Files cleaned:", count)
