import json
from pathlib import Path


SESSION_FILE = Path.cwd() / "pharmaguard_session.json"


def save_current_user(user):
    SESSION_FILE.write_text(
        json.dumps(user, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def get_current_user():
    if not SESSION_FILE.exists():
        return {
            "user_id": None,
            "username": "guest",
            "full_name": "Guest User",
            "role": "viewer",
        }

    try:
        return json.loads(SESSION_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {
            "user_id": None,
            "username": "guest",
            "full_name": "Guest User",
            "role": "viewer",
        }


def get_current_user_id():
    user = get_current_user()
    return user.get("user_id")


def get_current_user_label():
    user = get_current_user()
    return f"{user.get('username', 'guest')} ({user.get('role', 'viewer')})"


def clear_current_user():
    if SESSION_FILE.exists():
        SESSION_FILE.unlink()
