import os
import hashlib
import hmac
from database.schema import connect, init_database, seed_basic_data


ROLES = [
    "admin",
    "pharmacy_manager",
    "pharmacist",
    "storekeeper",
    "procurement",
    "viewer",
]


def ensure_security_columns():
    init_database()
    seed_basic_data()

    conn = connect()
    cur = conn.cursor()

    cur.execute("PRAGMA table_info(users);")
    existing_columns = {row[1] for row in cur.fetchall()}

    def add_column_if_missing(column_name, column_sql):
        if column_name not in existing_columns:
            cur.execute(f"ALTER TABLE users ADD COLUMN {column_sql};")

    add_column_if_missing("password_hash", "password_hash TEXT")
    add_column_if_missing("password_salt", "password_salt TEXT")
    add_column_if_missing("last_login_at", "last_login_at TEXT")

    conn.commit()
    conn.close()

    ensure_default_admin_password()


def hash_password(password, salt=None):
    if not password:
        raise ValueError("Password cannot be empty.")

    if salt is None:
        salt = os.urandom(16).hex()

    password_bytes = password.encode("utf-8")
    salt_bytes = salt.encode("utf-8")

    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password_bytes,
        salt_bytes,
        120000
    ).hex()

    return digest, salt


def verify_password(password, stored_hash, stored_salt):
    if not stored_hash or not stored_salt:
        return False

    new_hash, _ = hash_password(password, stored_salt)
    return hmac.compare_digest(new_hash, stored_hash)


def ensure_default_admin_password():
    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT user_id, password_hash
        FROM users
        WHERE username = 'admin'
        LIMIT 1;
    """)
    row = cur.fetchone()

    if row is None:
        password_hash, salt = hash_password("admin123")
        cur.execute("""
            INSERT INTO users (
                username, full_name, role, is_active, password_hash, password_salt
            )
            VALUES (?, ?, ?, 1, ?, ?);
        """, ("admin", "System Admin", "admin", password_hash, salt))
    else:
        user_id, existing_hash = row
        if not existing_hash:
            password_hash, salt = hash_password("admin123")
            cur.execute("""
                UPDATE users
                SET password_hash = ?, password_salt = ?, role = 'admin', is_active = 1
                WHERE user_id = ?;
            """, (password_hash, salt, user_id))

    conn.commit()
    conn.close()


def create_user(username, full_name, role, password):
    ensure_security_columns()

    username = str(username).strip()
    full_name = str(full_name).strip()
    role = str(role).strip()

    if not username:
        raise ValueError("Username is required.")

    if role not in ROLES:
        raise ValueError(f"Invalid role: {role}")

    password_hash, salt = hash_password(password)

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO users (
            username, full_name, role, is_active, password_hash, password_salt
        )
        VALUES (?, ?, ?, 1, ?, ?);
    """, (username, full_name, role, password_hash, salt))

    conn.commit()
    user_id = cur.lastrowid
    conn.close()

    return user_id


def list_users():
    ensure_security_columns()

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            user_id,
            username,
            full_name,
            role,
            is_active,
            created_at,
            COALESCE(last_login_at, '') AS last_login_at
        FROM users
        ORDER BY username;
    """)

    rows = cur.fetchall()
    conn.close()
    return rows


def authenticate_user(username, password):
    ensure_security_columns()

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            user_id,
            username,
            full_name,
            role,
            is_active,
            password_hash,
            password_salt
        FROM users
        WHERE username = ?
        LIMIT 1;
    """, (username,))

    row = cur.fetchone()

    if row is None:
        conn.close()
        return None

    user_id, username, full_name, role, is_active, stored_hash, stored_salt = row

    if int(is_active) != 1:
        conn.close()
        return None

    if not verify_password(password, stored_hash, stored_salt):
        conn.close()
        return None

    cur.execute("""
        UPDATE users
        SET last_login_at = CURRENT_TIMESTAMP
        WHERE user_id = ?;
    """, (user_id,))

    conn.commit()
    conn.close()

    return {
        "user_id": user_id,
        "username": username,
        "full_name": full_name,
        "role": role,
    }


def set_user_active(user_id, is_active):
    ensure_security_columns()

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        UPDATE users
        SET is_active = ?
        WHERE user_id = ?;
    """, (1 if is_active else 0, user_id))

    conn.commit()
    conn.close()


def change_password(username, new_password):
    ensure_security_columns()

    password_hash, salt = hash_password(new_password)

    conn = connect()
    cur = conn.cursor()

    cur.execute("""
        UPDATE users
        SET password_hash = ?, password_salt = ?
        WHERE username = ?;
    """, (password_hash, salt, username))

    conn.commit()
    changed = cur.rowcount
    conn.close()

    return changed


if __name__ == "__main__":
    ensure_security_columns()
    print("Security columns ready.")
    print("Default admin login:")
    print("username: admin")
    print("password: admin123")
    print("")
    print("Users:")
    for user in list_users():
        print(user)
