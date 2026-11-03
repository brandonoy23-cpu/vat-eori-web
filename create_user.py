import base64
import getpass
import hashlib
import secrets
import sqlite3

from datetime import datetime

DATABASE_FILE = "app.db"

ITERATIONS = 600_000


def make_hash(password: str) -> str:
    salt = secrets.token_bytes(32)

    value = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        ITERATIONS,
    )

    return (
        f"pbkdf2_sha256$"
        f"{ITERATIONS}$"
        f"{base64.b64encode(salt).decode('ascii')}$"
        f"{base64.b64encode(value).decode('ascii')}"
    )


username = input(
    "Username: "
).strip()

if not username:
    raise SystemExit(
        "Κενό username."
    )

password = getpass.getpass(
    "Password: "
)

if len(password) < 10:
    raise SystemExit(
        "Τουλάχιστον 10 χαρακτήρες."
    )

expires_at = input(
    "Expiration (YYYY-MM-DD): "
).strip()

connection = sqlite3.connect(
    DATABASE_FILE
)

try:

    connection.execute(
        """
        INSERT INTO users
        (
            username,
            password_hash,
            role,
            expires_at,
            is_active,
            created_at
        )
        VALUES
        (
            ?,
            ?,
            ?,
            ?,
            ?,
            ?
        )
        """,
        (
            username,
            make_hash(password),
            "USER",
            expires_at,
            1,
            datetime.now()
            .astimezone()
            .isoformat(),
        ),
    )

    connection.commit()

    print(
        "Ο χρήστης δημιουργήθηκε."
    )

finally:

    connection.close()
