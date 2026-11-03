import sqlite3
from datetime import datetime

DATABASE_FILE = "app.db"

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
            "admin",
            "TEMP_HASH",
            "ADMIN",
            None,
            1,
            datetime.now().isoformat()
        )
    )

    connection.commit()

    print(
        "Admin inserted"
    )

finally:

    connection.close()