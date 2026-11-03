import sqlite3

connection = sqlite3.connect(
    "app.db"
)

connection.row_factory = sqlite3.Row

try:
    rows = connection.execute(
        """
        SELECT
            id,
            username,
            role,
            expires_at,
            is_active,
            created_at
        FROM users
        """
    ).fetchall()

    for row in rows:
        print(dict(row))

finally:
    connection.close()