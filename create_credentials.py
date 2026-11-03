import base64
import getpass
import hashlib
import secrets
from pathlib import Path

ITERATIONS = 600_000

def make_hash(password: str) -> str:
    salt = secrets.token_bytes(32)
    value = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, ITERATIONS)
    return f"pbkdf2_sha256${ITERATIONS}${base64.b64encode(salt).decode('ascii')}${base64.b64encode(value).decode('ascii')}"

username = input('Username: ').strip()
if not username:
    raise SystemExit('Το username είναι κενό.')
p1 = getpass.getpass('Password: ')
p2 = getpass.getpass('Επανάληψη password: ')
if p1 != p2:
    raise SystemExit('Τα passwords δεν είναι ίδια.')
if len(p1) < 10:
    raise SystemExit('Χρειάζονται τουλάχιστον 10 χαρακτήρες.')
Path('.env').write_text(
    f"APP_USERNAME={username}\nAPP_PASSWORD_HASH={make_hash(p1)}\nSESSION_SECRET={secrets.token_urlsafe(64)}\n",
    encoding='utf-8',
)
print('Το .env δημιουργήθηκε επιτυχώς.')
