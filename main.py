import asyncio
import base64
import hashlib
import os
import secrets
import sqlite3
import psycopg

from pathlib import Path
from fastapi.staticfiles import StaticFiles
from datetime import datetime, date
from io import BytesIO
from typing import Literal
from urllib.parse import quote, urlparse
from xml.etree import ElementTree

import httpx

from dotenv import load_dotenv

from fastapi import (
    BackgroundTasks,
    FastAPI,
    File,
    Form,
    HTTPException,
    Request,
    UploadFile,
)


from fastapi.responses import (
    RedirectResponse,
    StreamingResponse,
)
from fastapi.openapi.docs import get_swagger_ui_html
from fastapi.openapi.utils import get_openapi

from fastapi.templating import Jinja2Templates

from openpyxl import (
    Workbook,
    load_workbook,
)

from openpyxl.styles import (
    Alignment,
    Font,
    PatternFill,
)

from openpyxl.utils import get_column_letter

from pydantic import BaseModel

from starlette.middleware.sessions import (
    SessionMiddleware,
)

load_dotenv()
APP_USERNAME=os.getenv("APP_USERNAME","")
APP_PASSWORD_HASH=os.getenv("APP_PASSWORD_HASH","")
SESSION_SECRET=os.getenv("SESSION_SECRET","")
if not APP_USERNAME or not APP_PASSWORD_HASH or len(SESSION_SECRET)<32:
    raise RuntimeError("Λείπουν οι ρυθμίσεις login από το .env. Τρέξε: python create_credentials.py")
app=FastAPI(title="VAT & EORI Verification",version="0.8.0",docs_url=None,redoc_url=None,openapi_url=None)
app.add_middleware(SessionMiddleware,secret_key=SESSION_SECRET,session_cookie="vat_eori_session",max_age=8*60*60,same_site="lax",https_only=False)
templates = Jinja2Templates(
    directory="templates"
)
app.mount(
    "/static",
    StaticFiles(directory="static"),
    name="static",
)
LAST_RESULTS = {
    "summary": {},
    "preview_rows": [],
    "excel_file": None,
    "filename": None,
}

DATABASE_FILE = "app.db"
DATABASE_URL = os.getenv(
    "DATABASE_URL"
)
RESULTS_DIRECTORY = Path(
    "job_results"
)


RESULTS_DIRECTORY.mkdir(
    exist_ok=True
)
def test_postgresql_connection():

    if not DATABASE_URL:

        print(
            "PostgreSQL URL not configured"
        )

        return

    try:

        with psycopg.connect(
            DATABASE_URL
        ) as connection:

            with connection.cursor() as cursor:

                cursor.execute(
                    "SELECT 1"
                )

                print(
                    "PostgreSQL OK"
                )

    except Exception as error:

        print(
            "PostgreSQL Error:",
            error
        )



def initialize_postgresql():

    if not DATABASE_URL:

        return

    try:

        with psycopg.connect(
            DATABASE_URL
        ) as connection:

            with connection.cursor() as cursor:

                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS users
                    (
                        id SERIAL PRIMARY KEY,

                        username TEXT UNIQUE NOT NULL,

                        password_hash TEXT NOT NULL,

                        role TEXT NOT NULL,

                        expires_at TEXT,

                        is_active INTEGER NOT NULL DEFAULT 1,

                        created_at TEXT NOT NULL
                    )
                    """
                )

            connection.commit()

        print(
            "PostgreSQL users table ready"
        )

    except Exception as error:

        print(
            "PostgreSQL table error:",
            error
        )

def migrate_users_to_postgresql():

    if not DATABASE_URL:

        return

    sqlite_connection = sqlite3.connect(
        DATABASE_FILE
    )

    sqlite_connection.row_factory = (
        sqlite3.Row
    )

    try:

        users = sqlite_connection.execute(
            """
            SELECT
                username,
                password_hash,
                role,
                expires_at,
                is_active,
                created_at
            FROM users
            """
        ).fetchall()

    finally:

        sqlite_connection.close()

    try:

        with psycopg.connect(
            DATABASE_URL
        ) as pg_connection:

            with pg_connection.cursor() as cursor:

                for user in users:

                    cursor.execute(
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
                            %s,
                            %s,
                            %s,
                            %s,
                            %s,
                            %s
                        )
                        ON CONFLICT (username)
                        DO NOTHING
                        """,
                        (
                            user["username"],
                            user["password_hash"],
                            user["role"],
                            user["expires_at"],
                            user["is_active"],
                            user["created_at"],
                        ),
                    )

            pg_connection.commit()

        print(
            f"Migrated {len(users)} users"
        )

    except Exception as error:

        print(
            "Migration error:",
            error
        )
test_postgresql_connection()
initialize_postgresql()
migrate_users_to_postgresql()

def initialize_database():

    connection = sqlite3.connect(
        DATABASE_FILE
    )

    try:

        cursor = connection.cursor()

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS jobs
            (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                status TEXT NOT NULL,
                total_rows INTEGER NOT NULL DEFAULT 0,
                processed_rows INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT,
                error_message TEXT,
                result_filename TEXT
            )
            """
        )

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS vat_cache
            (
                vat_number TEXT PRIMARY KEY,
                status TEXT,
                company_name TEXT,
                company_address TEXT,
                request_date TEXT,
                checked_at TEXT
            )
            """
        )

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS users
            (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL,
                expires_at TEXT,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            )
            """
        )

        connection.commit()

    finally:

        connection.close()


def update_job_progress(
    job_id: int,
    processed_rows: int,
    status: str = "RUNNING",
    error_message: str | None = None,
    result_filename: str | None = None,
) -> None:
    connection = sqlite3.connect(
        DATABASE_FILE
    )

    try:
        updated_at = (
            datetime.now()
            .astimezone()
            .isoformat()
        )

        connection.execute(
            """
            UPDATE jobs
            SET
                status = ?,
                processed_rows = ?,
                updated_at = ?,
                error_message = ?,
                result_filename = ?
            WHERE id = ?
            """,
            (
                status,
                processed_rows,
                updated_at,
                error_message,
                result_filename,
                job_id,
            ),
        )

        connection.commit()

    finally:
        connection.close()


def get_job_status(
    job_id: int,
) -> dict | None:
    connection = sqlite3.connect(
        DATABASE_FILE
    )

    connection.row_factory = (
        sqlite3.Row
    )

    try:
        row = connection.execute(
            """
            SELECT
                id,
                status,
                total_rows,
                processed_rows,
                created_at,
                updated_at,
                error_message,
                result_filename
            FROM jobs
            WHERE id = ?
            """,
            (
                job_id,
            ),
        ).fetchone()

        if row is None:
            return None

        total_rows = int(
            row["total_rows"]
            or 0
        )

        processed_rows = int(
            row["processed_rows"]
            or 0
        )

        if total_rows > 0:
            percentage = round(
                (
                    processed_rows
                    / total_rows
                )
                * 100
            )
        else:
            percentage = 0

        percentage = max(
            0,
            min(
                percentage,
                100,
            ),
        )

        return {
            "id": row["id"],
            "status": row["status"],
            "total_rows": total_rows,
            "processed_rows": processed_rows,
            "percentage": percentage,
            "created_at": row["created_at"],
            "updated_at": row["updated_at"],
            "error_message": row[
                "error_message"
            ],
            "result_filename": row[
                "result_filename"
            ],
        }

    finally:
        connection.close()




def user_access_expired(
    user: dict,
) -> bool:

    expires_at = user.get(
        "expires_at"
    )

    if not expires_at:
        return False

    return (
        date.today()
        >
        date.fromisoformat(
            expires_at
        )
    )

def validate_user(
    username: str,
):

    user = get_user_by_username(
        username
    )

    if user is None:
        return None

    if (
        user["is_active"]
        != 1
    ):
        return None

    if user_access_expired(
        user
    ):
        return None

    return user

def authenticate_sqlite_user(
    username: str,
    password: str,
):
    user = validate_user(
        username
    )

    if user is None:
        return None

    if not verify_password(
        password,
        user["password_hash"]
    ):
        return None

    return user

def create_user(
    username: str,
    password_hash: str,
    role: str,
    expires_at: str | None,
):
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
                1,
                ?
            )
            """,
            (
                username,
                password_hash,
                role,
                expires_at,
                datetime.now()
                .isoformat(),
            ),
        )

        connection.commit()

    finally:

        connection.close()
def get_user_by_username(
    username: str,
):

    if not DATABASE_URL:

        return None

    try:

        with psycopg.connect(
            DATABASE_URL
        ) as connection:

            with connection.cursor() as cursor:

                cursor.execute(
                    """
                    SELECT
                        id,
                        username,
                        password_hash,
                        role,
                        expires_at,
                        is_active
                    FROM users
                    WHERE username = %s
                    """,
                    (
                        username,
                    ),
                )

                row = cursor.fetchone()

                if row is None:

                    return None

                return {
                    "id": row[0],
                    "username": row[1],
                    "password_hash": row[2],
                    "role": row[3],
                    "expires_at": row[4],
                    "is_active": row[5],
                }

    except Exception as error:

        print(
            "PostgreSQL user lookup error:",
            error
        )

        return None
def ensure_default_admin():
    connection = sqlite3.connect(
        DATABASE_FILE
    )

    connection.row_factory = sqlite3.Row

    try:

        row = connection.execute(
            """
            SELECT id
            FROM users
            LIMIT 1
            """
        ).fetchone()

        if row is not None:
            return

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
                APP_USERNAME,
                APP_PASSWORD_HASH,
                "ADMIN",
                None,
                1,
                datetime.now().isoformat(),
            )
        )

        connection.commit()

        print(
            "Default admin created"
        )

    finally:

        connection.close()

initialize_database()

ensure_default_admin()

def ensure_default_admin():
    connection = sqlite3.connect(
        DATABASE_FILE
    )

    connection.row_factory = sqlite3.Row

    try:

        row = connection.execute(
            """
            SELECT id
            FROM users
            LIMIT 1
            """
        ).fetchone()

        if row is not None:
            return

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
                APP_USERNAME,
                APP_PASSWORD_HASH,
                "ADMIN",
                None,
                1,
                datetime.now().isoformat(),
            )
        )

        connection.commit()

        print(
            "Default admin created"
        )

    finally:

        connection.close()


async def run_vat_list_job(
    job_id: int,
    country: str,
    vats: list[str],
) -> None:
    try:
        update_job_progress(
            job_id=job_id,
            processed_rows=0,
            status="RUNNING",
        )

        results = []

        eori_country = (
            normalize_eori_country_code(
                country
            )
        )

        for index, vat_number in enumerate(
            vats,
            start=1,
        ):
            vat_result = await perform_vat_check(
                country,
                vat_number,
            )

            eori_result = await perform_eori_check(
                eori_country,
                vat_number,
            )

            results.append(
                {
                    "vat": vat_number,
                    "vat_result": vat_result,
                    "eori_result": eori_result,
                }
            )

            update_job_progress(
                job_id=job_id,
                processed_rows=index,
                status="RUNNING",
            )

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "VAT EORI Results"

        headers = [
            "VAT Number",
            "Full VAT",
            "VAT Status",
            "VAT Company Name",
            "VAT Address",
            "VAT Request Date",
            "Full EORI",
            "EORI Status",
            "EORI Company Name",
            "EORI Address",
            "EORI Request Date",
            "Checked At",
            "Remarks",
        ]

        worksheet.append(headers)

        header_fill = PatternFill(
            "solid",
            fgColor="1F4E78",
        )

        for cell in worksheet[1]:
            cell.font = Font(
                color="FFFFFF",
                bold=True,
            )

            cell.fill = header_fill

            cell.alignment = Alignment(
                horizontal="center",
            )

        checked_at = (
            datetime.now()
            .astimezone()
            .strftime(
                "%Y-%m-%d %H:%M:%S %Z"
            )
        )

        green_fill = PatternFill(
            "solid",
            fgColor="C6EFCE",
        )

        red_fill = PatternFill(
            "solid",
            fgColor="FFC7CE",
        )

        orange_fill = PatternFill(
            "solid",
            fgColor="FCE4D6",
        )

        for item in results:
            vat_result = item[
                "vat_result"
            ]

            eori_result = item[
                "eori_result"
            ]

            remarks = "; ".join(
                message
                for message in [
                    vat_result.error,
                    eori_result.error,
                ]
                if message
            )

            worksheet.append(
                [
                    item["vat"],
                    vat_result.full_vat,
                    vat_result.status,
                    vat_result.company_name,
                    vat_result.company_address,
                    vat_result.request_date,
                    eori_result.full_eori,
                    eori_result.status,
                    eori_result.company_name,
                    eori_result.company_address,
                    eori_result.request_date,
                    checked_at,
                    remarks,
                ]
            )

            row_number = (
                worksheet.max_row
            )

            worksheet.cell(
                row_number,
                1,
            ).number_format = "@"

            worksheet.cell(
                row_number,
                2,
            ).number_format = "@"

            worksheet.cell(
                row_number,
                7,
            ).number_format = "@"

            for column, status in (
                (
                    3,
                    vat_result.status,
                ),
                (
                    8,
                    eori_result.status,
                ),
            ):
                if status == "VALID":
                    fill = green_fill

                elif status == "INVALID":
                    fill = red_fill

                else:
                    fill = orange_fill

                worksheet.cell(
                    row_number,
                    column,
                ).fill = fill

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        column_widths = {
            1: 18,
            2: 20,
            3: 18,
            4: 35,
            5: 45,
            6: 18,
            7: 20,
            8: 18,
            9: 35,
            10: 45,
            11: 18,
            12: 24,
            13: 55,
        }

        for column, width in (
            column_widths.items()
        ):
            worksheet.column_dimensions[
                get_column_letter(
                    column
                )
            ].width = width

        timestamp = (
            datetime.now()
            .strftime(
                "%Y%m%d_%H%M%S"
            )
        )

        filename = (
            f"vat_list_results_"
            f"{timestamp}_job_{job_id}.xlsx"
        )

        file_path = (
            RESULTS_DIRECTORY
            / filename
        )

        workbook.save(
            file_path
        )

        update_job_progress(
            job_id=job_id,
            processed_rows=len(vats),
            status="COMPLETED",
            result_filename=filename,
        )

    except Exception as exc:
        current_job = get_job_status(
            job_id
        )

        processed_rows = 0

        if current_job is not None:
            processed_rows = (
                current_job[
                    "processed_rows"
                ]
            )

        update_job_progress(
            job_id=job_id,
            processed_rows=processed_rows,
            status="FAILED",
            error_message=str(exc),
        )


TRANSLATIONS = {
    "el": {
        "language": "Γλώσσα",
        "greek": "Ελληνικά",
        "english": "Αγγλικά",
        "logout": "Αποσύνδεση",
        "change_user": "Αλλαγή χρήστη",
        "bulk_verification": "Μαζικός Έλεγχος",
        "single_verification": "Μεμονωμένος Έλεγχος",
        "login": "Σύνδεση",
        "download_results": "Λήψη Αποτελεσμάτων",
        "download_template": "Λήψη Template",
    },

    "en": {
        "language": "Language",
        "greek": "Greek",
        "english": "English",
        "logout": "Logout",
        "change_user": "Change User",
        "bulk_verification": "Bulk Verification",
        "single_verification": "Single Verification",
        "login": "Login",
        "download_results": "Download Results",
        "download_template": "Download Template",
    }
}
VIES_URL = "https://ec.europa.eu/taxation_customs/vies/services/checkVatService"
EORI_URL = "https://ec.europa.eu/taxation_customs/dds2/eos/validation/services/validation"
StatusType = Literal["VALID", "INVALID", "INPUT_ERROR", "TIMEOUT", "SERVICE_ERROR", "INVALID_RESPONSE"]

class VatResult(BaseModel):
    country_code: str
    vat_number: str
    full_vat: str
    status: StatusType
    valid: bool | None = None
    company_name: str | None = None
    company_address: str | None = None
    request_date: str | None = None
    error: str | None = None

class EoriResult(BaseModel):
    country_code: str
    eori_number: str
    full_eori: str
    status: StatusType
    valid: bool | None = None
    confirmed_eori: str | None = None
    company_name: str | None = None
    company_address: str | None = None
    request_date: str | None = None
    status_description: str | None = None
    error: str | None = None

def clean_number(value: str, field_name: str) -> str:
    cleaned = value.strip().upper().replace(" ", "").replace("-", "").replace(".", "").replace("/", "")
    if not cleaned:
        raise ValueError(f"Το {field_name} είναι υποχρεωτικό.")
    if not cleaned.isalnum():
        raise ValueError(f"Το {field_name} πρέπει να περιέχει μόνο γράμματα και αριθμούς.")
    if len(cleaned) > 30:
        raise ValueError(f"Το {field_name} είναι μεγαλύτερο από το αναμενόμενο.")
    return cleaned

def normalize_vat_country_code(code: str) -> str:
    code = code.strip().upper()
    if code == "GR": code = "EL"
    if len(code) != 2 or not code.isalpha():
        raise ValueError("Το VAT country code πρέπει να έχει δύο γράμματα.")
    return code

def normalize_eori_country_code(code: str) -> str:
    code = code.strip().upper()
    if code == "EL": code = "GR"
    if len(code) != 2 or not code.isalpha():
        raise ValueError("Το EORI country code πρέπει να έχει δύο γράμματα.")
    return code

def remove_country_prefix(number: str, prefixes: list[str]) -> str:
    for prefix in prefixes:
        if prefix and number.startswith(prefix) and len(number) > 2:
            return number[2:]
    return number

def build_vies_request(country_code: str, vat_number: str) -> bytes:
    soap="http://schemas.xmlsoap.org/soap/envelope/"
    vies="urn:ec.europa.eu:taxud:vies:services:checkVat:types"
    ElementTree.register_namespace("soap", soap); ElementTree.register_namespace("urn", vies)
    env=ElementTree.Element(f"{{{soap}}}Envelope")
    body=ElementTree.SubElement(env, f"{{{soap}}}Body")
    check=ElementTree.SubElement(body, f"{{{vies}}}checkVat")
    ElementTree.SubElement(check, f"{{{vies}}}countryCode").text=country_code
    ElementTree.SubElement(check, f"{{{vies}}}vatNumber").text=vat_number
    return ElementTree.tostring(env, encoding="utf-8", xml_declaration=True)

def build_eori_request(full_eori: str) -> bytes:
    soap="http://schemas.xmlsoap.org/soap/envelope/"; eori="http://eori.ws.eos.dds.s/"
    ElementTree.register_namespace("soapenv", soap); ElementTree.register_namespace("eori", eori)
    env=ElementTree.Element(f"{{{soap}}}Envelope")
    ElementTree.SubElement(env, f"{{{soap}}}Header")
    body=ElementTree.SubElement(env, f"{{{soap}}}Body")
    validate=ElementTree.SubElement(body, f"{{{eori}}}validateEORI")
    ElementTree.SubElement(validate, f"{{{eori}}}eori").text=full_eori
    return ElementTree.tostring(env, encoding="utf-8", xml_declaration=True)

def parse_xml(content: bytes):
    try: return ElementTree.fromstring(content)
    except ElementTree.ParseError: return None

def find_xml_value(root, field_name: str):
    for element in root.iter():
        if element.tag.split("}")[-1] == field_name:
            value=(element.text or "").strip()
            return None if not value or value == "---" else value
    return None

def normalize_text(value):
    return " ".join(value.split()) if value else None

def build_eori_address(root):
    vals=[find_xml_value(root,x) for x in ("address","street","postalCode","city","country")]
    out=[]
    for value in vals:
        value=normalize_text(value)
        if value and value not in out: out.append(value)
    return ", ".join(out) if out else None
async def perform_vat_check(country_code: str, vat_number: str) -> VatResult:
    try:
        country = normalize_vat_country_code(country_code)
        number = remove_country_prefix(
            clean_number(vat_number, "VAT number"),
            [country, "EL", "GR"]
        )
    except ValueError as exc:
        return VatResult(
            country_code=country_code.strip().upper(),
            vat_number=vat_number.strip(),
            full_vat="",
            status="INPUT_ERROR",
            error=str(exc)
        )

    full = country + number

    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": "urn:ec.europa.eu:taxud:vies:services:checkVat:checkVat",
    }

    timeout = httpx.Timeout(
        connect=10,
        read=25,
        write=10,
        pool=10,
    )

    response = None

    try:

        for attempt in range(5):

            async with httpx.AsyncClient(
                timeout=timeout,
                follow_redirects=True
            ) as client:

                response = await client.post(
                    VIES_URL,
                    content=build_vies_request(country, number),
                    headers=headers
                )

            root = parse_xml(response.content)

            if root is not None:

                fault = find_xml_value(
                    root,
                    "faultstring"
                )

                if (
                    fault
                    and
                    "MS_MAX_CONCURRENT_REQ" in fault
                ):

                    if attempt < 2:

                        await asyncio.sleep(
                            (attempt + 1) * 4
                        )

                        continue

            break

    except httpx.TimeoutException:
        return VatResult(
            country_code=country,
            vat_number=number,
            full_vat=full,
            status="TIMEOUT",
            error="Το VIES δεν απάντησε εγκαίρως."
        )

    except httpx.RequestError as exc:
        return VatResult(
            country_code=country,
            vat_number=number,
            full_vat=full,
            status="SERVICE_ERROR",
            error=f"Αποτυχία σύνδεσης VIES: {type(exc).__name__}"
        )

    if response.status_code != 200:
        return VatResult(
            country_code=country,
            vat_number=number,
            full_vat=full,
            status="SERVICE_ERROR",
            error=f"VIES HTTP {response.status_code}"
        )

    root = parse_xml(response.content)

    if root is None:
        return VatResult(
            country_code=country,
            vat_number=number,
            full_vat=full,
            status="INVALID_RESPONSE",
            error="Μη έγκυρη XML απάντηση VIES."
        )

    fault = find_xml_value(root, "faultstring")

    if fault:

        if "MS_MAX_CONCURRENT_REQ" in fault:
            return VatResult(
                country_code=country,
                vat_number=number,
                full_vat=full,
                status="SERVICE_ERROR",
                error="Η υπηρεσία VIES είναι προσωρινά απασχολημένη. Δοκιμάστε ξανά σε λίγα δευτερόλεπτα."
            )

        return VatResult(
            country_code=country,
            vat_number=number,
            full_vat=full,
            status="SERVICE_ERROR",
            error=f"VIES SOAP fault: {fault}"
        )

    valid = find_xml_value(root, "valid")

    if valid is None:
        return VatResult(
            country_code=country,
            vat_number=number,
            full_vat=full,
            status="INVALID_RESPONSE",
            error="Λείπει το πεδίο valid."
        )

    ok = valid.lower() == "true"

    return VatResult(
        country_code=country,
        vat_number=number,
        full_vat=full,
        status="VALID" if ok else "INVALID",
        valid=ok,
        company_name=normalize_text(
            find_xml_value(root, "name")
        ),
        company_address=normalize_text(
            find_xml_value(root, "address")
        ),
        request_date=find_xml_value(
            root,
            "requestDate"
        ),
    )

async def perform_eori_check(country_code: str, eori_number: str) -> EoriResult:
    try:
        country=normalize_eori_country_code(country_code)
        number=remove_country_prefix(clean_number(eori_number,"EORI number"),[country,"GR","EL"])
    except ValueError as exc:
        return EoriResult(country_code=country_code.strip().upper(),eori_number=eori_number.strip(),full_eori="",status="INPUT_ERROR",error=str(exc))
    full=country+number
    headers={"Content-Type":"text/xml; charset=utf-8","SOAPAction":"","Accept":"text/xml, application/xml","User-Agent":"VAT-EORI-Verification/0.6"}
    timeout=httpx.Timeout(connect=15,read=40,write=15,pool=15)
    response=None
    for attempt in range(1,4):
        try:
            async with httpx.AsyncClient(timeout=timeout,follow_redirects=True) as client:
                response=await client.post(EORI_URL,content=build_eori_request(full),headers=headers)
        except (httpx.TimeoutException,httpx.RequestError) as exc:
            if attempt < 3:
                await asyncio.sleep(attempt*2); continue
            status="TIMEOUT" if isinstance(exc,httpx.TimeoutException) else "SERVICE_ERROR"
            return EoriResult(country_code=country,eori_number=number,full_eori=full,status=status,error=f"Αποτυχία EORI μετά από 3 προσπάθειες: {type(exc).__name__}")
        if response.status_code in {429,500,502,503,504} and attempt < 3:
            await asyncio.sleep(attempt*2); continue
        break
    if response is None or response.status_code != 200:
        code=response.status_code if response else "χωρίς απάντηση"
        return EoriResult(country_code=country,eori_number=number,full_eori=full,status="SERVICE_ERROR",error=f"EORI HTTP {code} μετά από προσπάθειες.")
    root=parse_xml(response.content)
    if root is None:
        return EoriResult(country_code=country,eori_number=number,full_eori=full,status="INVALID_RESPONSE",error="Μη έγκυρη XML απάντηση EORI.")
    fault=find_xml_value(root,"faultstring")
    if fault:
        return EoriResult(country_code=country,eori_number=number,full_eori=full,status="SERVICE_ERROR",error=f"EORI SOAP fault: {fault}")
    descr=find_xml_value(root,"statusDescr")
    if descr is None:
        return EoriResult(country_code=country,eori_number=number,full_eori=full,status="INVALID_RESPONSE",error="Λείπει το statusDescr.")
    ok=descr.strip().upper()=="VALID"
    return EoriResult(country_code=country,eori_number=number,full_eori=full,status="VALID" if ok else "INVALID",valid=ok,confirmed_eori=find_xml_value(root,"eori"),company_name=normalize_text(find_xml_value(root,"name")),company_address=build_eori_address(root),request_date=find_xml_value(root,"requestDate"),status_description=descr)

def normalize_excel_header(value):
    return "" if value is None else " ".join(str(value).strip().lower().split())

def excel_cell_to_text(value):
    if value is None: return ""
    if isinstance(value,float) and value.is_integer(): return str(int(value))
    return str(value).strip()

def extract_rows(file_content: bytes):
    wb=load_workbook(BytesIO(file_content),read_only=True,data_only=True)
    ws=wb.active
    headers=next(ws.iter_rows(min_row=1,max_row=1,values_only=True),None)
    if not headers:
        wb.close(); raise ValueError("Δεν βρέθηκε γραμμή επικεφαλίδων.")
    mapping={normalize_excel_header(v):i for i,v in enumerate(headers) if normalize_excel_header(v)}
    country=next((mapping[x] for x in ("country code","countrycode","country") if x in mapping),None)
    vat=next((mapping[x] for x in ("vat number","vatnumber","vat no","vat no.","vat") if x in mapping),None)
    missing=[]
    if country is None: missing.append("Country Code")
    if vat is None: missing.append("VAT Number")
    if missing:
        wb.close(); raise ValueError("Λείπουν οι στήλες: "+", ".join(missing))
    rows=[]
    for row_no,values in enumerate(ws.iter_rows(min_row=2,values_only=True),start=2):
        cv=values[country] if country < len(values) else None
        vv=values[vat] if vat < len(values) else None
        if cv is None and vv is None: continue
        rows.append((row_no,excel_cell_to_text(cv),excel_cell_to_text(vv)))
    wb.close()
    if not rows: raise ValueError("Δεν βρέθηκαν εγγραφές.")
    return rows

ITERATIONS = 600_000


def make_hash(
    password: str,
) -> str:

    salt = secrets.token_bytes(
        32
    )

    value = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode(
            "utf-8"
        ),
        salt,
        ITERATIONS,
    )

    return (
        f"pbkdf2_sha256$"
        f"{ITERATIONS}$"
        f"{base64.b64encode(salt).decode('ascii')}$"
        f"{base64.b64encode(value).decode('ascii')}"
    )

def verify_password(password: str, stored: str) -> bool:
    try:
        alg,it,salt_text,hash_text=stored.split("$",3)
        if alg!="pbkdf2_sha256": return False
        salt=base64.b64decode(salt_text); expected=base64.b64decode(hash_text)
        actual=hashlib.pbkdf2_hmac("sha256",password.encode(),salt,int(it))
        return secrets.compare_digest(actual,expected)
    except Exception: return False

def is_authenticated(
    request: Request,
) -> bool:

    return (
        request.session.get(
            "authenticated"
        ) is True
    )
def redirect_login(request: Request):
    target=request.url.path + (("?"+request.url.query) if request.url.query else "")
    return RedirectResponse("/login?next_url="+quote(target,safe="/?=&"),status_code=303)

def require_login(request: Request):
    if not is_authenticated(request): raise HTTPException(status_code=401,detail="Authentication required")
def require_admin(
    request: Request,
):
    print(
        "SESSION:",
        dict(request.session)
    )

    require_login(
        request
    )

    if (
        request.session.get(
            "role"
        )
        != "ADMIN"
    ):
        raise HTTPException(
            status_code=403,
            detail="Admin only"
        )
def get_texts(
    request: Request,
):
    lang = request.session.get(
        "lang",
        "el"
    )

    return TRANSLATIONS.get(
        lang,
        TRANSLATIONS["el"]
    )
@app.get("/login")
async def login_page(request: Request,next_url: str="/excel"):
    if is_authenticated(request): return RedirectResponse("/excel",status_code=303)
    return templates.TemplateResponse(request=request,name="login.html",context={"error":None,"next_url":next_url if next_url.startswith("/") and not next_url.startswith("//") else "/excel"})

@app.post("/login")
async def login_submit(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    next_url: str = Form("/excel"),
):
    user = get_user_by_username(
        username.strip()
    )
    print(
    "LOGIN USER:",
    user
)
    if user is None:
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "error":
                "Λανθασμένο username ή password.",
                "next_url":
                next_url,
            },
            status_code=401,
        )

    if not user["is_active"]:
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "error":
                "Ο λογαριασμός είναι απενεργοποιημένος.",
                "next_url":
                next_url,
            },
            status_code=401,
        )
    if user_access_expired(
        user
    ):
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "error":
                "Η πρόσβασή σας έχει λήξει.",
                "next_url":
                next_url,
            },
            status_code=401,
        )

    if not verify_password(
        password,
        user["password_hash"],
    ):
        return templates.TemplateResponse(
            request=request,
            name="login.html",
            context={
                "error":
                "Λανθασμένο username ή password.",
                "next_url":
                next_url,
            },
            status_code=401,
        )
    print(
    "LOGIN SUCCESS:",
    user["username"],
    user["role"]
)
    request.session.clear()

    request.session[
        "authenticated"
    ] = True

    request.session[
        "username"
    ] = user["username"]

    request.session[
        "role"
    ] = user["role"]

    print(
    "AFTER LOGIN SESSION:",
    dict(request.session)
)
    print(
    	"LOGIN:",
    	user["username"],
    	user["role"]
)
    return RedirectResponse(
        next_url,
        status_code=303,
    )
@app.get("/debug-users")
async def debug_users():

    connection = sqlite3.connect(
        DATABASE_FILE
    )

    try:

        rows = connection.execute(
            """
            SELECT
                username,
                role,
                expires_at,
                is_active
            FROM users
            """
        ).fetchall()

        return rows

    finally:

        connection.close()

@app.get("/logout")
async def logout(request: Request):
    request.session.clear(); response=RedirectResponse("/login",status_code=303); response.delete_cookie("vat_eori_session"); return response

@app.get("/change-user")
@app.get("/set-language/{lang}")
async def set_language(
    request: Request,
    lang: str,
):
    if lang not in ["el", "en"]:
        lang = "el"

    request.session["lang"] = lang

    referer = (
        request.headers.get(
            "referer"
        )
        or "/excel"
    )

    return RedirectResponse(
        url=referer,
        status_code=303,
    )
async def change_user(request: Request):
    request.session.clear(); response=RedirectResponse("/login?next_url=/excel",status_code=303); response.delete_cookie("vat_eori_session"); return response

@app.get("/docs",include_in_schema=False)
async def docs(request: Request):
    if not is_authenticated(request): return redirect_login(request)
    return get_swagger_ui_html(openapi_url="/openapi.json",title="VAT EORI API")

@app.get("/openapi.json",include_in_schema=False)
async def openapi(request: Request):
    require_login(request); return get_openapi(title=app.title,version=app.version,routes=app.routes)

@app.get("/")
async def home(request: Request):

    if not is_authenticated(request):
        return redirect_login(request)

    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "vat_result": None,
            "eori_result": None,
            "vat_country_code": "EL",
            "eori_country_code": "GR",
            "number": "",
            "t": get_texts(request),
        },
    )
@app.post("/check-all")
async def check_all_form(request: Request,vat_country_code: str=Form(...),eori_country_code: str=Form(...),number: str=Form(...)):
    if not is_authenticated(request): return redirect_login(request)
    vat,eori=await asyncio.gather(perform_vat_check(vat_country_code,number),perform_eori_check(eori_country_code,number))
    return templates.TemplateResponse(request=request,name="index.html",context={"vat_result":vat,"eori_result":eori,"vat_country_code":vat_country_code,"eori_country_code":eori_country_code,"number":number})

@app.get("/excel")
async def excel_page(request: Request):

    if not is_authenticated(request):
        return redirect_login(request)

    return templates.TemplateResponse(
        request=request,
        name="upload.html",
        context={
            "error": None,

            "t": get_texts(request),

            "current_user": request.session.get(
                "username"
            ),

            "current_role": request.session.get(
                "role"
            ),
        },
    )

@app.post("/process-excel")
async def process_excel(request: Request,excel_file: UploadFile=File(...)):
    require_login(request)
    start_time = datetime.now()
    filename=excel_file.filename or ""
    if not filename.lower().endswith(".xlsx"):
    
        return StreamingResponse(BytesIO("Επιτρέπονται μόνο .xlsx".encode("utf-8")),media_type="text/plain",status_code=400)
    content=await excel_file.read()
    if not content or len(content)>5*1024*1024:
        return StreamingResponse(BytesIO("Κενό ή πολύ μεγάλο αρχείο".encode("utf-8")),media_type="text/plain",status_code=400)
    try:
        input_rows=extract_rows(content)
    except Exception as exc:
        return StreamingResponse(BytesIO(str(exc).encode("utf-8")),media_type="text/plain",status_code=400)
    if len(input_rows)>200:
        return StreamingResponse(BytesIO("Μέγιστο όριο: 200 εγγραφές ανά αρχείο.".encode("utf-8")),media_type="text/plain",status_code=400)

    semaphore=asyncio.Semaphore(2)
    async def process_one(row_no,country,number):
        async with semaphore:
            try:
                vat_country=normalize_vat_country_code(country)
                eori_country=normalize_eori_country_code(country)
                vat,eori=await asyncio.gather(perform_vat_check(vat_country,number),perform_eori_check(eori_country,number))
                return row_no,country,number,vat,eori
            except Exception as exc:
                vat=VatResult(country_code=country,vat_number=number,full_vat="",status="INPUT_ERROR",error=str(exc))
                eori=EoriResult(country_code=country,eori_number=number,full_eori="",status="INPUT_ERROR",error=str(exc))
                return row_no,country,number,vat,eori

    results = await asyncio.gather(*(process_one(*row) for row in input_rows))
    total_records = len(results)

    vat_valid = sum(
        1
        for _, _, _, vat, _
        in results
        if vat.status == "VALID"
    )

    vat_invalid = sum(
        1
        for _, _, _, vat, _
        in results
        if vat.status == "INVALID"
    )

    eori_valid = sum(
        1
        for _, _, _, _, eori
        in results
        if eori.status == "VALID"
    )

    eori_invalid = sum(
        1
        for _, _, _, _, eori
        in results
        if eori.status == "INVALID"
    )

    total_errors = sum(
        1
        for _, _, _, vat, eori
        in results
        if (
            vat.status
            not in ["VALID", "INVALID"]
        )
        or (
            eori.status
            not in ["VALID", "INVALID"]
        )
    )
    LAST_RESULTS["summary"] = {
        "total_records": total_records,
        "vat_valid": vat_valid,
        "vat_invalid": vat_invalid,
        "eori_valid": eori_valid,
        "eori_invalid": eori_invalid,
        "total_errors": total_errors,
    }
    preview_rows = []
    for (
        row_no,
        country,
        number,
        vat,
        eori,
    ) in results:

        preview_rows.append(
            {
                "full_vat": vat.full_vat,
                "vat_status": vat.status,
                "eori_status": eori.status,
            }
        )
    LAST_RESULTS["preview_rows"] = preview_rows

    wb=Workbook(); ws=wb.active; ws.title="VAT EORI Results"
    headers=["Source Row","Country Code","VAT Number","Full VAT","VAT Status","VAT Company Name","VAT Address","VAT Request Date","EORI Country Code","Full EORI","EORI Status","Confirmed EORI","EORI Company Name","EORI Address","EORI Request Date","Checked At","Remarks"]
    ws.append(headers)
    header_fill=PatternFill("solid",fgColor="1F4E78")
    for cell in ws[1]:
        cell.font=Font(color="FFFFFF",bold=True); cell.fill=header_fill; cell.alignment=Alignment(horizontal="center")
    checked=datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")
    green=PatternFill("solid",fgColor="C6EFCE"); red=PatternFill("solid",fgColor="FFC7CE"); orange=PatternFill("solid",fgColor="FCE4D6")
    for row_no,country,number,vat,eori in results:
        remarks="; ".join(x for x in [vat.error,eori.error] if x)
        ws.append([row_no,country,number,vat.full_vat,vat.status,vat.company_name,vat.company_address,vat.request_date,eori.country_code,eori.full_eori,eori.status,eori.confirmed_eori,eori.company_name,eori.company_address,eori.request_date,checked,remarks])
        r=ws.max_row
        ws.cell(r,3).number_format="@"
        ws.cell(r,4).number_format="@"
        ws.cell(r,10).number_format="@"
        for col,status in ((5,vat.status),(11,eori.status)):
            ws.cell(r,col).fill=green if status=="VALID" else red if status=="INVALID" else orange
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    widths={1:12,2:14,3:18,4:20,5:18,6:35,7:45,8:18,9:18,10:20,11:18,12:20,13:35,14:45,15:18,16:24,17:55}
    for col,width in widths.items(): ws.column_dimensions[get_column_letter(col)].width=width
    output = BytesIO()

    wb.save(output)

    output.seek(0)

    stamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    out_name = (
        f"vat_eori_results_{stamp}.xlsx"
    )

    return StreamingResponse(
        output,
        media_type=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition":
            f'attachment; filename="{out_name}"'
        },
    )
@app.get("/results")
async def results_page(
    request: Request,
):
    require_login(request)

    summary = LAST_RESULTS.get(
        "summary",
        {}
    )

    preview_rows = LAST_RESULTS.get(
        "preview_rows",
        []
    )

    return templates.TemplateResponse(
        request=request,
        name="results.html",
        context={"t": get_texts(request),
            "total_records": summary.get(
                "total_records",
                0
            ),
            "vat_valid": summary.get(
                "vat_valid",
                0
            ),
            "vat_invalid": summary.get(
                "vat_invalid",
                0
            ),
            "eori_valid": summary.get(
                "eori_valid",
                0
            ),
            "eori_invalid": summary.get(
                "eori_invalid",
                0
            ),
            "total_errors": summary.get(
                "total_errors",
                0
            ),
            "preview_rows": preview_rows,
        },
    )
@app.get("/download-results")
async def download_results(
    request: Request,
):
    require_login(request)

    excel_bytes = LAST_RESULTS.get(
        "excel_file"
    )

    if excel_bytes is None:

        return RedirectResponse(
            url="/results",
            status_code=303,
        )

    filename = LAST_RESULTS.get(
        "filename",
        "vat_eori_results.xlsx",
    )

    return StreamingResponse(
        BytesIO(excel_bytes),

        media_type=
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        headers={
            "Content-Disposition":
            f'attachment; filename="{filename}"'
        },
    )
@app.post("/start-vat-list-job")
async def start_vat_list_job(
    request: Request,
    background_tasks: BackgroundTasks,
):
    require_login(request)

    data = await request.json()

    country = (
        str(
            data.get(
                "country",
                ""
            )
        )
        .strip()
        .upper()
    )

    raw_vats = data.get(
        "vats",
        []
    )

    if not isinstance(
        raw_vats,
        list,
    ):
        raise HTTPException(
            status_code=400,
            detail=(
                "Η λίστα VAT δεν είναι "
                "έγκυρη."
            ),
        )

    vats = []

    for value in raw_vats:
        cleaned = str(
            value
        ).strip()

        if cleaned:
            vats.append(
                cleaned
            )

    if not vats:
        raise HTTPException(
            status_code=400,
            detail=(
                "Δεν βρέθηκαν VAT "
                "numbers."
            ),
        )

    if len(vats) > 100:
        raise HTTPException(
            status_code=400,
            detail=(
                "Μέγιστο 100 VAT "
                "ανά εκτέλεση."
            ),
        )

    try:
        normalize_vat_country_code(
            country
        )

    except ValueError as exc:
        raise HTTPException(
            status_code=400,
            detail=str(exc),
        ) from exc

    job_id = create_job(
        total_rows=len(vats)
    )

    background_tasks.add_task(
        run_vat_list_job,
        job_id,
        country,
        vats,
    )

    return {
        "job_id": job_id,
        "status": "PENDING",
        "total_rows": len(vats),
        "status_url": (
            f"/job-status/{job_id}"
        ),
        "download_url": (
            f"/download-job/{job_id}"
        ),
    }
@app.post("/process-vat-list")
async def process_vat_list(
    request: Request,
):
    require_login(request)

    data = await request.json()

    country = (
        data.get(
            "country",
            ""
        )
        .strip()
        .upper()
    )

    vats = data.get(
        "vats",
        []
    )

    results = []

    for vat_number in vats:

        vat_result = await perform_vat_check(
            country,
            vat_number,
        )

        eori_country = (
            normalize_eori_country_code(
                country
            )
        )

        eori_result = await perform_eori_check(
            eori_country,
            vat_number,
        )

        results.append(
    {
        "vat": vat_number,

        "vat_result":
            vat_result,

        "eori_result":
            eori_result,
    }
)

    

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = (
        "VAT EORI Results"
    )

    headers = [
    "VAT Number",
    "Full VAT",
    "VAT Status",
    "VAT Company Name",
    "VAT Address",
    "VAT Request Date",

    "Full EORI",
    "EORI Status",
    "EORI Company Name",
    "EORI Address",
    "EORI Request Date",
]

    worksheet.append(headers)
    for item in results:

        vat_result = item["vat_result"]

        eori_result = item["eori_result"]

        worksheet.append(
            [
                item["vat"],

                vat_result.full_vat,
                vat_result.status,
                vat_result.company_name,
                vat_result.company_address,
                vat_result.request_date,

                eori_result.full_eori,
                eori_result.status,
                eori_result.company_name,
                eori_result.company_address,
                eori_result.request_date,
            ]
        )

    output = BytesIO()

    workbook.save(output)

    output.seek(0)

    file_name = (
        "vat_list_results.xlsx"
    )

    return StreamingResponse(
        output,

        media_type=
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        headers={
            "Content-Disposition":
            f'attachment; filename="{file_name}"'
        },
    )


@app.get("/job-status/{job_id}")
async def job_status(
    request: Request,
    job_id: int,
):
    require_login(request)

    job = get_job_status(
        job_id
    )

    if job is None:
        raise HTTPException(
            status_code=404,
            detail="Το job δεν βρέθηκε.",
        )

    return job 
@app.get("/download-job/{job_id}")
async def download_job(
    request: Request,
    job_id: int,
):
    require_login(request)

    job = get_job_status(
        job_id
    )

    if job is None:
        raise HTTPException(
            status_code=404,
            detail="Το job δεν βρέθηκε.",
        )

    if job["status"] != "COMPLETED":
        raise HTTPException(
            status_code=409,
            detail=(
                "Το job δεν έχει "
                "ολοκληρωθεί ακόμη."
            ),
        )

    filename = job[
        "result_filename"
    ]

    if not filename:
        raise HTTPException(
            status_code=404,
            detail=(
                "Δεν βρέθηκε αρχείο "
                "αποτελεσμάτων."
            ),
        )

    safe_filename = Path(
        filename
    ).name

    file_path = (
        RESULTS_DIRECTORY
        / safe_filename
    )

    if not file_path.exists():
        raise HTTPException(
            status_code=404,
            detail=(
                "Το αρχείο αποτελεσμάτων "
                "δεν υπάρχει."
            ),
        )

    return StreamingResponse(
        file_path.open(
            "rb"
        ),
        media_type=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition":
            (
                "attachment; "
                f'filename="{safe_filename}"'
            )
        },
    ) 
@app.get("/admin/users")
async def admin_users(
    request: Request,
):
    require_admin(
        request
    )

    connection = sqlite3.connect(
        DATABASE_FILE
    )

    connection.row_factory = (
        sqlite3.Row
    )

    try:

        users = connection.execute(
            """
            SELECT
                id,
                username,
                role,
                expires_at,
                is_active
            FROM users
            ORDER BY username
            """
        ).fetchall()

        total_users = len(
            users
        )

        active_users = len(
            [
                u
                for u in users
                if u["is_active"] == 1
            ]
        )

        disabled_users = len(
            [
                u
                for u in users
                if u["is_active"] == 0
            ]
        )

        admin_users = len(
            [
                u
                for u in users
                if u["role"] == "ADMIN"
            ]
        )

        return templates.TemplateResponse(
            request=request,
            name="admin_users.html",
            context={
                "users": users,
          	"current_user": request.session.get(
        "username"
    ),
		"current_user": request.session.get(
    "username"
),

		"current_role": request.session.get(
    "role"
),

                "error": request.query_params.get(
                    "error"
                ),

                "success": request.query_params.get(
                    "success"
                ),

                "total_users": total_users,

                "active_users": active_users,

                "disabled_users": disabled_users,

                "admin_users": admin_users,
            },
        )

    finally:

        connection.close()
@app.post("/admin/create-user")
async def admin_create_user(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    role: str = Form(...),
    expires_at: str = Form(""),
):
    require_admin(
        request
    )

    connection = sqlite3.connect(
        DATABASE_FILE
    )

    try:

        existing_user = connection.execute(
            """
            SELECT id
            FROM users
            WHERE username = ?
            """,
            (
                username.strip(),
            ),
        ).fetchone()

        if existing_user:

            return RedirectResponse(
                "/admin/users?error=username_exists",
                status_code=303,
            )

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
                username.strip(),
                make_hash(
                    password
                ),
                role,
                expires_at or None,
                1,
                datetime.now()
                .isoformat(),
            ),
        )

        connection.commit()

    finally:

        connection.close()

    return RedirectResponse(
    "/admin/users?success=user_created",
    status_code=303,
)

@app.post(
    "/admin/disable-user/{user_id}"
)
async def disable_user(
    request: Request,
    user_id: int,
):
    require_admin(
        request
    )

    connection = sqlite3.connect(
        DATABASE_FILE
    )

    try:

        connection.execute(
            """
            UPDATE users
            SET is_active = 0
            WHERE id = ?
            """,
            (
                user_id,
            ),
        )

        connection.commit()

    finally:

        connection.close()

    return RedirectResponse(
        "/admin/users",
        status_code=303,
    )
@app.post(
    "/admin/enable-user/{user_id}"
)
async def enable_user(
    request: Request,
    user_id: int,
):
    require_admin(
        request
    )

    connection = sqlite3.connect(
        DATABASE_FILE
    )

    try:

        connection.execute(
            """
            UPDATE users
            SET is_active = 1
            WHERE id = ?
            """,
            (
                user_id,
            ),
        )

        connection.commit()

    finally:

        connection.close()

    return RedirectResponse(
        "/admin/users",
        status_code=303,
    )
@app.post(
    "/admin/update-expiration/{user_id}"
)
async def update_expiration(
    request: Request,
    user_id: int,
    expires_at: str = Form(...),
):
    require_admin(
        request
    )

    connection = sqlite3.connect(
        DATABASE_FILE
    )

    try:

        connection.execute(
            """
            UPDATE users
            SET expires_at = ?
            WHERE id = ?
            """,
            (
                expires_at,
                user_id,
            ),
        )

        connection.commit()

    finally:

        connection.close()

    return RedirectResponse(
        "/admin/users",
        status_code=303,
    )
@app.post(
    "/admin/reset-password/{user_id}"
)
async def reset_password(
    request: Request,
    user_id: int,
    new_password: str = Form(...),
):
    require_admin(
        request
    )

    if len(
        new_password
    ) < 10:
        return RedirectResponse(
            "/admin/users",
            status_code=303,
        )

    connection = sqlite3.connect(
        DATABASE_FILE
    )

    try:

        connection.execute(
            """
            UPDATE users
            SET password_hash = ?
            WHERE id = ?
            """,
            (
                make_hash(
                    new_password
                ),
                user_id,
            ),
        )

        connection.commit()

    finally:

        connection.close()

    return RedirectResponse(
        "/admin/users",
        status_code=303,
    )
@app.post(
    "/admin/delete-user/{user_id}"
)
async def delete_user(
    request: Request,
    user_id: int,
):
    require_admin(
        request
    )

    connection = sqlite3.connect(
        DATABASE_FILE
    )

    connection.row_factory = sqlite3.Row

    try:

        row = connection.execute(
            """
            SELECT username
            FROM users
            WHERE id = ?
            """,
            (
                user_id,
            ),
        ).fetchone()

        if (
            row
            and row["username"] == "admin"
        ):
            return RedirectResponse(
                "/admin/users",
                status_code=303,
            )

        connection.execute(
            """
            DELETE FROM users
            WHERE id = ?
            """,
            (
                user_id,
            ),
        )

        connection.commit()

    finally:

        connection.close()

    return RedirectResponse(
        "/admin/users",
        status_code=303,
    )
@app.get("/health")
async def health(request: Request):
    require_login(request)
    return {"status":"ok","version":"0.8.0"}

@app.post("/api/check-vat",response_model=VatResult)
async def vat_api(request: Request,country_code: str,vat_number: str):
    require_login(request)
    return await perform_vat_check(country_code,vat_number)

@app.post("/api/check-eori",response_model=EoriResult)
async def eori_api(request: Request,country_code: str,eori_number: str):
    require_login(request)
    return await perform_eori_check(country_code,eori_number)
