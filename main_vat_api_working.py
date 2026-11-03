from typing import Literal
from xml.etree import ElementTree

import httpx
from fastapi import FastAPI
from pydantic import BaseModel


app = FastAPI(
    title="VAT-EORI Web Application",
    version="0.2.0",
)


VIES_URL = (
    "https://ec.europa.eu/"
    "taxation_customs/vies/services/checkVatService"
)


class VatResult(BaseModel):
    country_code: str
    vat_number: str
    full_vat: str

    status: Literal[
        "VALID",
        "INVALID",
        "INPUT_ERROR",
        "TIMEOUT",
        "SERVICE_ERROR",
        "INVALID_RESPONSE",
    ]

    valid: bool | None = None
    company_name: str | None = None
    company_address: str | None = None
    request_date: str | None = None
    error: str | None = None


def normalize_country_code(country_code: str) -> str:
    normalized = country_code.strip().upper()

    # Για ελληνικό VAT το VIES χρησιμοποιεί EL.
    if normalized == "GR":
        normalized = "EL"

    if len(normalized) != 2:
        raise ValueError(
            "Το country code πρέπει να έχει δύο γράμματα."
        )

    if not normalized.isalpha():
        raise ValueError(
            "Το country code πρέπει να περιέχει μόνο γράμματα."
        )

    return normalized


def normalize_vat_number(vat_number: str) -> str:
    normalized = (
        vat_number.strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
        .replace(".", "")
    )

    if not normalized:
        raise ValueError(
            "Το VAT number είναι υποχρεωτικό."
        )

    if not normalized.isalnum():
        raise ValueError(
            "Το VAT number πρέπει να περιέχει μόνο "
            "γράμματα και αριθμούς."
        )

    if len(normalized) > 20:
        raise ValueError(
            "Το VAT number είναι μεγαλύτερο "
            "από το αναμενόμενο."
        )

    return normalized


def build_vies_request(
    country_code: str,
    vat_number: str,
) -> str:
    return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope
    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:urn="urn:ec.europa.eu:taxud:vies:services:checkVat:types">
    <soap:Body>
        <urn:checkVat>
            <urn:countryCode>{country_code}</urn:countryCode>
            <urn:vatNumber>{vat_number}</urn:vatNumber>
        </urn:checkVat>
    </soap:Body>
</soap:Envelope>"""


def find_xml_value(
    root: ElementTree.Element,
    field_name: str,
) -> str | None:
    """
    Βρίσκει ένα XML element με βάση το local name,
    ανεξάρτητα από το namespace prefix.
    """
    for element in root.iter():
        local_name = element.tag.split("}")[-1]

        if local_name == field_name:
            if element.text is None:
                return None

            value = element.text.strip()

            if not value or value == "---":
                return None

            return value

    return None


async def perform_vat_check(
    country_code: str,
    vat_number: str,
) -> VatResult:
    try:
        normalized_country = normalize_country_code(
            country_code
        )

        normalized_vat = normalize_vat_number(
            vat_number
        )

    except ValueError as exc:
        return VatResult(
            country_code=country_code.strip().upper(),
            vat_number=vat_number.strip(),
            full_vat="",
            status="INPUT_ERROR",
            error=str(exc),
        )

    full_vat = normalized_country + normalized_vat

    soap_request = build_vies_request(
        normalized_country,
        normalized_vat,
    )

    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": (
            "urn:ec.europa.eu:taxud:"
            "vies:services:checkVat:checkVat"
        ),
    }

    timeout = httpx.Timeout(
        connect=10.0,
        read=25.0,
        write=10.0,
        pool=10.0,
    )

    try:
        async with httpx.AsyncClient(
            timeout=timeout
        ) as client:
            response = await client.post(
                VIES_URL,
                content=soap_request.encode("utf-8"),
                headers=headers,
            )

    except httpx.TimeoutException:
        return VatResult(
            country_code=normalized_country,
            vat_number=normalized_vat,
            full_vat=full_vat,
            status="TIMEOUT",
            error=(
                "Το VIES δεν απάντησε μέσα "
                "στο χρονικό όριο."
            ),
        )

    except httpx.RequestError as exc:
        return VatResult(
            country_code=normalized_country,
            vat_number=normalized_vat,
            full_vat=full_vat,
            status="SERVICE_ERROR",
            error=f"Αποτυχία σύνδεσης με το VIES: {exc}",
        )

    if response.status_code != 200:
        return VatResult(
            country_code=normalized_country,
            vat_number=normalized_vat,
            full_vat=full_vat,
            status="SERVICE_ERROR",
            error=(
                "Το VIES επέστρεψε HTTP status "
                f"{response.status_code}."
            ),
        )

    try:
        root = ElementTree.fromstring(
            response.content
        )

    except ElementTree.ParseError:
        return VatResult(
            country_code=normalized_country,
            vat_number=normalized_vat,
            full_vat=full_vat,
            status="INVALID_RESPONSE",
            error=(
                "Η απάντηση του VIES δεν είναι "
                "έγκυρο XML."
            ),
        )

    fault_message = find_xml_value(
        root,
        "faultstring",
    )

    if fault_message:
        return VatResult(
            country_code=normalized_country,
            vat_number=normalized_vat,
            full_vat=full_vat,
            status="SERVICE_ERROR",
            error=f"VIES SOAP fault: {fault_message}",
        )

    valid_value = find_xml_value(
        root,
        "valid",
    )

    if valid_value is None:
        return VatResult(
            country_code=normalized_country,
            vat_number=normalized_vat,
            full_vat=full_vat,
            status="INVALID_RESPONSE",
            error=(
                "Δεν βρέθηκε το πεδίο valid "
                "στην απάντηση του VIES."
            ),
        )

    is_valid = valid_value.lower() == "true"

    company_name = find_xml_value(
        root,
        "name",
    )

    company_address = find_xml_value(
        root,
        "address",
    )

    request_date = find_xml_value(
        root,
        "requestDate",
    )

    if company_address:
        company_address = " ".join(
            company_address.split()
        )

    return VatResult(
        country_code=normalized_country,
        vat_number=normalized_vat,
        full_vat=full_vat,
        status="VALID" if is_valid else "INVALID",
        valid=is_valid,
        company_name=company_name,
        company_address=company_address,
        request_date=request_date,
    )


@app.get("/")
async def home():
    return {
        "message": "Η εφαρμογή VAT-EORI λειτουργεί",
        "version": "0.2.0",
        "documentation": "/docs",
    }


@app.get("/health")
async def health():
    return {
        "status": "ok"
    }


@app.post(
    "/api/check-vat",
    response_model=VatResult,
)
async def check_vat(
    country_code: str,
    vat_number: str,
) -> VatResult:
    return await perform_vat_check(
        country_code=country_code,
        vat_number=vat_number,
    )