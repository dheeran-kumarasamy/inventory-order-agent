import json
import os
import re
from datetime import datetime

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, SpreadsheetNotFound

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


MONTH_PATTERNS = ["%B-%Y", "%b-%Y"]


class ConfigError(ValueError):
    pass


def _looks_like_service_account_mapping(data) -> bool:
    if not isinstance(data, dict):
        return False
    required = {"type", "client_email", "private_key", "token_uri"}
    return required.issubset(set(data.keys()))


def _load_service_account_info() -> dict:
    # Support either raw JSON string or structured secret/env representations.
    if "GOOGLE_SERVICE_ACCOUNT_JSON" in st.secrets:
        raw = st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"]
        if isinstance(raw, dict):
            return dict(raw)
        try:
            return json.loads(str(raw))
        except Exception as err:
            raise ConfigError(
                "Invalid GOOGLE_SERVICE_ACCOUNT_JSON. Expected valid JSON service-account content."
            ) from err

    if "google_service_account" in st.secrets:
        sa = st.secrets["google_service_account"]
        if isinstance(sa, dict):
            return dict(sa)
        try:
            return dict(sa)
        except Exception as err:
            raise ConfigError(
                "Invalid [google_service_account] in Streamlit secrets. Expected a key/value object."
            ) from err

    # Also accept top-level secrets.toml service account fields.
    if _looks_like_service_account_mapping(dict(st.secrets)):
        return dict(st.secrets)

    env_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if env_json:
        try:
            return json.loads(env_json)
        except Exception as err:
            raise ConfigError(
                "Invalid GOOGLE_SERVICE_ACCOUNT_JSON environment variable. Expected valid JSON."
            ) from err

    env_json_alt = os.getenv("GOOGLE_SERVICE_ACCOUNT", "").strip()
    if env_json_alt:
        try:
            return json.loads(env_json_alt)
        except Exception as err:
            raise ConfigError(
                "Invalid GOOGLE_SERVICE_ACCOUNT environment variable. Expected valid JSON."
            ) from err

    raise ConfigError(
        "Missing Google credentials. Add GOOGLE_SERVICE_ACCOUNT_JSON to Streamlit secrets "
        "(or add [google_service_account] fields, or top-level service account fields), "
        "or set GOOGLE_SERVICE_ACCOUNT_JSON env var."
    )


def get_gspread_client():
    creds_dict = _load_service_account_info()
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def _service_account_email() -> str:
    try:
        creds_dict = _load_service_account_info()
        return str(creds_dict.get("client_email", "")).strip()
    except ConfigError:
        return ""


def _raise_sheet_access_error(url: str, err: Exception) -> None:
    email = _service_account_email()
    share_hint = (
        f"Share the Google Sheet with this service account email: {email}."
        if email
        else "Share the Google Sheet with the configured service account email in Streamlit secrets."
    )
    raise ValueError(
        "Unable to access the Google Sheet. "
        "This is usually a permissions issue or an invalid URL. "
        f"{share_hint} "
        f"URL: {url}. "
        f"Original error: {type(err).__name__}: {str(err) or repr(err)}"
    ) from err


def _open_spreadsheet(url: str):
    try:
        client = get_gspread_client()
        return client.open_by_key(sheet_id_from_url(url))
    except ConfigError:
        raise
    except (PermissionError, SpreadsheetNotFound, APIError, ValueError) as err:
        _raise_sheet_access_error(url, err)


def sheet_id_from_url(url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        raise ValueError(f"Could not extract Sheet ID from URL: {url}")
    return match.group(1)


def _normalize_key(name: str) -> str:
    return str(name).strip().lower().replace(" ", "")


def _map_column(df: pd.DataFrame, canonical: str, aliases: list[str]) -> pd.DataFrame:
    lookup = {_normalize_key(c): c for c in df.columns}
    for alias in aliases:
        key = _normalize_key(alias)
        if key in lookup:
            df = df.rename(columns={lookup[key]: canonical})
            return df
    return df


def _parse_month_tab_title(title: str):
    value = title.strip()
    for pattern in MONTH_PATTERNS:
        try:
            return datetime.strptime(value, pattern)
        except ValueError:
            continue
    return None


def discover_master_sheet_structure(url: str) -> dict:
    spreadsheet = _open_spreadsheet(url)

    all_tabs = [ws.title for ws in spreadsheet.worksheets()]
    has_software_stock = any(tab.strip().lower() == "software stock" for tab in all_tabs)

    monthly_tabs = []
    for title in all_tabs:
        month_date = _parse_month_tab_title(title)
        if month_date:
            monthly_tabs.append(
                {
                    "Sheet": title,
                    "Month": month_date.strftime("%Y-%m"),
                }
            )

    monthly_tabs = sorted(monthly_tabs, key=lambda r: r["Month"])

    return {
        "has_software_stock": has_software_stock,
        "monthly_tabs": monthly_tabs,
        "all_tabs": all_tabs,
    }


def load_stock_from_sheet(url: str) -> pd.DataFrame:
    spreadsheet = _open_spreadsheet(url)

    worksheet = None
    for ws in spreadsheet.worksheets():
        if ws.title.strip().lower() == "software stock":
            worksheet = ws
            break

    if worksheet is None:
        raise ValueError("Sheet tab 'software stock' not found.")

    data = worksheet.get_all_records()
    if not data:
        raise ValueError("'software stock' sheet is empty.")

    df = pd.DataFrame(data)
    df.columns = [str(c).strip() for c in df.columns]

    df = _map_column(df, "ProductName", ["ProductName", "Product Name", "Item", "ItemName", "Name"])
    df = _map_column(df, "StockLevel", ["StockLevel", "Stocklevel", "Stock Level"])
    df = _map_column(df, "ReorderLevel", ["ReorderLevel", "Reorderlevel", "Reorder Level"])

    required = ["ProductName", "StockLevel", "ReorderLevel"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"Missing required columns in 'software stock': {missing}. Found columns: {list(df.columns)}"
        )

    df["ProductName"] = df["ProductName"].astype(str).str.strip()
    df["StockLevel"] = pd.to_numeric(df["StockLevel"], errors="coerce").fillna(0)
    df["ReorderLevel"] = pd.to_numeric(df["ReorderLevel"], errors="coerce").fillna(0)

    return df


def load_sales_from_sheet(url: str) -> pd.DataFrame:
    spreadsheet = _open_spreadsheet(url)

    sales_frames = []
    ignored = []

    for ws in spreadsheet.worksheets():
        month_date = _parse_month_tab_title(ws.title)
        if not month_date:
            ignored.append(ws.title)
            continue

        records = ws.get_all_records()
        if not records:
            continue

        df = pd.DataFrame(records)
        df.columns = [str(c).strip() for c in df.columns]

        df = _map_column(df, "ProductName", ["ProductName", "Product Name", "Item", "ItemName", "Name"])
        df = _map_column(df, "Total", ["Total", "Sales", "Qty", "Quantity", "Units", "Sold"])

        required = ["ProductName", "Total"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(
                f"Missing required columns in sales tab '{ws.title}': {missing}. "
                f"Found columns: {list(df.columns)}"
            )

        df = df[["ProductName", "Total"]].copy()
        df["Month"] = month_date
        df["ProductName"] = df["ProductName"].astype(str).str.strip()
        df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
        sales_frames.append(df)

    if not sales_frames:
        raise ValueError(
            "No month-wise sales tabs found. Expected tab names like 'June-2024'."
        )

    if ignored:
        st.info(f"Ignoring non-sales tabs: {', '.join(ignored)}")

    sales_df = pd.concat(sales_frames, ignore_index=True)
    return sales_df


def load_leadtime_from_sheet(url: str) -> pd.DataFrame:
    spreadsheet = _open_spreadsheet(url)
    ws = spreadsheet.get_worksheet(0)
    data = ws.get_all_records()
    df = pd.DataFrame(data)
    df.columns = [str(c).strip() for c in df.columns]
    return df
