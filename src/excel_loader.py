import os
import re
import urllib.parse
from typing import Dict, List, Optional, Tuple, Any, Union
import requests
import pandas as pd
import openpyxl

DEFAULT_SHAREPOINT_URL = (
    "https://diazacarreos-my.sharepoint.com/:x:/r/personal/logistica_dag_com_pe/"
    "_layouts/15/Doc.aspx?sourcedoc=%7BB8F7A53A-F369-497E-AF5F-7BA382AF512F%7D"
    "&file=Requ%20%26%20OC.xlsx&action=default&mobileredirect=true&wdOrigin=OUTLOOK-METAOS.FILEBROWSER"
)

LOCAL_DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
DEFAULT_LOCAL_FILE = os.path.join(LOCAL_DATA_DIR, "Requ & OC.xlsx")


def ensure_data_dir() -> str:
    """Ensures local data directory exists."""
    os.makedirs(LOCAL_DATA_DIR, exist_ok=True)
    return LOCAL_DATA_DIR


def transform_sharepoint_download_url(url: str) -> str:
    """
    Transforms a SharePoint web view URL into a direct download URL.
    """
    if not url:
        return ""

    # Replace action=default or action=view with action=download
    if "action=" in url:
        transformed = re.sub(r"action=[^&]+", "action=download", url)
    else:
        sep = "&" if "?" in url else "?"
        transformed = f"{url}{sep}action=download"

    # Add download=1 if not present
    if "download=1" not in transformed:
        transformed += "&download=1"

    return transformed


def download_sharepoint_file(
    url: str,
    cookies_str: Optional[str] = None,
    dest_path: str = DEFAULT_LOCAL_FILE,
) -> Tuple[bool, str]:
    """
    Attempts to download the Excel file from SharePoint.
    Returns (success: bool, message: str).
    """
    ensure_data_dir()
    download_url = transform_sharepoint_download_url(url)

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*",
    }

    cookies = {}
    if cookies_str:
        for cookie in cookies_str.split(";"):
            if "=" in cookie:
                k, v = cookie.strip().split("=", 1)
                cookies[k] = v

    try:
        session = requests.Session()
        resp = session.get(download_url, headers=headers, cookies=cookies, timeout=30, allow_redirects=True)

        # Check if redirected to Microsoft Online Login
        final_url = resp.url
        if "login.microsoftonline.com" in final_url:
            return (
                False,
                "El enlace de SharePoint requiere inicio de sesión en Microsoft 365 / Azure AD. "
                "Para mayor comodidad y seguridad, puedes subir el archivo 'Requ & OC.xlsx' directamente "
                "usando el selector de archivos o colocarlo en la carpeta sincronizada.",
            )

        # Check if the content is an actual Excel file
        content_type = resp.headers.get("Content-Type", "")
        is_zip = resp.content.startswith(b"PK\x03\x04")  # Standard ZIP magic bytes for .xlsx

        if not is_zip:
            if resp.status_code == 200:
                return (
                    False,
                    f"La respuesta del servidor no fue un archivo Excel válido (recibido: {content_type}). "
                    "El tenant de SharePoint bloquea descargas anónimas.",
                )
            else:
                return False, f"Error HTTP {resp.status_code} al intentar descargar el archivo."

        # Save downloaded file
        with open(dest_path, "wb") as f:
            f.write(resp.content)

        return True, f"Archivo descargado exitosamente y guardado en {dest_path}"

    except Exception as e:
        return False, f"Excepción durante la descarga: {str(e)}"


def get_excel_sheet_names(file_source: Union[str, Any]) -> List[str]:
    """Retrieves sheet names from an Excel file path or uploaded file buffer."""
    try:
        wb = openpyxl.load_workbook(file_source, read_only=True)
        sheets = wb.sheetnames
        wb.close()
        return sheets
    except Exception:
        # Fallback using pd.ExcelFile
        try:
            xl = pd.ExcelFile(file_source)
            return xl.sheet_names
        except Exception:
            return ["Hoja1"]


def load_excel_data(
    file_source: Union[str, Any],
    sheet_name: Optional[str] = None,
    header_row: int = 0,
) -> Tuple[Optional[pd.DataFrame], List[str], str]:
    """
    Loads Excel data into a pandas DataFrame.
    Returns: (df, sheet_names, error_message)
    """
    try:
        sheets = get_excel_sheet_names(file_source)
        selected_sheet = sheet_name if sheet_name in sheets else sheets[0]

        df = pd.read_excel(
            file_source,
            sheet_name=selected_sheet,
            header=header_row,
            dtype=object,  # keep all values as strings/objects to avoid losing leading zeros
        )

        # Clean column names
        cleaned_cols = []
        for i, col in enumerate(df.columns):
            if pd.isna(col) or str(col).startswith("Unnamed:"):
                cleaned_cols.append(f"Columna_{chr(65 + i) if i < 26 else i+1}")
            else:
                cleaned_cols.append(str(col).strip())
        df.columns = cleaned_cols

        return df, sheets, ""
    except Exception as e:
        return None, [], f"Error al cargar archivo Excel: {str(e)}"


def get_column_p_name(df: pd.DataFrame) -> Tuple[Optional[str], int]:
    """
    Finds the 16th column (Column P, 0-based index 15) in the DataFrame.
    Returns (column_name, column_index).
    """
    if df is None or df.empty:
        return None, -1

    # Column P is index 15
    if len(df.columns) > 15:
        return df.columns[15], 15

    # If fewer than 16 columns, check for names like 'Factura', 'Comprobante', 'N° Doc', 'Documento'
    for col in df.columns:
        c_upper = str(col).upper()
        if any(term in c_upper for term in ["FACTURA", "COMPROBANTE", "N° DOC", "NRO DOC", "DOC", "SERIE"]):
            return col, list(df.columns).index(col)

    # Fallback to last column
    return df.columns[-1], len(df.columns) - 1
