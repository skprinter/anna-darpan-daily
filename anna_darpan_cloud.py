import os
import re
import json
from io import StringIO
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import requests
import pandas as pd


# ============================================================
# ANNA DARPAN DAILY AUTOMATION
# LOCAL CMD + GITHUB ACTIONS
#
# DSI + DSR + INSPECTION + 13 SHEDS + MASTER EXCEL
# + WHATSAPP TEXT + WHATSAPP EXCEL
# ============================================================


# ============================================================
# PATHS
# ============================================================

BASE = Path(__file__).resolve().parent

OUTPUT = BASE / "OUTPUT"
OUTPUT.mkdir(exist_ok=True)

TOKEN_FILE = BASE / "token.txt"
WHATSAPP_CONFIG_FILE = BASE / "whatsapp_config.txt"


# ============================================================
# SETTINGS
# ============================================================

DEPOT_ID = "1696"

WHATSAPP_API_VERSION = "v25.0"

TARGET_SHEDS = {
    "74",
    "75",
    "76",
    "77",
    "79",
    "80",
    "81",
    "82",
    "85",
    "88",
    "89",
    "90",
    "103",
}


IST = ZoneInfo("Asia/Kolkata")


# ============================================================
# COMMON HELPERS
# ============================================================

def clean_secret(value):
    """
    Makes secret/config values safe.

    Supports:
        TOKEN=xxxxx
        WHATSAPP_TOKEN=xxxxx
        "xxxxx"
        'xxxxx'
        Bearer xxxxx
    """

    if value is None:
        return ""

    value = str(value).strip()

    if not value:
        return ""

    # Remove surrounding quotes
    if len(value) >= 2:
        if (
            value.startswith('"')
            and value.endswith('"')
        ):
            value = value[1:-1].strip()

        elif (
            value.startswith("'")
            and value.endswith("'")
        ):
            value = value[1:-1].strip()

    # Remove common prefixes
    prefixes = [
        "TOKEN=",
        "ANNA_DARPAN_TOKEN=",
        "WHATSAPP_TOKEN=",
        "PHONE_NUMBER_ID=",
        "TO_NUMBER=",
    ]

    upper = value.upper()

    for prefix in prefixes:

        if upper.startswith(prefix):

            value = value[
                len(prefix):
            ].strip()

            break

    # Remove quotes again
    if len(value) >= 2:

        if (
            value.startswith('"')
            and value.endswith('"')
        ):
            value = value[1:-1].strip()

        elif (
            value.startswith("'")
            and value.endswith("'")
        ):
            value = value[1:-1].strip()

    return value.strip()


def normalize_bearer_token(token):
    token = clean_secret(token)

    if token.lower().startswith("bearer "):
        token = token[7:].strip()

    return token


# ============================================================
# LOCAL CONFIG FILE
# ============================================================

def load_local_whatsapp_config():

    config = {}

    if not WHATSAPP_CONFIG_FILE.exists():
        return config

    try:

        lines = WHATSAPP_CONFIG_FILE.read_text(
            encoding="utf-8"
        ).splitlines()

    except Exception as e:

        print(
            "WARNING: Could not read whatsapp_config.txt:",
            e
        )

        return config

    for line in lines:

        line = line.strip()

        if not line:
            continue

        if line.startswith("#"):
            continue

        if "=" not in line:
            continue

        key, value = line.split(
            "=",
            1
        )

        key = key.strip().upper()
        value = clean_secret(value)

        config[key] = value

    return config


# ============================================================
# ANNA DARPAN TOKEN
# ============================================================

def load_anna_darpan_token():

    # --------------------------------------------------------
    # 1. GitHub / Environment
    # --------------------------------------------------------

    token = os.getenv(
        "ANNA_DARPAN_TOKEN",
        ""
    ).strip()

    if token:
        return normalize_bearer_token(token)

    # --------------------------------------------------------
    # 2. Local token.txt
    # --------------------------------------------------------

    if TOKEN_FILE.exists():

        try:

            content = TOKEN_FILE.read_text(
                encoding="utf-8"
            ).strip()

            if content:

                # Handle TOKEN=xxxxx
                if "=" in content:

                    first_line = content.splitlines()[0]

                    key, value = first_line.split(
                        "=",
                        1
                    )

                    if key.strip().upper() in [
                        "TOKEN",
                        "ANNA_DARPAN_TOKEN",
                    ]:
                        content = value.strip()

                return normalize_bearer_token(
                    content
                )

        except Exception as e:

            print(
                "WARNING: token.txt read error:",
                e
            )

    return ""


# ============================================================
# WHATSAPP CONFIG
# ============================================================

def load_whatsapp_config():

    local = load_local_whatsapp_config()

    # Environment / GitHub Secrets first
    whatsapp_token = os.getenv(
        "WHATSAPP_TOKEN",
        ""
    ).strip()

    phone_number_id = os.getenv(
        "PHONE_NUMBER_ID",
        ""
    ).strip()

    to_number = os.getenv(
        "TO_NUMBER",
        ""
    ).strip()

    # Local config fallback
    if not whatsapp_token:

        whatsapp_token = local.get(
            "WHATSAPP_TOKEN",
            ""
        )

        if not whatsapp_token:

            # Some older config files may use TOKEN
            whatsapp_token = local.get(
                "TOKEN",
                ""
            )

    if not phone_number_id:

        phone_number_id = local.get(
            "PHONE_NUMBER_ID",
            ""
        )

    if not to_number:

        to_number = local.get(
            "TO_NUMBER",
            ""
        )

    return (
        normalize_bearer_token(
            whatsapp_token
        ),
        clean_secret(phone_number_id),
        clean_secret(to_number),
    )


# ============================================================
# VALIDATE SECRETS
# ============================================================

def validate_secrets():

    anna_token = load_anna_darpan_token()

    whatsapp_token, phone_number_id, to_number = (
        load_whatsapp_config()
    )

    print(
        "\n" + "=" * 70
    )

    print(
        "CHECKING CREDENTIALS"
    )

    print(
        "=" * 70
    )

    print(
        "Anna Darpan Token :",
        "OK" if anna_token else "MISSING"
    )

    print(
        "WhatsApp Token    :",
        "OK" if whatsapp_token else "MISSING"
    )

    print(
        "Phone Number ID   :",
        "OK" if phone_number_id else "MISSING"
    )

    print(
        "WhatsApp To       :",
        "OK" if to_number else "MISSING"
    )

    print(
        "WhatsApp API      :",
        WHATSAPP_API_VERSION
    )

    if not anna_token:

        raise RuntimeError(
            "Anna Darpan token missing."
        )

    if not whatsapp_token:

        raise RuntimeError(
            "WhatsApp token missing."
        )

    if not phone_number_id:

        raise RuntimeError(
            "PHONE_NUMBER_ID missing."
        )

    if not to_number:

        raise RuntimeError(
            "TO_NUMBER missing."
        )

    return (
        anna_token,
        whatsapp_token,
        phone_number_id,
        to_number,
    )


# ============================================================
# DAILY DATE
# ============================================================

def get_daily_dates():

    now = datetime.now(IST)

    date_string = now.strftime(
        "%d-%m-%Y"
    )

    return (
        date_string,
        date_string,
    )


# ============================================================
# ANNA DARPAN API
# ============================================================

def download_report(
    name,
    endpoint,
    referer,
    from_date,
    to_date,
    anna_token,
):

    print(
        "\n" + "=" * 70
    )

    print(
        "DOWNLOADING",
        name
    )

    print(
        "=" * 70
    )

    token = normalize_bearer_token(
        anna_token
    )

    headers = {
        "accept": (
            "application/json, "
            "text/plain, */*"
        ),
        "accept-language": "en",
        "authorization": f"Bearer {token}",
        "content-type": "application/json",
        "depotid": DEPOT_ID,
        "origin": (
            "https://www.annadarpan.in"
        ),
        "referer": referer,
        "user-agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/151.0.0.0 Safari/537.36"
        ),
    }

    params = {
        "fromDate": from_date,
        "toDate": to_date,
        "commodity": 0,
        "cropyearId": 0,
        "shed": 0,
        "format": "html",
        "unit": "MT",
    }

    try:

        response = requests.get(
            endpoint,
            headers=headers,
            params=params,
            timeout=300,
        )

    except requests.RequestException as e:

        raise RuntimeError(
            f"{name} API connection failed: {e}"
        )

    print(
        "HTTP STATUS :",
        response.status_code
    )

    if response.status_code != 200:

        print(
            response.text[:1500]
        )

        if response.status_code == 401:

            raise RuntimeError(
                f"{name} API failed HTTP 401. "
                "Anna Darpan token is invalid/expired "
                "or authentication is not accepted."
            )

        raise RuntimeError(
            f"{name} API failed HTTP "
            f"{response.status_code}"
        )

    try:

        data = response.json()

    except Exception:

        raise RuntimeError(
            f"{name} API returned invalid JSON."
        )

    json_file = (
        OUTPUT /
        f"{name}_RESPONSE.json"
    )

    html_file = (
        OUTPUT /
        f"{name}_REPORT.html"
    )

    json_file.write_text(
        json.dumps(
            data,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    html = data.get(
        "value",
        "",
    )

    if html is None:
        html = ""

    html = str(html)

    html_file.write_text(
        html,
        encoding="utf-8",
    )

    print(
        "JSON Saved :",
        json_file
    )

    print(
        "HTML Saved :",
        html_file
    )

    print(
        "HTML Length:",
        len(html)
    )

    if not html.strip():

        raise RuntimeError(
            f"{name} API returned empty HTML."
        )

    return html


# ============================================================
# FLATTEN COLUMNS
# ============================================================

def flatten_columns(df):

    df = df.copy()

    new_columns = []

    for col in df.columns:

        if isinstance(col, tuple):

            parts = []

            for item in col:

                text = str(item).strip()

                if text:
                    parts.append(text)

            name = " ".join(parts)

        else:

            name = str(col).strip()

        name = re.sub(
            r"\s+",
            " ",
            name
        )

        new_columns.append(
            name.strip()
        )

    df.columns = new_columns

    return df


# ============================================================
# SHED EXTRACTION
# ============================================================

def extract_shed(value):

    if pd.isna(value):
        return ""

    text = str(value).strip()

    if not text:
        return ""

    # --------------------------------------------------------
    # 87/87D06
    # 100/100A08
    # 13/13C05
    # --------------------------------------------------------

    match = re.match(
        r"^\s*(\d+)\s*/",
        text,
    )

    if match:

        return match.group(1)

    # --------------------------------------------------------
    # Shed 85
    # --------------------------------------------------------

    match = re.search(
        r"\bShed\s*[:\-]?\s*(\d+)\b",
        text,
        re.I,
    )

    if match:

        return match.group(1)

    # --------------------------------------------------------
    # Shed/Stack: 85/85A01
    # --------------------------------------------------------

    match = re.search(
        r"\b(\d+)\s*/\s*\d+[A-Za-z]",
        text,
        re.I,
    )

    if match:

        return match.group(1)

    # --------------------------------------------------------
    # Plain numeric
    # --------------------------------------------------------

    match = re.match(
        r"^\s*(\d+)(?:\.0)?\s*$",
        text,
    )

    if match:

        return match.group(1)

    return ""


# ============================================================
# FIND SHED COLUMN
# ============================================================

def score_shed_column(series):

    try:

        sample = (
            series.dropna()
            .astype(str)
            .head(300)
        )

    except Exception:

        return 0

    if sample.empty:
        return 0

    score = 0

    for value in sample:

        text = value.strip()

        if re.match(
            r"^\d+\s*/",
            text,
        ):
            score += 3

        elif re.search(
            r"\bShed\s*[:\-]?\s*\d+",
            text,
            re.I,
        ):
            score += 3

        elif re.match(
            r"^\d+(?:\.0)?$",
            text,
        ):
            score += 1

    return score


def find_shed_stack_column(
    df,
    report_name,
):

    # --------------------------------------------------------
    # Named header
    # --------------------------------------------------------

    preferred_names = [
        "shed/stack",
        "shed / stack",
        "shedstack",
        "shed",
        "shed no",
        "shed number",
        "stack",
    ]

    for col in df.columns:

        name = str(col).strip().lower()

        name = re.sub(
            r"\s+",
            " ",
            name,
        )

        if name in preferred_names:

            return col

        if (
            "shed/stack" in name
            or "shed / stack" in name
            or "shed stack" in name
        ):

            return col

    # --------------------------------------------------------
    # Score every column
    # This is especially important for DSR.
    # --------------------------------------------------------

    scored = []

    for col in df.columns:

        score = score_shed_column(
            df[col]
        )

        if score > 0:

            scored.append(
                (
                    score,
                    col,
                )
            )

    if scored:

        scored.sort(
            key=lambda x: x[0],
            reverse=True,
        )

        best_score, best_col = scored[0]

        print(
            "Auto detected Shed/Stack column:",
            best_col,
            "Score:",
            best_score,
        )

        return best_col

    # --------------------------------------------------------
    # Legacy fallback
    # DSI / DSR often have column 2
    # --------------------------------------------------------

    if report_name in [
        "DSI",
        "DSR",
    ]:

        if len(df.columns) > 2:

            col = df.columns[2]

            if (
                score_shed_column(
                    df[col]
                ) > 0
            ):

                print(
                    "Using fallback Shed/Stack Column: 2"
                )

                return col

    return None


# ============================================================
# FILTER TARGET SHEDS
# ============================================================

def filter_target_sheds(
    df,
    report_name,
):

    if df is None or df.empty:

        return pd.DataFrame()

    df = flatten_columns(
        df.copy()
    )

    shed_col = find_shed_stack_column(
        df,
        report_name,
    )

    if shed_col is None:

        print(
            "\nWARNING:",
            report_name,
            "Shed/Stack column not found."
        )

        print(
            "Available columns:"
        )

        for col in df.columns:

            print(
                " -",
                col,
            )

        return pd.DataFrame(
            columns=df.columns
        )

    print(
        "\nUsing Shed Column:",
        shed_col
    )

    df["Shed_Clean"] = (
        df[shed_col]
        .apply(extract_shed)
    )

    # --------------------------------------------------------
    # Debug detected sheds
    # --------------------------------------------------------

    detected = (
        df.loc[
            df["Shed_Clean"] != "",
            "Shed_Clean",
        ]
        .value_counts()
        .sort_index()
    )

    print(
        "\nDetected Shed Counts"
    )

    print(
        "-" * 50
    )

    if len(detected) > 0:

        print(
            detected.to_string()
        )

    else:

        print(
            "No shed numbers detected."
        )

    # --------------------------------------------------------
    # Filter target
    # --------------------------------------------------------

    clean = df[
        df["Shed_Clean"].isin(
            TARGET_SHEDS
        )
    ].copy()

    return clean


# ============================================================
# PARSE REPORT
# ============================================================

def parse_report(
    report_name,
    html,
):

    print(
        "\n" + "=" * 70
    )

    print(
        "PARSING",
        report_name,
    )

    print(
        "=" * 70
    )

    try:

        tables = pd.read_html(
            StringIO(html)
        )

    except Exception as e:

        raise RuntimeError(
            f"{report_name} HTML parsing failed: {e}"
        )

    print(
        "Tables Found:",
        len(tables)
    )

    frames = []

    for index, table in enumerate(tables):

        if table is None:
            continue

        if table.empty:
            continue

        frame = pd.DataFrame(
            table
        )

        frame = flatten_columns(
            frame
        )

        print(
            f"Table {index + 1}:",
            len(frame),
            "rows x",
            len(frame.columns),
            "columns"
        )

        frames.append(
            frame
        )

    if not frames:

        raise RuntimeError(
            f"No HTML tables found for {report_name}"
        )

    raw = pd.concat(
        frames,
        ignore_index=True,
        sort=False,
    )

    raw = flatten_columns(
        raw
    )

    print(
        "All Rows:",
        len(raw)
    )

    # --------------------------------------------------------
    # Save complete data
    # --------------------------------------------------------

    all_file = (
        OUTPUT /
        f"{report_name}_ALL_DATA.xlsx"
    )

    try:

        raw.to_excel(
            all_file,
            index=False,
        )

        print(
            "All Data Saved:",
            all_file
        )

    except PermissionError:

        print(
            "WARNING:",
            all_file.name,
            "is open."
        )

    # --------------------------------------------------------
    # Filter 13 sheds
    # --------------------------------------------------------

    clean = filter_target_sheds(
        raw,
        report_name,
    )

    target_file = (
        OUTPUT /
        f"{report_name}_13_TARGET_SHEDS.xlsx"
    )

    try:

        clean.to_excel(
            target_file,
            index=False,
        )

    except PermissionError:

        target_file = (
            OUTPUT /
            (
                f"{report_name}_13_TARGET_SHEDS_"
                +
                datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )
                +
                ".xlsx"
            )
        )

        clean.to_excel(
            target_file,
            index=False,
        )

    print(
        "\n13 Shed Rows:",
        len(clean)
    )

    print(
        "Target Output:",
        target_file
    )

    if not clean.empty:

        print(
            "\nShed Wise Count"
        )

        print(
            "-" * 50
        )

        print(
            clean[
                "Shed_Clean"
            ]
            .value_counts()
            .sort_index()
            .to_string()
        )

    return clean


# ============================================================
# INSPECTION
# ============================================================

def find_inspection_file():

    # Exact known file first
    candidates = [
        OUTPUT /
        "STACKWISE_INSPECTION_13_TARGET_SHEDS.xlsx",

        OUTPUT /
        "INSPECTION_13_TARGET_SHEDS.xlsx",

        BASE /
        "STACKWISE_INSPECTION_13_TARGET_SHEDS.xlsx",

        BASE /
        "INSPECTION_13_TARGET_SHEDS.xlsx",
    ]

    for file in candidates:

        if file.exists():

            return file

    # Search for inspection xlsx
    search_locations = [
        OUTPUT,
        BASE,
    ]

    for location in search_locations:

        try:

            files = list(
                location.glob(
                    "*INSPECTION*.xlsx"
                )
            )

        except Exception:

            files = []

        for file in files:

            if file.name == (
                "ANNA_DARPAN_DAILY_MASTER.xlsx"
            ):
                continue

            return file

    return None


def process_inspection():

    print(
        "\n" + "=" * 70
    )

    print(
        "INSPECTION DATA"
    )

    print(
        "=" * 70
    )

    source = find_inspection_file()

    if source is None:

        print(
            "Inspection file not found."
        )

        print(
            "Inspection will be treated as empty."
        )

        return pd.DataFrame()

    print(
        "Inspection Source:",
        source
    )

    try:

        df = pd.read_excel(
            source
        )

    except Exception as e:

        print(
            "WARNING: Inspection read failed:",
            e
        )

        return pd.DataFrame()

    clean = filter_target_sheds(
        df,
        "INSPECTION",
    )

    print(
        "Inspection Target Rows:",
        len(clean)
    )

    return clean


# ============================================================
# MASTER DASHBOARD
# ============================================================

def build_master(
    dsi,
    dsr,
    inspection,
    from_date,
    to_date,
):

    print(
        "\n" + "=" * 70
    )

    print(
        "BUILDING MASTER DASHBOARD"
    )

    print(
        "=" * 70
    )

    rows = []

    ordered_sheds = sorted(
        TARGET_SHEDS,
        key=lambda x: int(x)
    )

    for shed in ordered_sheds:

        dsi_count = 0

        if (
            dsi is not None
            and not dsi.empty
            and "Shed_Clean" in dsi.columns
        ):

            dsi_count = int(
                (
                    dsi[
                        "Shed_Clean"
                    ] == shed
                ).sum()
            )

        dsr_count = 0

        if (
            dsr is not None
            and not dsr.empty
            and "Shed_Clean" in dsr.columns
        ):

            dsr_count = int(
                (
                    dsr[
                        "Shed_Clean"
                    ] == shed
                ).sum()
            )

        inspection_count = 0

        if (
            inspection is not None
            and not inspection.empty
            and "Shed_Clean"
            in inspection.columns
        ):

            inspection_count = int(
                (
                    inspection[
                        "Shed_Clean"
                    ] == shed
                ).sum()
            )

        rows.append(
            {
                "Shed": shed,
                "DSI_Rows": dsi_count,
                "DSR_Rows": dsr_count,
                "Inspection_Rows":
                    inspection_count,
            }
        )

    shed_summary = pd.DataFrame(
        rows
    )

    summary = pd.DataFrame(
        [
            [
                "Generated On",
                datetime.now(
                    IST
                ).strftime(
                    "%d-%m-%Y %I:%M:%S %p"
                ),
            ],
            [
                "From Date",
                from_date,
            ],
            [
                "To Date",
                to_date,
            ],
            [
                "Target Sheds",
                len(TARGET_SHEDS),
            ],
            [
                "DSI Records",
                len(dsi),
            ],
            [
                "DSR Records",
                len(dsr),
            ],
            [
                "Inspection Records",
                len(inspection),
            ],
        ],
        columns=[
            "Item",
            "Value",
        ],
    )

    master = (
        OUTPUT /
        "ANNA_DARPAN_DAILY_MASTER.xlsx"
    )

    try:

        with pd.ExcelWriter(
            master,
            engine="openpyxl",
        ) as writer:

            summary.to_excel(
                writer,
                sheet_name="Summary",
                index=False,
            )

            shed_summary.to_excel(
                writer,
                sheet_name="Shed_Summary",
                index=False,
            )

            dsi.to_excel(
                writer,
                sheet_name="DSI_13_Sheds",
                index=False,
            )

            dsr.to_excel(
                writer,
                sheet_name="DSR_13_Sheds",
                index=False,
            )

            inspection.to_excel(
                writer,
                sheet_name="Inspection_13_Sheds",
                index=False,
            )

    except PermissionError:

        master = (
            OUTPUT /
            (
                "ANNA_DARPAN_DAILY_MASTER_"
                +
                datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )
                +
                ".xlsx"
            )
        )

        with pd.ExcelWriter(
            master,
            engine="openpyxl",
        ) as writer:

            summary.to_excel(
                writer,
                sheet_name="Summary",
                index=False,
            )

            shed_summary.to_excel(
                writer,
                sheet_name="Shed_Summary",
                index=False,
            )

            dsi.to_excel(
                writer,
                sheet_name="DSI_13_Sheds",
                index=False,
            )

            dsr.to_excel(
                writer,
                sheet_name="DSR_13_Sheds",
                index=False,
            )

            inspection.to_excel(
                writer,
                sheet_name="Inspection_13_Sheds",
                index=False,
            )

    print(
        "\nSUCCESS"
    )

    print(
        "Master:",
        master
    )

    print(
        "\nShed Wise Summary"
    )

    print(
        "-" * 70
    )

    print(
        shed_summary.to_string(
            index=False
        )
    )

    return master, shed_summary


# ============================================================
# WHATSAPP PHONE NORMALIZATION
# ============================================================

def normalize_phone_number(number):

    number = str(
        number or ""
    ).strip()

    # Remove +, spaces, -, brackets
    number = re.sub(
        r"[\s\-\(\)\+]",
        "",
        number
    )

    # If Indian 10 digit number supplied,
    # add 91.
    if (
        len(number) == 10
        and number.startswith("6")
        or len(number) == 10
        and number.startswith("7")
        or len(number) == 10
        and number.startswith("8")
        or len(number) == 10
        and number.startswith("9")
    ):

        number = "91" + number

    return number


# ============================================================
# WHATSAPP SUMMARY
# ============================================================

def build_whatsapp_message(
    from_date,
    to_date,
    dsi,
    dsr,
    inspection,
    shed_summary,
):

    dsi_total = len(dsi)

    dsr_total = len(dsr)

    inspection_total = len(
        inspection
    )

    lines = []

    lines.append(
        "🌾 ANNA DARPAN DAILY REPORT"
    )

    lines.append(
        "--------------------------------"
    )

    lines.append(
        f"📅 From : {from_date}"
    )

    lines.append(
        f"📅 To   : {to_date}"
    )

    lines.append("")

    lines.append(
        "📊 13 SHED SUMMARY"
    )

    lines.append(
        "--------------------------------"
    )

    lines.append(
        f"DSI         : {dsi_total}"
    )

    lines.append(
        f"DSR         : {dsr_total}"
    )

    lines.append(
        f"Inspection  : {inspection_total}"
    )

    lines.append("")

    lines.append(
        "Shed | DSI | DSR | Inspection"
    )

    lines.append(
        "--------------------------------"
    )

    for _, row in shed_summary.iterrows():

        lines.append(
            f"{row['Shed']} | "
            f"{int(row['DSI_Rows'])} | "
            f"{int(row['DSR_Rows'])} | "
            f"{int(row['Inspection_Rows'])}"
        )

    lines.append("")

    lines.append(
        "📎 Master Excel attached below."
    )

    return "\n".join(
        lines
    )


# ============================================================
# WHATSAPP SEND TEXT
# ============================================================

def send_whatsapp_text(
    whatsapp_token,
    phone_number_id,
    to_number,
    message,
):

    print(
        "\n" + "=" * 70
    )

    print(
        "SENDING WHATSAPP MESSAGE"
    )

    print(
        "=" * 70
    )

    phone_number_id = clean_secret(
        phone_number_id
    )

    to_number = normalize_phone_number(
        to_number
    )

    url = (
        f"https://graph.facebook.com/"
        f"{WHATSAPP_API_VERSION}/"
        f"{phone_number_id}/messages"
    )

    headers = {
        "Authorization":
            f"Bearer {normalize_bearer_token(whatsapp_token)}",

        "Content-Type":
            "application/json",
    }

    payload = {
        "messaging_product":
            "whatsapp",

        "to":
            to_number,

        "type":
            "text",

        "text":
            {
                "preview_url": False,
                "body": message,
            },
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=120,
    )

    print(
        "HTTP STATUS :",
        response.status_code
    )

    print(
        "RESPONSE    :",
        response.text[:2000]
    )

    if response.status_code not in [
        200,
        201,
    ]:

        raise RuntimeError(
            "WhatsApp text send failed: "
            f"HTTP {response.status_code}"
        )

    print(
        "SUCCESS WhatsApp text sent."
    )


# ============================================================
# WHATSAPP UPLOAD EXCEL
# ============================================================

def upload_whatsapp_document(
    whatsapp_token,
    phone_number_id,
    excel_file,
):

    print(
        "\n" + "=" * 70
    )

    print(
        "UPLOADING MASTER EXCEL TO WHATSAPP"
    )

    print(
        "=" * 70
    )

    phone_number_id = clean_secret(
        phone_number_id
    )

    url = (
        f"https://graph.facebook.com/"
        f"{WHATSAPP_API_VERSION}/"
        f"{phone_number_id}/media"
    )

    headers = {
        "Authorization":
            f"Bearer {normalize_bearer_token(whatsapp_token)}",
    }

    mime_type = (
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    )

    try:

        with open(
            excel_file,
            "rb",
        ) as file_handle:

            files = {
                "file": (
                    excel_file.name,
                    file_handle,
                    mime_type,
                ),
            }

            data = {
                "messaging_product":
                    "whatsapp",

                "type":
                    mime_type,
            }

            response = requests.post(
                url,
                headers=headers,
                files=files,
                data=data,
                timeout=300,
            )

    except requests.RequestException as e:

        raise RuntimeError(
            f"WhatsApp media upload failed: {e}"
        )

    print(
        "UPLOAD STATUS :",
        response.status_code
    )

    print(
        "UPLOAD RESPONSE:",
        response.text[:2000]
    )

    if response.status_code not in [
        200,
        201,
    ]:

        raise RuntimeError(
            "WhatsApp Excel upload failed: "
            f"HTTP {response.status_code}"
        )

    try:

        result = response.json()

    except Exception:

        raise RuntimeError(
            "WhatsApp media upload returned invalid JSON."
        )

    media_id = result.get(
        "id",
        ""
    )

    if not media_id:

        raise RuntimeError(
            "WhatsApp media ID not received."
        )

    print(
        "Media ID:",
        media_id
    )

    return media_id


# ============================================================
# WHATSAPP SEND EXCEL
# ============================================================

def send_whatsapp_document(
    whatsapp_token,
    phone_number_id,
    to_number,
    media_id,
    excel_file,
):

    print(
        "\n" + "=" * 70
    )

    print(
        "SENDING MASTER EXCEL TO WHATSAPP"
    )

    print(
        "=" * 70
    )

    phone_number_id = clean_secret(
        phone_number_id
    )

    to_number = normalize_phone_number(
        to_number
    )

    url = (
        f"https://graph.facebook.com/"
        f"{WHATSAPP_API_VERSION}/"
        f"{phone_number_id}/messages"
    )

    headers = {
        "Authorization":
            f"Bearer {normalize_bearer_token(whatsapp_token)}",

        "Content-Type":
            "application/json",
    }

    payload = {
        "messaging_product":
            "whatsapp",

        "to":
            to_number,

        "type":
            "document",

        "document":
            {
                "id":
                    media_id,

                "caption":
                    (
                        "🌾 Anna Darpan Daily Master Report\n"
                        f"📅 {datetime.now(IST).strftime('%d-%m-%Y')}"
                    ),

                "filename":
                    excel_file.name,
            },
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=120,
    )

    print(
        "DOCUMENT STATUS :",
        response.status_code
    )

    print(
        "DOCUMENT RESPONSE:",
        response.text[:2000]
    )

    if response.status_code not in [
        200,
        201,
    ]:

        raise RuntimeError(
            "WhatsApp Excel send failed: "
            f"HTTP {response.status_code}"
        )

    print(
        "SUCCESS Excel document sent to WhatsApp."
    )


# ============================================================
# SAVE WHATSAPP PREVIEW
# ============================================================

def save_whatsapp_preview(message):

    file = (
        OUTPUT /
        "WHATSAPP_MESSAGE_PREVIEW.txt"
    )

    file.write_text(
        message,
        encoding="utf-8",
    )

    return file


# ============================================================
# MAIN
# ============================================================

def main():

    print(
        "\n" + "=" * 70
    )

    print(
        "ANNA DARPAN DAILY AUTOMATION - CLOUD + LOCAL"
    )

    print(
        "=" * 70
    )

    # --------------------------------------------------------
    # Credentials
    # --------------------------------------------------------

    (
        anna_token,
        whatsapp_token,
        phone_number_id,
        to_number,
    ) = validate_secrets()

    # --------------------------------------------------------
    # Date
    # --------------------------------------------------------

    from_date, to_date = (
        get_daily_dates()
    )

    print(
        "\n" + "=" * 70
    )

    print(
        "DAILY REPORT DATE"
    )

    print(
        "-" * 70
    )

    print(
        "From :",
        from_date
    )

    print(
        "To   :",
        to_date
    )

    print(
        "\nTarget Sheds:"
    )

    print(
        ", ".join(
            sorted(
                TARGET_SHEDS,
                key=lambda x: int(x),
            )
        )
    )

    # --------------------------------------------------------
    # Endpoints
    # --------------------------------------------------------

    DSI_ENDPOINT = (
        "https://adbackend.annadarpan.in/"
        "prdannadarpan.in/reports/api/v2/DSIReport"
    )

    DSR_ENDPOINT = (
        "https://adbackend.annadarpan.in/"
        "prdannadarpan.in/reports/api/v2/DSRReport"
    )

    DSI_REFERER = (
        "https://www.annadarpan.in/"
        "reporting/depotDSIReport"
    )

    DSR_REFERER = (
        "https://www.annadarpan.in/"
        "reporting/depotDSRReport"
    )

    try:

        # ====================================================
        # DSI
        # ====================================================

        dsi_html = download_report(
            "DSI",
            DSI_ENDPOINT,
            DSI_REFERER,
            from_date,
            to_date,
            anna_token,
        )

        # ====================================================
        # DSR
        # ====================================================

        dsr_html = download_report(
            "DSR",
            DSR_ENDPOINT,
            DSR_REFERER,
            from_date,
            to_date,
            anna_token,
        )

        # ====================================================
        # PARSE
        # ====================================================

        dsi = parse_report(
            "DSI",
            dsi_html,
        )

        dsr = parse_report(
            "DSR",
            dsr_html,
        )

        # ====================================================
        # INSPECTION
        # ====================================================

        inspection = process_inspection()

        # ====================================================
        # MASTER
        # ====================================================

        master, shed_summary = build_master(
            dsi,
            dsr,
            inspection,
            from_date,
            to_date,
        )

        # ====================================================
        # WHATSAPP MESSAGE
        # ====================================================

        whatsapp_message = build_whatsapp_message(
            from_date,
            to_date,
            dsi,
            dsr,
            inspection,
            shed_summary,
        )

        preview_file = save_whatsapp_preview(
            whatsapp_message
        )

        print(
            "\n" + "=" * 70
        )

        print(
            "WHATSAPP MESSAGE PREVIEW"
        )

        print(
            "=" * 70
        )

        print(
            whatsapp_message
        )

        print(
            "\nPreview Saved:",
            preview_file
        )

        # ====================================================
        # SEND WHATSAPP TEXT
        # ====================================================

        send_whatsapp_text(
            whatsapp_token,
            phone_number_id,
            to_number,
            whatsapp_message,
        )

        # ====================================================
        # UPLOAD MASTER EXCEL
        # ====================================================

        media_id = upload_whatsapp_document(
            whatsapp_token,
            phone_number_id,
            master,
        )

        # ====================================================
        # SEND MASTER EXCEL
        # ====================================================

        send_whatsapp_document(
            whatsapp_token,
            phone_number_id,
            to_number,
            media_id,
            master,
        )

        # ====================================================
        # FINAL
        # ====================================================

        print(
            "\n" + "=" * 70
        )

        print(
            "ALL DAILY PROCESSING COMPLETED"
        )

        print(
            "=" * 70
        )

        print(
            "\nMaster File:"
        )

        print(
            master
        )

        print(
            "\nWhatsApp Text: SENT"
        )

        print(
            "WhatsApp Excel: SENT"
        )

    except Exception as e:

        print(
            "\n" + "=" * 70
        )

        print(
            "DAILY PROCESS FAILED"
        )

        print(
            "=" * 70
        )

        print(
            type(e).__name__,
            ":",
            e,
        )

        raise


# ============================================================
# START
# ============================================================

if __name__ == "__main__":

    main()
