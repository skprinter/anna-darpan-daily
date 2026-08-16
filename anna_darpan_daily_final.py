import os
import re
import json
from io import StringIO
from pathlib import Path
from datetime import datetime

import requests
import pandas as pd


# ============================================================
# ANNA DARPAN DAILY AUTOMATION - FINAL
#
# DSI + DSR + 13 SHEDS + INSPECTION
# + MASTER EXCEL
# + WHATSAPP TEXT
# + WHATSAPP EXCEL DOCUMENT
#
# DSR FIX:
#   - Parse HTML table-by-table
#   - Do NOT depend only on "Shed/Stack" header
#   - Automatically detect stack-like column
#   - Handles changing table/header positions
# ============================================================


BASE = Path(__file__).resolve().parent
OUTPUT = BASE / "OUTPUT"
OUTPUT.mkdir(exist_ok=True)

TARGET_SHEDS = {
    "74", "75", "76", "77", "79", "80", "81",
    "82", "85", "88", "89", "90", "103"
}

DEPOT_ID = "1696"

WHATSAPP_API_VERSION = "v23.0"


# ============================================================
# TOKEN
# ============================================================

def load_token():
    token_file = BASE / "token.txt"

    if token_file.exists():
        token = token_file.read_text(
            encoding="utf-8"
        ).strip()

        if token:
            return token

    return os.getenv(
        "ANNA_DARPAN_TOKEN",
        ""
    ).strip()


# ============================================================
# WHATSAPP CONFIG
# ============================================================

def load_whatsapp_config():
    config_file = BASE / "whatsapp_config.txt"

    if not config_file.exists():
        raise RuntimeError(
            "whatsapp_config.txt not found"
        )

    config = {}

    for line in config_file.read_text(
        encoding="utf-8"
    ).splitlines():

        line = line.strip()

        if not line or line.startswith("#"):
            continue

        if "=" not in line:
            continue

        key, value = line.split("=", 1)
        config[key.strip()] = value.strip()

    token = config.get("TOKEN", "")
    phone_number_id = config.get("PHONE_NUMBER_ID", "")
    to_number = config.get("TO_NUMBER", "")

    if not token:
        raise RuntimeError(
            "WhatsApp TOKEN missing"
        )

    if not phone_number_id:
        raise RuntimeError(
            "PHONE_NUMBER_ID missing"
        )

    if not to_number:
        raise RuntimeError(
            "TO_NUMBER missing"
        )

    to_number = re.sub(
        r"\D",
        "",
        to_number
    )

    return token, phone_number_id, to_number


# ============================================================
# WHATSAPP TEXT
# ============================================================

def send_whatsapp_message(message):

    print("\n" + "=" * 70)
    print("SENDING WHATSAPP MESSAGE")
    print("=" * 70)

    token, phone_number_id, to_number = (
        load_whatsapp_config()
    )

    url = (
        f"https://graph.facebook.com/"
        f"{WHATSAPP_API_VERSION}/"
        f"{phone_number_id}/messages"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    payload = {
        "messaging_product": "whatsapp",
        "to": to_number,
        "type": "text",
        "text": {
            "preview_url": False,
            "body": message
        }
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=60
    )

    print("HTTP STATUS :", response.status_code)
    print("RESPONSE    :", response.text)

    if response.status_code not in (200, 201):
        raise RuntimeError(
            "WhatsApp text message failed"
        )

    print("SUCCESS WhatsApp text sent.")


# ============================================================
# WHATSAPP EXCEL DOCUMENT
# ============================================================

def upload_whatsapp_document(file_path):

    print("\n" + "=" * 70)
    print("UPLOADING MASTER EXCEL TO WHATSAPP")
    print("=" * 70)

    token, phone_number_id, _ = (
        load_whatsapp_config()
    )

    url = (
        f"https://graph.facebook.com/"
        f"{WHATSAPP_API_VERSION}/"
        f"{phone_number_id}/media"
    )

    headers = {
        "Authorization": f"Bearer {token}"
    }

    mime_type = (
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    )

    with open(file_path, "rb") as f:

        files = {
            "file": (
                Path(file_path).name,
                f,
                mime_type
            )
        }

        data = {
            "messaging_product": "whatsapp",
            "type": mime_type
        }

        response = requests.post(
            url,
            headers=headers,
            data=data,
            files=files,
            timeout=120
        )

    print("UPLOAD STATUS :", response.status_code)
    print("UPLOAD RESPONSE:", response.text)

    if response.status_code not in (200, 201):
        raise RuntimeError(
            "WhatsApp media upload failed"
        )

    result = response.json()
    media_id = result.get("id")

    if not media_id:
        raise RuntimeError(
            "WhatsApp media ID not returned"
        )

    print("Media ID :", media_id)

    return media_id


def send_whatsapp_document(
    media_id,
    file_name,
    caption=""
):

    print("\n" + "=" * 70)
    print("SENDING MASTER EXCEL TO WHATSAPP")
    print("=" * 70)

    token, phone_number_id, to_number = (
        load_whatsapp_config()
    )

    url = (
        f"https://graph.facebook.com/"
        f"{WHATSAPP_API_VERSION}/"
        f"{phone_number_id}/messages"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    document = {
        "id": media_id,
        "filename": file_name
    }

    if caption:
        document["caption"] = caption

    payload = {
        "messaging_product": "whatsapp",
        "to": to_number,
        "type": "document",
        "document": document
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload,
        timeout=60
    )

    print("DOCUMENT STATUS :", response.status_code)
    print("DOCUMENT RESPONSE:", response.text)

    if response.status_code not in (200, 201):
        raise RuntimeError(
            "WhatsApp document message failed"
        )

    print(
        "SUCCESS Excel document sent to WhatsApp."
    )


# ============================================================
# ANNA DARPAN DOWNLOAD
# ============================================================

def download_report(
    name,
    endpoint,
    referer,
    from_date,
    to_date
):

    print("\n" + "=" * 70)
    print("Downloading", name)
    print("=" * 70)

    token = load_token()

    if not token:
        raise RuntimeError(
            "Anna Darpan token not found"
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
        "origin": "https://www.annadarpan.in",
        "referer": referer,
        "user-agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/151.0.0.0 Safari/537.36"
        )
    }

    params = {
        "fromDate": from_date,
        "toDate": to_date,
        "commodity": 0,
        "cropyearId": 0,
        "shed": 0,
        "format": "html",
        "unit": "MT"
    }

    response = requests.get(
        endpoint,
        headers=headers,
        params=params,
        timeout=300
    )

    print("HTTP Status :", response.status_code)

    if response.status_code != 200:
        print(response.text[:1000])

        raise RuntimeError(
            f"{name} API failed: "
            f"HTTP {response.status_code}"
        )

    data = response.json()

    json_file = OUTPUT / f"{name}_RESPONSE.json"
    html_file = OUTPUT / f"{name}_REPORT.html"

    json_file.write_text(
        json.dumps(
            data,
            ensure_ascii=False
        ),
        encoding="utf-8"
    )

    html = data.get("value", "")

    if not html:
        raise RuntimeError(
            f"{name} API returned empty HTML"
        )

    html_file.write_text(
        html,
        encoding="utf-8"
    )

    print("JSON Saved :", json_file)
    print("HTML Saved :", html_file)
    print("HTML Length:", len(html))

    return html


# ============================================================
# SHED EXTRACTION
# ============================================================

def extract_shed(value):

    if pd.isna(value):
        return ""

    text = str(value).strip()

    # 87/87D06
    # 100/100A08
    # 13/13C05
    match = re.match(
        r"^\s*(\d+)\s*/",
        text
    )

    if match:
        return match.group(1)

    # Shed 85 / Shed: 85
    match = re.search(
        r"\bShed\s*[:\-]?\s*(\d+)\b",
        text,
        re.I
    )

    if match:
        return match.group(1)

    # Plain number
    match = re.match(
        r"^\s*(\d+)(?:\.0)?\s*$",
        text
    )

    if match:
        return match.group(1)

    return ""


def is_stack_value(value):

    if pd.isna(value):
        return False

    text = str(value).strip()

    if re.match(
        r"^\s*\d+\s*/",
        text
    ):
        return True

    if re.search(
        r"\bShed\s*[:\-]?\s*\d+\b",
        text,
        re.I
    ):
        return True

    return False


# ============================================================
# ROBUST SHED COLUMN DETECTION
# ============================================================

def score_shed_column(series):

    values = (
        series
        .dropna()
        .astype(str)
        .str.strip()
    )

    if values.empty:
        return 0

    stack_count = int(
        values.apply(is_stack_value).sum()
    )

    target_count = int(
        values.apply(
            lambda x:
            extract_shed(x) in TARGET_SHEDS
        ).sum()
    )

    # Target shed matches are stronger evidence.
    return (
        stack_count * 2
        + target_count * 10
    )


def find_best_shed_column(
    df,
    report_name
):

    # 1. Explicit header
    for col in df.columns:

        name = str(col).strip().lower()

        if (
            "shed/stack" in name
            or "shed / stack" in name
            or name == "shed"
            or "shed stack" in name
        ):
            return col

    # 2. Score every column.
    # This is the important DSR fix.
    scores = []

    for col in df.columns:

        score = score_shed_column(
            df[col]
        )

        scores.append(
            (score, col)
        )

    scores.sort(
        key=lambda x: x[0],
        reverse=True
    )

    if scores and scores[0][0] > 0:

        best_score, best_col = scores[0]

        print(
            f"Using detected Shed/Stack column "
            f"{best_col} (score={best_score})"
        )

        return best_col

    return None


# ============================================================
# ROBUST HTML PARSER
# ============================================================

def parse_report(
    report_name,
    html
):

    print(
        "\nParsing",
        report_name,
        "HTML..."
    )

    print("-" * 70)

    # header=None is intentional:
    # different tables have different header rows.
    tables = pd.read_html(
        StringIO(html),
        header=None
    )

    print(
        "Tables Found :",
        len(tables)
    )

    # --------------------------------------------------------
    # Save complete raw parse
    # --------------------------------------------------------

    raw_frames = []

    for table in tables:

        if table is None or table.empty:
            continue

        frame = pd.DataFrame(table).copy()
        frame.columns = range(frame.shape[1])

        raw_frames.append(frame)

    if not raw_frames:
        raise RuntimeError(
            f"No HTML tables found for {report_name}"
        )

    raw = pd.concat(
        raw_frames,
        ignore_index=True,
        sort=False
    )

    print(
        "All Rows     :",
        len(raw)
    )

    all_file = (
        OUTPUT /
        f"{report_name}_ALL_DATA.xlsx"
    )

    try:
        raw.to_excel(
            all_file,
            index=False
        )
    except PermissionError:
        print(
            "WARNING:",
            all_file.name,
            "is open."
        )

    # --------------------------------------------------------
    # Extract rows table-by-table.
    #
    # This prevents DSR tables with different headers from
    # destroying the Shed/Stack position after concat.
    # --------------------------------------------------------

    selected_frames = []
    table_hits = []

    for table_no, table in enumerate(
        tables,
        start=1
    ):

        if table is None or table.empty:
            continue

        frame = pd.DataFrame(table).copy()
        frame.columns = range(frame.shape[1])

        shed_col = find_best_shed_column(
            frame,
            report_name
        )

        if shed_col is None:
            continue

        frame["Shed_Clean"] = (
            frame[shed_col]
            .apply(extract_shed)
        )

        selected = frame[
            frame["Shed_Clean"].isin(
                TARGET_SHEDS
            )
        ].copy()

        if not selected.empty:

            selected["Source_Table"] = table_no

            selected_frames.append(
                selected
            )

            table_hits.append(
                (table_no, len(selected))
            )

    if selected_frames:

        clean = pd.concat(
            selected_frames,
            ignore_index=True,
            sort=False
        )

        # Remove exact duplicates caused by repeated HTML
        # header/table rendering.
        clean = clean.drop_duplicates(
            ignore_index=True
        )

    else:
        clean = pd.DataFrame()

    # --------------------------------------------------------
    # Output
    # --------------------------------------------------------

    target_file = (
        OUTPUT /
        f"{report_name}_13_TARGET_SHEDS.xlsx"
    )

    try:

        clean.to_excel(
            target_file,
            index=False
        )

    except PermissionError:

        target_file = (
            OUTPUT /
            (
                f"{report_name}_13_TARGET_SHEDS_"
                + datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )
                + ".xlsx"
            )
        )

        clean.to_excel(
            target_file,
            index=False
        )

        print(
            "WARNING: Original target file was open."
        )

        print(
            "Saved new file:",
            target_file
        )

    print(
        "\n13 Shed Rows:",
        len(clean)
    )

    print(
        "Output      :",
        target_file
    )

    if table_hits:

        print(
            "\nTables contributing target rows:"
        )

        print(
            table_hits
        )

    if not clean.empty:

        print(
            "\nShed Wise Count"
        )

        print("-" * 50)

        print(
            clean["Shed_Clean"]
            .value_counts()
            .sort_index()
            .to_string()
        )

    else:

        print(
            "\nWARNING:",
            report_name,
            "returned ZERO target shed rows."
        )

        # For debugging, show top column scores.
        print(
            "Column detection failed for target rows."
        )

    return clean


# ============================================================
# INSPECTION
# ============================================================

def process_inspection():

    print("\n" + "=" * 70)
    print("Inspection Data")
    print("=" * 70)

    source = (
        OUTPUT /
        "STACKWISE_INSPECTION_13_TARGET_SHEDS.xlsx"
    )

    if not source.exists():

        print(
            "Inspection file not found:",
            source
        )

        return pd.DataFrame()

    try:

        df = pd.read_excel(source)

    except PermissionError:

        print(
            "WARNING: Inspection file is open."
        )

        return pd.DataFrame()

    # Inspection normally has Shed column.
    shed_col = None

    for col in df.columns:

        if str(col).strip().lower() == "shed":
            shed_col = col
            break

    if shed_col is None:

        for col in df.columns:

            name = str(col).strip().lower()

            if "shed" in name:
                shed_col = col
                break

    if shed_col is None:

        print(
            "WARNING: Inspection Shed column not found."
        )

        return pd.DataFrame()

    print(
        "Using Shed Column :",
        shed_col
    )

    df = df.copy()

    df["Shed_Clean"] = (
        df[shed_col]
        .apply(extract_shed)
    )

    clean = df[
        df["Shed_Clean"].isin(
            TARGET_SHEDS
        )
    ].copy()

    print(
        "\nInspection target rows:",
        len(clean)
    )

    if not clean.empty:

        print(
            clean["Shed_Clean"]
            .value_counts()
            .sort_index()
            .to_string()
        )

    return clean


# ============================================================
# BUILD MASTER
# ============================================================

def build_master(
    dsi,
    dsr,
    inspection,
    from_date,
    to_date
):

    print("\n" + "=" * 70)
    print("BUILDING MASTER DASHBOARD")
    print("=" * 70)

    rows = []

    ordered_sheds = sorted(
        TARGET_SHEDS,
        key=lambda x: int(x)
    )

    for shed in ordered_sheds:

        dsi_count = 0
        dsr_count = 0
        inspection_count = 0

        if "Shed_Clean" in dsi.columns:

            dsi_count = int(
                (
                    dsi["Shed_Clean"] == shed
                ).sum()
            )

        if "Shed_Clean" in dsr.columns:

            dsr_count = int(
                (
                    dsr["Shed_Clean"] == shed
                ).sum()
            )

        if "Shed_Clean" in inspection.columns:

            inspection_count = int(
                (
                    inspection["Shed_Clean"] == shed
                ).sum()
            )

        rows.append({
            "Shed": shed,
            "DSI_Rows": dsi_count,
            "DSR_Rows": dsr_count,
            "Inspection_Rows": inspection_count
        })

    shed_summary = pd.DataFrame(rows)

    summary = pd.DataFrame(
        [
            [
                "Generated On",
                datetime.now().strftime(
                    "%d-%m-%Y %I:%M:%S %p"
                )
            ],
            ["From Date", from_date],
            ["To Date", to_date],
            ["Target Sheds", len(TARGET_SHEDS)],
            ["DSI Records", len(dsi)],
            ["DSR Records", len(dsr)],
            ["Inspection Records", len(inspection)]
        ],
        columns=["Item", "Value"]
    )

    master = (
        OUTPUT /
        "ANNA_DARPAN_DAILY_MASTER.xlsx"
    )

    def write_master(path):

        with pd.ExcelWriter(
            path,
            engine="openpyxl"
        ) as writer:

            summary.to_excel(
                writer,
                sheet_name="Summary",
                index=False
            )

            shed_summary.to_excel(
                writer,
                sheet_name="Shed_Summary",
                index=False
            )

            dsi.to_excel(
                writer,
                sheet_name="DSI_13_Sheds",
                index=False
            )

            dsr.to_excel(
                writer,
                sheet_name="DSR_13_Sheds",
                index=False
            )

            inspection.to_excel(
                writer,
                sheet_name="Inspection_13_Sheds",
                index=False
            )

    try:

        write_master(master)

    except PermissionError:

        master = (
            OUTPUT /
            (
                "ANNA_DARPAN_DAILY_MASTER_"
                + datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )
                + ".xlsx"
            )
        )

        write_master(master)

        print(
            "WARNING: Existing master file was open."
        )

        print(
            "New master:",
            master
        )

    print("\nSUCCESS")
    print("Master :", master)

    print("\nShed Wise Summary")
    print("-" * 70)
    print(
        shed_summary.to_string(
            index=False
        )
    )

    return master, shed_summary


# ============================================================
# WHATSAPP MESSAGE
# ============================================================

def create_whatsapp_message(
    dsi,
    dsr,
    inspection,
    from_date,
    to_date
):

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
        f"DSI         : {len(dsi)}"
    )

    lines.append(
        f"DSR         : {len(dsr)}"
    )

    lines.append(
        f"Inspection  : {len(inspection)}"
    )

    lines.append("")

    lines.append(
        "Shed | DSI | DSR | Inspection"
    )

    lines.append(
        "--------------------------------"
    )

    for shed in sorted(
        TARGET_SHEDS,
        key=lambda x: int(x)
    ):

        dsi_count = 0
        dsr_count = 0
        inspection_count = 0

        if "Shed_Clean" in dsi.columns:

            dsi_count = int(
                (
                    dsi["Shed_Clean"] == shed
                ).sum()
            )

        if "Shed_Clean" in dsr.columns:

            dsr_count = int(
                (
                    dsr["Shed_Clean"] == shed
                ).sum()
            )

        if "Shed_Clean" in inspection.columns:

            inspection_count = int(
                (
                    inspection["Shed_Clean"] == shed
                ).sum()
            )

        lines.append(
            f"{shed} | {dsi_count} | "
            f"{dsr_count} | "
            f"{inspection_count}"
        )

    lines.append("")
    lines.append(
        "📎 Master Excel attached below."
    )

    return "\n".join(lines)


# ============================================================
# MAIN
# ============================================================

print("\n" + "=" * 70)
print("ANNA DARPAN DAILY AUTOMATION - FINAL")
print("=" * 70)

if not load_token():

    print(
        "\nERROR: Anna Darpan token not found."
    )

    raise SystemExit

print("\nAnna Darpan Token : OK")

try:

    whatsapp_token, phone_id, to_number = (
        load_whatsapp_config()
    )

    print("WhatsApp Config   : OK")
    print(
        "WhatsApp Number   :",
        to_number
    )

except Exception as e:

    print(
        "\nWhatsApp Config Error:",
        e
    )

    raise SystemExit


from zoneinfo import ZoneInfo

# ============================================================
# AUTOMATIC DAILY DATE - INDIA
# ============================================================

india_now = datetime.now(
    ZoneInfo("Asia/Kolkata")
)

from_date = india_now.strftime("%d-%m-%Y")
to_date = from_date

print(
    "\nAutomatic Daily Report Date"
)

print(
    "-" * 70
)

print(
    "India Date :",
    from_date
)

print(
    "From Date  :",
    from_date
)

print(
    "To Date    :",
    to_date
)

print("\nParameters")
print("-" * 70)
print("From :", from_date)
print("To   :", to_date)
print(
    "Target Sheds:",
    ", ".join(
        sorted(
            TARGET_SHEDS,
            key=lambda x: int(x)
        )
    )
)


try:

    DSI_ENDPOINT = (
        "https://adbackend.annadarpan.in/"
        "prdannadarpan.in/reports/api/v2/DSIReport"
    )

    DSR_ENDPOINT = (
        "https://adbackend.annadarpan.in/"
        "prdannadarpan.in/reports/api/v2/DSRReport"
    )

    # --------------------------------------------------------
    # DSI
    # --------------------------------------------------------

    dsi_html = download_report(
        "DSI",
        DSI_ENDPOINT,
        (
            "https://www.annadarpan.in/"
            "reporting/depotDSIReport"
        ),
        from_date,
        to_date
    )

    # --------------------------------------------------------
    # DSR
    # --------------------------------------------------------

    dsr_html = download_report(
        "DSR",
        DSR_ENDPOINT,
        (
            "https://www.annadarpan.in/"
            "reporting/depotDSRReport"
        ),
        from_date,
        to_date
    )

    # --------------------------------------------------------
    # PARSE
    # --------------------------------------------------------

    dsi = parse_report(
        "DSI",
        dsi_html
    )

    dsr = parse_report(
        "DSR",
        dsr_html
    )

    # IMPORTANT:
    # Do not silently continue with DSR=0.
    # If API has HTML tables but parser gets zero target rows,
    # stop and show failure clearly.
    if len(dsr) == 0:

        raise RuntimeError(
            "DSR extraction returned 0 target shed rows. "
            "Open OUTPUT\\DSR_REPORT.html and verify the report "
            "structure/date range."
        )

    # --------------------------------------------------------
    # INSPECTION
    # --------------------------------------------------------

    inspection = process_inspection()

    # --------------------------------------------------------
    # MASTER
    # --------------------------------------------------------

    master, shed_summary = build_master(
        dsi,
        dsr,
        inspection,
        from_date,
        to_date
    )

    # --------------------------------------------------------
    # WHATSAPP TEXT
    # --------------------------------------------------------

    whatsapp_message = create_whatsapp_message(
        dsi,
        dsr,
        inspection,
        from_date,
        to_date
    )

    print("\n" + "=" * 70)
    print("WHATSAPP MESSAGE PREVIEW")
    print("=" * 70)
    print(whatsapp_message)

    send_whatsapp_message(
        whatsapp_message
    )

    # --------------------------------------------------------
    # WHATSAPP EXCEL ATTACHMENT
    # --------------------------------------------------------

    media_id = upload_whatsapp_document(
        master
    )

    send_whatsapp_document(
        media_id,
        Path(master).name,
        "📊 ANNA DARPAN DAILY MASTER EXCEL"
    )

    # --------------------------------------------------------
    # FINAL
    # --------------------------------------------------------

    print("\n" + "=" * 70)
    print("ALL DAILY PROCESSING COMPLETED")
    print("=" * 70)

    print("\nMaster File:")
    print(master)

    print("\nWhatsApp Text: SENT")
    print("WhatsApp Excel: SENT")

except Exception as e:

    print("\n" + "=" * 70)
    print("DAILY PROCESS FAILED")
    print("=" * 70)

    print(
        type(e).__name__,
        ":",
        e
    )

    raise
