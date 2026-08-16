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
# ANNA DARPAN DAILY AUTOMATION - CLOUD / GITHUB ACTIONS
#
# DAILY:
#   Today From Date = Today
#   Today To Date   = Today
#
# REPORTS:
#   DSI
#   DSR
#   INSPECTION
#   MASTER EXCEL
#
# WHATSAPP:
#   Text summary
#   Master Excel document
#
# GITHUB SECRETS:
#   ANNA_DARPAN_TOKEN
#   WHATSAPP_TOKEN
#   PHONE_NUMBER_ID
#   TO_NUMBER
#
# Optional:
#   WHATSAPP_API_VERSION
# ============================================================


# ============================================================
# BASE / OUTPUT
# ============================================================

BASE = Path(__file__).resolve().parent

OUTPUT = BASE / "OUTPUT"
OUTPUT.mkdir(parents=True, exist_ok=True)


# ============================================================
# TARGET SHEDS
# ============================================================

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


DEPOT_ID = "1696"


# ============================================================
# TIMEZONE
# ============================================================

IST = ZoneInfo("Asia/Kolkata")


# ============================================================
# GITHUB SECRETS
# ============================================================

def get_env(*names):
    """
    Multiple secret names support.
    First non-empty value will be used.
    """

    for name in names:

        value = os.getenv(name, "").strip()

        if value:
            return value

    return ""


def load_anna_token():

    token = get_env(
        "ANNA_DARPAN_TOKEN",
        "ANNA_TOKEN",
        "TOKEN",
    )

    return token


def load_whatsapp_config():

    whatsapp_token = get_env(
        "WHATSAPP_TOKEN",
        "WHATSAPP_ACCESS_TOKEN",
        "WA_TOKEN",
    )

    phone_number_id = get_env(
        "PHONE_NUMBER_ID",
        "WHATSAPP_PHONE_NUMBER_ID",
    )

    to_number = get_env(
        "TO_NUMBER",
        "WHATSAPP_TO_NUMBER",
    )

    api_version = get_env(
        "WHATSAPP_API_VERSION",
    )

    if not api_version:
        api_version = "v25.0"

    return (
        whatsapp_token,
        phone_number_id,
        to_number,
        api_version,
    )


# ============================================================
# TODAY DATE
# ============================================================

def get_today_dates():

    now = datetime.now(IST)

    today = now.strftime("%d-%m-%Y")

    return today, today


# ============================================================
# ANNA DARPAN API
# ============================================================

def download_report(
    name,
    endpoint,
    referer,
    from_date,
    to_date,
):

    print("\n" + "=" * 70)
    print("DOWNLOADING", name)
    print("=" * 70)

    token = load_anna_token()

    if not token:

        raise RuntimeError(
            "ANNA_DARPAN_TOKEN secret not found."
        )

    headers = {

        "accept": (
            "application/json, "
            "text/plain, */*"
        ),

        "accept-language": "en",

        "authorization":
            f"Bearer {token}",

        "content-type":
            "application/json",

        "depotid":
            DEPOT_ID,

        "origin":
            "https://www.annadarpan.in",

        "referer":
            referer,

        "user-agent":
            (
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Chrome/151.0.0.0 "
                "Safari/537.36"
            ),
    }


    params = {

        "fromDate":
            from_date,

        "toDate":
            to_date,

        "commodity":
            0,

        "cropyearId":
            0,

        "shed":
            0,

        "format":
            "html",

        "unit":
            "MT",
    }


    response = requests.get(

        endpoint,

        headers=headers,

        params=params,

        timeout=300,
    )


    print(
        "HTTP STATUS :",
        response.status_code,
    )


    if response.status_code != 200:

        print(
            response.text[:2000]
        )

        raise RuntimeError(
            f"{name} API failed "
            f"HTTP {response.status_code}"
        )


    try:

        data = response.json()

    except Exception:

        raise RuntimeError(
            f"{name}: API did not return JSON"
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
        json_file,
    )

    print(
        "HTML Saved :",
        html_file,
    )

    print(
        "HTML Length:",
        len(html),
    )


    if not html.strip():

        raise RuntimeError(
            f"{name}: HTML report is empty"
        )


    return html


# ============================================================
# SHED EXTRACTION
# ============================================================

def extract_shed(value):

    if pd.isna(value):

        return ""


    text = str(value).strip()


    # --------------------------------------------------------
    # Examples:
    #
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
    # Shed: 85
    # Shed-85
    # --------------------------------------------------------

    match = re.search(
        r"\bShed\s*[:\-]?\s*(\d+)\b",
        text,
        re.I,
    )


    if match:

        return match.group(1)


    # --------------------------------------------------------
    # Plain number
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

def find_shed_stack_column(
    df,
    report_name,
):

    # --------------------------------------------------------
    # 1. Named column
    # --------------------------------------------------------

    for col in df.columns:

        name = str(col).strip().lower()

        normalized = (
            name
            .replace(" ", "")
            .replace("_", "")
        )


        if (
            "shed/stack" in name
            or "shed / stack" in name
            or "shedstack" in normalized
        ):

            return col


    # --------------------------------------------------------
    # 2. Inspection -> Shed
    # --------------------------------------------------------

    if report_name == "INSPECTION":

        for col in df.columns:

            name = (
                str(col)
                .strip()
                .lower()
            )

            if name == "shed":

                return col


    # --------------------------------------------------------
    # 3. Generic content detection
    #
    # VERY IMPORTANT FOR DSR
    #
    # DSR tables can have numeric column names:
    #
    # 0, 1, 2, 3...
    #
    # Therefore we inspect ALL columns instead of blindly
    # assuming column 2.
    # --------------------------------------------------------

    best_col = None
    best_score = 0


    for col in df.columns:

        sample = (
            df[col]
            .dropna()
            .astype(str)
            .head(200)
        )


        if len(sample) == 0:
            continue


        score = 0


        for value in sample:

            value = value.strip()


            # 87/87D06
            if re.match(
                r"^\s*\d+\s*/",
                value,
            ):

                score += 1


            # Shed 85
            elif re.search(
                r"\bShed\s*[:\-]?\s*\d+\b",
                value,
                re.I,
            ):

                score += 1


        if score > best_score:

            best_score = score
            best_col = col


    if best_col is not None:

        print(
            f"Detected {report_name} "
            f"Shed/Stack column:",
            best_col,
            "score=",
            best_score,
        )

        return best_col


    # --------------------------------------------------------
    # 4. DSI / DSR fallback
    # --------------------------------------------------------

    if report_name in ["DSI", "DSR"]:

        if len(df.columns) > 2:

            print(
                f"{report_name}: "
                "Using fallback column 2"
            )

            return df.columns[2]


    return None


# ============================================================
# FILTER TARGET SHEDS
# ============================================================

def filter_target_sheds(
    df,
    report_name,
):

    if df.empty:

        return pd.DataFrame()


    df = df.copy()


    shed_col = find_shed_stack_column(
        df,
        report_name,
    )


    if shed_col is None:

        print(
            "\nWARNING:",
            report_name,
            "Shed/Stack column NOT FOUND",
        )

        print(
            "Available columns:"
        )

        for col in df.columns:

            print(
                " -",
                col,
            )


        return pd.DataFrame()


    print(
        "\nUsing Shed Column:",
        shed_col,
    )


    df["Shed_Clean"] = (
        df[shed_col]
        .apply(extract_shed)
    )


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
            "NO SHEDS DETECTED"
        )


    clean = df[
        df["Shed_Clean"].isin(
            TARGET_SHEDS
        )
    ].copy()


    return clean


# ============================================================
# HTML PARSER
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
            f"{report_name}: "
            f"HTML table parsing failed: {e}"
        )


    print(
        "Tables Found:",
        len(tables),
    )


    if not tables:

        raise RuntimeError(
            f"{report_name}: "
            "No HTML tables found"
        )


    frames = []


    for table in tables:

        if table is None:
            continue

        if table.empty:
            continue


        frame = pd.DataFrame(
            table
        )


        if not frame.empty:

            frames.append(
                frame
            )


    if not frames:

        raise RuntimeError(
            f"{report_name}: "
            "No usable tables"
        )


    raw = pd.concat(
        frames,
        ignore_index=True,
    )


    print(
        "All Rows:",
        len(raw),
    )


    # --------------------------------------------------------
    # Save complete report
    # --------------------------------------------------------

    all_file = (
        OUTPUT /
        f"{report_name}_ALL_DATA.xlsx"
    )


    raw.to_excel(
        all_file,
        index=False,
    )


    print(
        "All Data:",
        all_file,
    )


    # --------------------------------------------------------
    # Filter target sheds
    # --------------------------------------------------------

    clean = filter_target_sheds(
        raw,
        report_name,
    )


    target_file = (
        OUTPUT /
        f"{report_name}_13_TARGET_SHEDS.xlsx"
    )


    clean.to_excel(
        target_file,
        index=False,
    )


    print(
        "13 Shed Rows:",
        len(clean),
    )


    print(
        "Target File:",
        target_file,
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


    possible_files = [

        OUTPUT /
        "STACKWISE_INSPECTION_13_TARGET_SHEDS.xlsx",

        BASE /
        "STACKWISE_INSPECTION_13_TARGET_SHEDS.xlsx",

        OUTPUT /
        "INSPECTION_13_TARGET_SHEDS.xlsx",

        BASE /
        "INSPECTION_13_TARGET_SHEDS.xlsx",
    ]


    source = None


    for file in possible_files:

        if file.exists():

            source = file
            break


    if source is None:

        print(
            "Inspection source file NOT FOUND."
        )

        print(
            "Inspection will be 0."
        )

        return pd.DataFrame()


    print(
        "Inspection Source:",
        source,
    )


    try:

        df = pd.read_excel(
            source
        )

    except Exception as e:

        print(
            "Inspection read error:",
            e,
        )

        return pd.DataFrame()


    clean = filter_target_sheds(
        df,
        "INSPECTION",
    )


    print(
        "Inspection Target Rows:",
        len(clean),
    )


    return clean


# ============================================================
# SHED SUMMARY
# ============================================================

def get_shed_count(
    df,
    shed,
):

    if df is None:
        return 0


    if df.empty:
        return 0


    if "Shed_Clean" not in df.columns:
        return 0


    return int(
        (
            df["Shed_Clean"]
            .astype(str)
            == str(shed)
        ).sum()
    )


# ============================================================
# BUILD MASTER EXCEL
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
        "BUILDING MASTER EXCEL"
    )

    print(
        "=" * 70
    )


    rows = []


    ordered_sheds = sorted(
        TARGET_SHEDS,
        key=lambda x: int(x),
    )


    for shed in ordered_sheds:

        rows.append({

            "Shed":
                shed,

            "DSI":
                get_shed_count(
                    dsi,
                    shed,
                ),

            "DSR":
                get_shed_count(
                    dsr,
                    shed,
                ),

            "Inspection":
                get_shed_count(
                    inspection,
                    shed,
                ),
        })


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
        "\nMASTER EXCEL CREATED:"
    )

    print(
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
# WHATSAPP TEXT
# ============================================================

def build_whatsapp_message(
    from_date,
    to_date,
    dsi,
    dsr,
    inspection,
    shed_summary,
):

    message = []


    message.append(
        "🌾 ANNA DARPAN DAILY REPORT"
    )

    message.append(
        "--------------------------------"
    )

    message.append(
        f"📅 From : {from_date}"
    )

    message.append(
        f"📅 To   : {to_date}"
    )

    message.append("")

    message.append(
        "📊 13 SHED SUMMARY"
    )

    message.append(
        "--------------------------------"
    )

    message.append(
        f"DSI         : {len(dsi)}"
    )

    message.append(
        f"DSR         : {len(dsr)}"
    )

    message.append(
        f"Inspection  : {len(inspection)}"
    )

    message.append("")

    message.append(
        "Shed | DSI | DSR | Inspection"
    )

    message.append(
        "--------------------------------"
    )


    for _, row in shed_summary.iterrows():

        message.append(

            f"{row['Shed']} | "
            f"{row['DSI']} | "
            f"{row['DSR']} | "
            f"{row['Inspection']}"

        )


    message.append("")

    message.append(
        "📎 Master Excel attached below."
    )


    return "\n".join(
        message
    )


# ============================================================
# WHATSAPP API
# ============================================================

def whatsapp_headers(
    whatsapp_token,
):

    return {

        "Authorization":
            f"Bearer {whatsapp_token}",

        "Content-Type":
            "application/json",
    }


# ============================================================
# SEND WHATSAPP TEXT
# ============================================================

def send_whatsapp_text(
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


    (
        whatsapp_token,
        phone_number_id,
        to_number,
        api_version,
    ) = load_whatsapp_config()


    if not whatsapp_token:

        raise RuntimeError(
            "WHATSAPP_TOKEN secret not found."
        )


    if not phone_number_id:

        raise RuntimeError(
            "PHONE_NUMBER_ID secret not found."
        )


    if not to_number:

        raise RuntimeError(
            "TO_NUMBER secret not found."
        )


    url = (

        f"https://graph.facebook.com/"
        f"{api_version}/"
        f"{phone_number_id}/messages"
    )


    payload = {

        "messaging_product":
            "whatsapp",

        "to":
            to_number,

        "type":
            "text",

        "text": {

            "preview_url":
                False,

            "body":
                message,
        },
    }


    response = requests.post(

        url,

        headers=whatsapp_headers(
            whatsapp_token
        ),

        json=payload,

        timeout=120,
    )


    print(
        "HTTP STATUS :",
        response.status_code,
    )

    print(
        "RESPONSE    :",
        response.text,
    )


    if response.status_code not in [
        200,
        201,
    ]:

        raise RuntimeError(
            "WhatsApp text message failed."
        )


    print(
        "SUCCESS WhatsApp text sent."
    )


# ============================================================
# UPLOAD EXCEL TO WHATSAPP
# ============================================================

def upload_whatsapp_document(
    file_path,
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


    (
        whatsapp_token,
        phone_number_id,
        to_number,
        api_version,
    ) = load_whatsapp_config()


    url = (

        f"https://graph.facebook.com/"
        f"{api_version}/"
        f"{phone_number_id}/media"
    )


    headers = {

        "Authorization":
            f"Bearer {whatsapp_token}",
    }


    data = {

        "messaging_product":
            "whatsapp",
    }


    with open(
        file_path,
        "rb",
    ) as file:

        files = {

            "file": (

                file_path.name,

                file,

                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet",
            )
        }


        response = requests.post(

            url,

            headers=headers,

            data=data,

            files=files,

            timeout=300,
        )


    print(
        "UPLOAD STATUS :",
        response.status_code,
    )

    print(
        "UPLOAD RESPONSE:",
        response.text,
    )


    if response.status_code not in [
        200,
        201,
    ]:

        raise RuntimeError(
            "WhatsApp Excel upload failed."
        )


    result = response.json()


    media_id = result.get(
        "id"
    )


    if not media_id:

        raise RuntimeError(
            "WhatsApp media ID not received."
        )


    print(
        "Media ID:",
        media_id,
    )


    return media_id


# ============================================================
# SEND EXCEL DOCUMENT
# ============================================================

def send_whatsapp_document(
    media_id,
    file_path,
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


    (
        whatsapp_token,
        phone_number_id,
        to_number,
        api_version,
    ) = load_whatsapp_config()


    url = (

        f"https://graph.facebook.com/"
        f"{api_version}/"
        f"{phone_number_id}/messages"
    )


    payload = {

        "messaging_product":
            "whatsapp",

        "to":
            to_number,

        "type":
            "document",

        "document": {

            "id":
                media_id,

            "filename":
                file_path.name,

            "caption":
                "🌾 Anna Darpan Daily Master Excel",
        },
    }


    response = requests.post(

        url,

        headers=whatsapp_headers(
            whatsapp_token
        ),

        json=payload,

        timeout=120,
    )


    print(
        "DOCUMENT STATUS :",
        response.status_code,
    )

    print(
        "DOCUMENT RESPONSE:",
        response.text,
    )


    if response.status_code not in [
        200,
        201,
    ]:

        raise RuntimeError(
            "WhatsApp Excel document send failed."
        )


    print(
        "SUCCESS Excel document sent to WhatsApp."
    )


# ============================================================
# VALIDATE SECRETS
# ============================================================

def validate_secrets():

    print(
        "\n" + "=" * 70
    )

    print(
        "CHECKING GITHUB SECRETS"
    )

    print(
        "=" * 70
    )


    anna_token = load_anna_token()


    (
        whatsapp_token,
        phone_number_id,
        to_number,
        api_version,
    ) = load_whatsapp_config()


    if anna_token:

        print(
            "Anna Darpan Token : OK"
        )

    else:

        print(
            "Anna Darpan Token : MISSING"
        )


    if whatsapp_token:

        print(
            "WhatsApp Token    : OK"
        )

    else:

        print(
            "WhatsApp Token    : MISSING"
        )


    if phone_number_id:

        print(
            "Phone Number ID   : OK"
        )

    else:

        print(
            "Phone Number ID   : MISSING"
        )


    if to_number:

        print(
            "WhatsApp To       : OK"
        )

    else:

        print(
            "WhatsApp To       : MISSING"
        )


    print(
        "WhatsApp API      :",
        api_version,
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
            "Phone Number ID missing."
        )


    if not to_number:

        raise RuntimeError(
            "WhatsApp TO number missing."
        )


# ============================================================
# MAIN
# ============================================================

def main():

    print(
        "\n" + "=" * 70
    )

    print(
        "ANNA DARPAN DAILY AUTOMATION - CLOUD"
    )

    print(
        "=" * 70
    )


    # --------------------------------------------------------
    # Secrets
    # --------------------------------------------------------

    validate_secrets()


    # --------------------------------------------------------
    # TODAY ONLY
    # --------------------------------------------------------

    from_date, to_date = (
        get_today_dates()
    )


    print(
        "\nDAILY REPORT DATE"
    )

    print(
        "-" * 70
    )

    print(
        "From :",
        from_date,
    )

    print(
        "To   :",
        to_date,
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


    # ========================================================
    # API ENDPOINTS
    # ========================================================

    DSI_ENDPOINT = (

        "https://adbackend.annadarpan.in/"
        "prdannadarpan.in/reports/api/v2/"
        "DSIReport"
    )


    DSR_ENDPOINT = (

        "https://adbackend.annadarpan.in/"
        "prdannadarpan.in/reports/api/v2/"
        "DSRReport"
    )


    # ========================================================
    # DOWNLOAD DSI
    # ========================================================

    dsi_html = download_report(

        "DSI",

        DSI_ENDPOINT,

        (
            "https://www.annadarpan.in/"
            "reporting/depotDSIReport"
        ),

        from_date,

        to_date,
    )


    # ========================================================
    # DOWNLOAD DSR
    # ========================================================

    dsr_html = download_report(

        "DSR",

        DSR_ENDPOINT,

        (
            "https://www.annadarpan.in/"
            "reporting/depotDSRReport"
        ),

        from_date,

        to_date,
    )


    # ========================================================
    # PARSE DSI
    # ========================================================

    dsi = parse_report(
        "DSI",
        dsi_html,
    )


    # ========================================================
    # PARSE DSR
    # ========================================================

    dsr = parse_report(
        "DSR",
        dsr_html,
    )


    # ========================================================
    # INSPECTION
    # ========================================================

    inspection = process_inspection()


    # ========================================================
    # MASTER
    # ========================================================

    master, shed_summary = build_master(

        dsi,

        dsr,

        inspection,

        from_date,

        to_date,
    )


    # ========================================================
    # WHATSAPP MESSAGE
    # ========================================================

    message = build_whatsapp_message(

        from_date,

        to_date,

        dsi,

        dsr,

        inspection,

        shed_summary,
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
        message
    )


    # ========================================================
    # SEND TEXT
    # ========================================================

    send_whatsapp_text(
        message
    )


    # ========================================================
    # UPLOAD EXCEL
    # ========================================================

    media_id = upload_whatsapp_document(
        master
    )


    # ========================================================
    # SEND EXCEL
    # ========================================================

    send_whatsapp_document(

        media_id,

        master,
    )


    # ========================================================
    # FINAL
    # ========================================================

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


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    try:

        main()

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
            e
        )

        raise
