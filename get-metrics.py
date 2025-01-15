#!/usr/bin/python3

import argparse
import json
import os
import re
import requests
import string
import subprocess
import sys
import tempfile

from datetime import datetime, timedelta
from tabulate import tabulate

import pandas as pd


THUNDERBIRD_STATUS_VERSIONS = [
    "cf_status_thunderbird_esr128",
    "cf_status_thunderbird_128",
    "cf_status_thunderbird_129",
    "cf_status_thunderbird_130",
    "cf_status_thunderbird_131",
    "cf_status_thunderbird_132",
    "cf_status_thunderbird_133",
    "cf_status_thunderbird_134",
    "cf_status_thunderbird_135",
    "cf_status_thunderbird_136",
]

THUNDERBIRD_DAILY_VERSIONS = [
    "136.0a1",
]

THUNDERBIRD_BETA_VERSIONS = [
    "134.0b1",
    "134.0b2",
    "134.0b3",
    "134.0b4",
    "134.0b5",
    "134.0b6",
    "135.0b1",
    "135.0b2",
    "135.0b3",
    "135.0b4",
    "135.0b5",
    "135.0b6",
]

THUNDERBIRD_RELEASE_VERSIONS = [
    "133.0",
    "133.0.1",
    "133.0.2",
    "133.0.3",
]

# bugzilla.mozilla.org
BMO_QUERY_TYPES = [
    "regression-all",
    "regression-severe",
    "non-regression-all",
    "non-regression-severe",
    "topcrash",
    "perf",
    "sec-crit-high",
    "sec-moderate-low",
]

# stats.thunderbird.net
STN_QUERY_TYPES = [
    "daily-installations",
    "beta-installations",
    "release-installations",
    "total-installations",
]

# crash-stats.mozilla.org
CSMO_QUERY_TYPES = [
    "release-crashes",
    "beta-crashes",
    "daily-crashes",
]


def get_today():
    """Get today's date in YYYY-MM-DD format"""
    today = datetime.now()
    return today.strftime("%Y-%m-%d")


def get_yesterday():
    """Get yesterday's date in YYYY-MM-DD format"""
    today = datetime.now()
    yesterday = today - timedelta(days=1)
    return yesterday.strftime("%Y-%m-%d")


def get_bmo_url(query_type):
    """Get bugzilla.mozilla.org URL"""
    index = 4
    f_version = ""
    o_comparison = ""
    v_status = ""
    for version in THUNDERBIRD_STATUS_VERSIONS:
        f_version = f"{f_version}f{index}={version}&"
        o_comparison = f"{o_comparison}o{index}=equals&"
        v_status = f"{v_status}v{index}=affected&"
        index += 1
    f_version = f"{f_version}f{index}=CP&"

    url = (
        "https://bugzilla.mozilla.org/rest/bug?include_fields=id,summary,status&"
        "bug_type=defect&"
        "chfield=%5BBug%20creation%5D&"
        "f1=short_desc&"
        "f2=component&"
        "f3=OP&"
        "j3=OR&"
        f"{f_version}"
        f"{o_comparison}"
        "resolution=---&"
        "v1=intermit%20perma%20assert%20debug%20ews&"
        "v2=%20add-on%20build%20upstream&"
        f"{v_status}"
    )

    match query_type:
        case "sec-crit-high":
            url = (
                f"{url}"
                "keywords=sec-crit%20sec-high&"
                "keywords_type=anywords&"
                "o1=nowordssubstr&"
                "o2=nowordssubstr&"
            )
        case "sec-moderate-low":
            url = (
                f"{url}"
                "keywords=sec-moderate%20sec-low&"
                "keywords_type=anywords&"
                "o1=nowordssubstr&"
                "o2=nowordssubstr&"
            )
        case "regression-all":
            url = f"{url}" "keywords=regression&" "keywords_type=allwords&"
        case "regression-severe":
            url = (
                f"{url}"
                "keywords=regression&"
                "keywords_type=allwords&"
                "bug_severity=S1&"
                "bug_severity=critical&"
                "bug_severity=S2&"
                "bug_severity=major&"
                "o1=nowordssubstr&"
                "o2=nowordssubstr&"
            )
        case "non-regression-all":
            url = (
                f"{url}"
                "keywords=regression&"
                "keywords_type=nowords&"
                "o1=nowordssubstr&"
                "o2=nowordssubstr&"
            )
        case "non-regression-severe":
            url = (
                f"{url}"
                "keywords=regression&"
                "keywords_type=nowords&"
                "bug_severity=S1&"
                "bug_severity=critical&"
                "bug_severity=S2&"
                "bug_severity=major&"
                "o1=nowordssubstr&"
                "o2=nowordssubstr&"
            )
        case "perf":
            url = (
                f"{url}"
                "keywords=perf&"
                "keywords_type=allwords&"
                "o1=nowordssubstr&"
                "o2=nowordssubstr&"
            )
        case "topcrash":
            url = (
                f"{url}"
                "keywords=topcrash-thunderbird&"
                "keywords_type=allwords&"
                "o1=nowordssubstr&"
                "o2=nowordssubstr&"
            )
        case _:
            sys.exit(f"Unknown query type: {query_type}")

    return url


def get_csmo_url(query_type, date):
    """Get crash-stats.mozilla.org URL"""
    versions = ""
    match query_type:
        case "daily-crashes":
            for version in THUNDERBIRD_DAILY_VERSIONS:
                versions = f"{versions}version={version}&"
        case "beta-crashes":
            for version in THUNDERBIRD_BETA_VERSIONS:
                versions = f"{versions}version={version}&"
        case "release-crashes":
            for version in THUNDERBIRD_RELEASE_VERSIONS:
                versions = f"{versions}version={version}&"

    today_formatted = get_today()
    yesterday_formatted = get_yesterday()
    start_date = f"date=>={yesterday_formatted}T00:00:00.000Z&"
    end_date = f"date=<{today_formatted}T00:00:00.000Z&"

    url = (
        "https://crash-stats.mozilla.org/api/SuperSearch/?product=Thunderbird&"
        f"{versions}"
        f"{start_date}"
        f"{end_date}"
        "_facets=platform&"
        "_facets=release_channel"
    )
    return url


def bmo_query(query_type):
    """Query bugzilla.mozilla.org"""
    api_key = os.getenv("BMO_API_KEY")
    if not api_key:
        sys.exit("BMO_API_KEY is empty. Please export your BMO key.")
    headers = {"Content-type": "application/json"}
    params = {"api_key": api_key}

    r = requests.get(get_bmo_url(query_type), headers=headers, params=params)
    count = len(json.loads(r.text)["bugs"])
    return count


def stn_query(query_type):
    """Query stats.thunderbird.net"""
    r = requests.get("https://stats.thunderbird.net/thunderbird_adi.json")
    yesterday = get_yesterday()
    match query_type:
        case "total-installations":
            count = json.loads(r.text)[yesterday]["count"]
        case "daily-installations":
            version = THUNDERBIRD_DAILY_VERSIONS[0]
            count = json.loads(r.text)[yesterday]["versions"][version]
        case "beta-installations":
            version = THUNDERBIRD_BETA_VERSIONS[0].split("b")[0]
            count = json.loads(r.text)[yesterday]["versions"][version]
        case "release-installations":
            count = 0
            for version in THUNDERBIRD_RELEASE_VERSIONS:
                if version in json.loads(r.text)[yesterday]["versions"]:
                    count += json.loads(r.text)[yesterday]["versions"][version]
        case _:
            sys.exit(f"Unknown query type: {query_type}")
    return count


def csmo_query(query_type, crash_stats_date):
    """Query crash-stats.mozilla.org"""
    r = requests.get(get_csmo_url(query_type, crash_stats_date))
    count = json.loads(r.text)["total"]
    return count


def print_versions():
    affected_versions = re.sub(
        "cf_status_thunderbird_", "", ", ".join(THUNDERBIRD_STATUS_VERSIONS)
    )
    daily_versions = f"{', '.join(THUNDERBIRD_DAILY_VERSIONS)}"
    beta_versions = f"{', '.join(THUNDERBIRD_BETA_VERSIONS)}"
    release_versions = f"{', '.join(THUNDERBIRD_RELEASE_VERSIONS)}"
    table_data = {
        "Category": [
            "bugzilla affected versions",
            "daily versions",
            "beta versions",
            "release versions",
        ],
        "Variable": [
            affected_versions,
            daily_versions,
            beta_versions,
            release_versions,
        ],
    }
    df = pd.DataFrame(table_data)
    table_data = df.values.tolist()
    print(tabulate(table_data, tablefmt="plain", colalign=("left", "left")))


def export_metrics_to_spreadsheet(release_readiness_metrics):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
        df = pd.DataFrame(
            [list(release_readiness_metrics.values())],
            columns=list(release_readiness_metrics.keys()),
        )
        df.insert(0, "Date", [get_today()])
        with pd.ExcelWriter(temp_file.name, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")

            workbook = writer.book
            worksheet = writer.sheets["Sheet1"]

            font = "Helvetica Neue"
            font_size = 10
            header_format = workbook.add_format(
                {"bg_color": "#B0B3B2", "align": "center"}
            )
            header_format.set_font_name(font)
            header_format.set_font_size(font_size)
            header_format.set_text_wrap()
            date_format = workbook.add_format(
                {"bg_color": "#D4D4D4", "align": "center"}
            )
            date_format.set_font_name(font)
            date_format.set_font_size(font_size)
            percentage_format = workbook.add_format(
                {"num_format": "0.00%", "align": "center"}
            )
            percentage_format.set_font_name(font)
            percentage_format.set_font_size(font_size)
            other_format = workbook.add_format({"align": "center"})
            other_format.set_font_name(font)
            other_format.set_font_size(font_size)

            percentage_columns = ["L", "O", "R"]
            other_columns = list(
                filter(
                    lambda x: x not in percentage_columns, list(string.ascii_uppercase)
                )
            )
            last_column = "S"
            column_width = 12

            for col_num, col_name in enumerate(df.columns):
                worksheet.write(0, col_num, col_name, header_format)
                worksheet.set_column(col_num, col_num, column_width, header_format)
            for column in percentage_columns:
                worksheet.set_column(
                    f"{column}:{column}", column_width, percentage_format
                )
            for column in other_columns[
                other_columns.index("A") : other_columns.index("A") + 1
            ]:
                worksheet.set_column(f"{column}:{column}", column_width, date_format)
            for column in other_columns[
                other_columns.index("B") : other_columns.index(last_column) + 1
            ]:
                worksheet.set_column(f"{column}:{column}", column_width, other_format)
        try:
            subprocess.run(["xdg-open", temp_file.name], check=True)
        except Exception as e:
            print(f"Failed to open the file: {e}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--crash-stats-date",
        type=str,
        default=None,
        help="Optional date to use for crash stats query in YYYY-MM-DD format",
    )
    args = parser.parse_args()

    release_readiness_metrics = {
        "regression-all": None,
        "regression-severe": None,
        "non-regression-all": None,
        "non-regression-severe": None,
        "topcrash": None,
        "perf": None,
        "sec-crit-high": None,
        "sec-moderate-low": None,
        "daily-installations": None,
        "daily-crashes": None,
        "daily-crash-rate": None,
        "beta-installations": None,
        "beta-crashes": None,
        "beta-crash-rate": None,
        "release-installations": None,
        "release-crashes": None,
        "release-crash-rate": None,
        "total-installations": None,
    }

    for query_type in BMO_QUERY_TYPES:
        count = bmo_query(query_type)
        release_readiness_metrics[query_type] = count

    for query_type in STN_QUERY_TYPES:
        count = stn_query(query_type)
        release_readiness_metrics[query_type] = count

    for query_type in CSMO_QUERY_TYPES:
        count = csmo_query(query_type, args.crash_stats_date)
        release_readiness_metrics[query_type] = count

    release_readiness_metrics["daily-crash-rate"] = (
        release_readiness_metrics["daily-crashes"]
        / release_readiness_metrics["daily-installations"]
    )
    release_readiness_metrics["beta-crash-rate"] = (
        release_readiness_metrics["beta-crashes"]
        / release_readiness_metrics["beta-installations"]
    )
    release_readiness_metrics["release-crash-rate"] = (
        release_readiness_metrics["release-crashes"]
        / release_readiness_metrics["release-installations"]
    )

    print_versions()
    export_metrics_to_spreadsheet(release_readiness_metrics)


if __name__ == "__main__":
    main()
