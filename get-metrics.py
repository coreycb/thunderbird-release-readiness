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


def get_bmo_url(query_type, rest_url=True):
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

    if rest_url:
        url_base = (
            "https://bugzilla.mozilla.org/rest/bug?include_fields=id,summary,status&"
        )
    else:
        url_base = "https://bugzilla.mozilla.org/buglist.cgi?"

    url = (
        f"{url_base}"
        "bug_type=defect&"
        "chfield=%5BBug%20creation%5D&"
        "f1=short_desc&"
        "f2=component&"
        "f3=OP&"
        f"{f_version}"
        "j3=OR&"
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


def get_csmo_url(query_type, rest_url=True):
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

    if rest_url:
        url_base = "https://crash-stats.mozilla.org/api/SuperSearch/?product=Thunderbird&"
    else:
        url_base = "https://crash-stats.mozilla.org/search/?product=Thunderbird&"

    url = (
        f"{url_base}"
        f"{versions}"
        f"{start_date}"
        f"{end_date}"
        "_facets=platform&"
    )

    if rest_url:
        url = (
            f"{url}"
            "_facets=release_channel"
        )
    else:
        url = (
            f"{url}"
            "_facets=release_channel&"
            "_sort=-date&"
            "_columns=date&"
            "_columns=signature&"
            "_columns=product&"
            "_columns=version&"
            "_columns=build_id&"
            "_columns=platform#facet-release_channel"
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


def csmo_query(query_type):
    """Query crash-stats.mozilla.org"""
    r = requests.get(get_csmo_url(query_type))
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
        metrics_df = pd.DataFrame(
            [list(metrics["count"] for metrics in release_readiness_metrics.values())],
            columns=list(release_readiness_metrics.keys()),
        )
        metrics_df.insert(0, "Date", [get_today()])
        url_df = pd.DataFrame(
            [
                (metrics["text"], metrics["url"])
                for key, metrics in release_readiness_metrics.items()
                if "url" in metrics
            ],
            columns=["Description", "URL"],
        )
        with pd.ExcelWriter(temp_file.name, engine="xlsxwriter") as writer:

            workbook = writer.book

            font = "Arial"
            font_size = 10
            header_format = workbook.add_format(
                {"bg_color": "#B0B3B2", "align": "center"}
            )
            header_format.set_font_name(font)
            header_format.set_font_size(font_size)
            link_format = workbook.add_format({"font_color": "#0000FF", "underline": 1})
            link_format.set_font_name(font)
            link_format.set_font_size(font_size)
            column_width = 50

            sheet1 = workbook.add_worksheet("Release Metrics Charts")
            sheet1.write(0, 0, "Query URLs", header_format)
            for row_num, (
                description,
                url,
            ) in enumerate(url_df.itertuples(index=False), start=1):
                sheet1.write_url(row_num, 0, url, link_format, description)
            sheet1.set_column(0, 0, column_width)

            font = "Helvetica Neue"

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

            metrics_df.to_excel(writer, index=False, sheet_name="Data from Queries")
            sheet2 = writer.sheets["Data from Queries"]
            for col_num, col_name in enumerate(metrics_df.columns):
                sheet2.write(0, col_num, col_name, header_format)
                sheet2.set_column(col_num, col_num, column_width, header_format)
            for column in percentage_columns:
                sheet2.set_column(f"{column}:{column}", column_width, percentage_format)
            for column in other_columns[
                other_columns.index("A") : other_columns.index("A") + 1
            ]:
                sheet2.set_column(f"{column}:{column}", column_width, date_format)
            for column in other_columns[
                other_columns.index("B") : other_columns.index(last_column) + 1
            ]:
                sheet2.set_column(f"{column}:{column}", column_width, other_format)

        try:
            subprocess.run(["xdg-open", temp_file.name], check=True)
        except Exception as e:
            print(f"Failed to open the file: {e}")


def main():
    def create_metrics_dict(keys_with_texts):
        return {key: {"text": text} for key, text in keys_with_texts}

    release_readiness_metrics = create_metrics_dict(
        [
            ("regression-all", "# of regressions (affecting 128+)"),
            ("regression-severe", "# of severe (S1/S2) regressions (affecting 128+)"),
            ("non-regression-all", "# of non-regressions (affecting 128+)"),
            ("non-regression-severe", "# of severe (S1/S2) non-regressions (affecting 128+)"),
            ("topcrash", "# of topcrash bugs (affecting 128+)"),
            ("perf", "# of perf bugs (affecting 128+)"),
            ("sec-crit-high", "# of sec-crit, sec-high bugs (affecting 128+)"),
            ("sec-moderate-low", "# of sec-moderate, sec-low (affecting 128+)"),
            ("daily-installations", None),
            ("daily-crashes", "Daily crashes (last 24 hours)"),
            ("daily-crash-rate", None),
            ("beta-installations", None),
            ("beta-crashes", "Beta crashes (last 24 hours)"),
            ("beta-crash-rate", None),
            ("release-installations", None),
            ("release-crashes", "Release crashes (last 24 hours)"),
            ("release-crash-rate", None),
            ("total-installations", None),
        ]
    )

    print_versions()

    for query_type in BMO_QUERY_TYPES:
        count = bmo_query(query_type)
        release_readiness_metrics[query_type]["count"] = count

    for query_type in STN_QUERY_TYPES:
        count = stn_query(query_type)
        release_readiness_metrics[query_type]["count"] = count

    for query_type in CSMO_QUERY_TYPES:
        count = csmo_query(query_type)
        release_readiness_metrics[query_type]["count"] = count

    release_readiness_metrics["daily-crash-rate"]["count"] = (
        release_readiness_metrics["daily-crashes"]["count"]
        / release_readiness_metrics["daily-installations"]["count"]
    )
    release_readiness_metrics["beta-crash-rate"]["count"] = (
        release_readiness_metrics["beta-crashes"]["count"]
        / release_readiness_metrics["beta-installations"]["count"]
    )
    release_readiness_metrics["release-crash-rate"]["count"] = (
        release_readiness_metrics["release-crashes"]["count"]
        / release_readiness_metrics["release-installations"]["count"]
    )

    for query_type in BMO_QUERY_TYPES:
        release_readiness_metrics[query_type]["url"] = get_bmo_url(
            query_type, rest_url=False
        )

    for query_type in CSMO_QUERY_TYPES:
        release_readiness_metrics[query_type]["url"] = get_csmo_url(
            query_type, rest_url=False
        )


if __name__ == "__main__":
    main()
