#!/usr/bin/env python3

import argparse
import cloudscraper
import json
import logging
import os
import re
import requests
import string
import subprocess
import sys
import tempfile

from datetime import datetime, timedelta
from functools import lru_cache

import pandas as pd

from tabulate import tabulate


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
    "dataloss",
]

# stats.thunderbird.net
STN_QUERY_TYPES = [
    "daily-adi",
    "beta-adi",
    "release-adi",
    "total-adi",
    "esr140-adi",
]

# crash-stats.mozilla.org
CSMO_QUERY_TYPES = [
    "release-crashes",
    "beta-crashes",
    "daily-crashes",
    "esr140-crashes",
]

INCLUDE_PREVIOUS_DAILIES = 0
INCLUDE_PREVIOUS_BETA = False
INCLUDE_PREVIOUS_RELEASES = 0
ESR_NEXT = False
INCLUDE_115 = False

scraper = cloudscraper.create_scraper()


@lru_cache
def current_thunderbird_versions():
    url = "https://product-details.mozilla.org/1.0/thunderbird_versions.json"
    response = scraper.get(url)
    data = response.json()
    return data


@lru_cache
def thunderbird_esr_major_version():
    thunderbird_versions = current_thunderbird_versions()
    version = thunderbird_versions["THUNDERBIRD_ESR"].split(".")[0]
    if ESR_NEXT:
        version = thunderbird_versions["THUNDERBIRD_ESR_NEXT"].split(".")[0]
    return version


@lru_cache
def thunderbird_status_versions():
    thunderbird_versions = current_thunderbird_versions()
    status_version_start = thunderbird_esr_major_version()
    status_version_end = thunderbird_versions[
        "LATEST_THUNDERBIRD_NIGHTLY_VERSION"
    ].split(".")[0]
    thunderbird_status_versions = [
        f"cf_status_thunderbird_esr{status_version_start}"
    ] + [
        f"cf_status_thunderbird_{index}"
        for index in range(int(status_version_start), int(status_version_end) + 1)
    ]
    return thunderbird_status_versions


@lru_cache
def thunderbird_daily_versions():
    thunderbird_versions = current_thunderbird_versions()
    daily = thunderbird_versions["LATEST_THUNDERBIRD_NIGHTLY_VERSION"]
    thunderbird_daily_versions = [daily]
    if INCLUDE_PREVIOUS_DAILIES > 0:
        # Include the specified number of previous dailies
        for i in range(1, INCLUDE_PREVIOUS_DAILIES + 1):
            previous_daily = f"{int(daily.split('.')[0]) - i}.0a1"
            thunderbird_daily_versions.append(previous_daily)
    return thunderbird_daily_versions


@lru_cache
def thunderbird_beta_versions():
    thunderbird_versions = current_thunderbird_versions()
    beta = thunderbird_versions["LATEST_THUNDERBIRD_DEVEL_VERSION"].split(".")[0]
    thunderbird_beta_versions = [f"{beta}.0b{index}" for index in range(1, 7)]
    if INCLUDE_PREVIOUS_BETA:
        previous_beta = f"{int(beta.split('.')[0]) - 1}"
        thunderbird_beta_versions = thunderbird_beta_versions + [
            f"{previous_beta}.0b{index}" for index in range(1, 7)
        ]
    return thunderbird_beta_versions


@lru_cache
def thunderbird_release_versions():
    thunderbird_versions = current_thunderbird_versions()
    release = thunderbird_versions["LATEST_THUNDERBIRD_VERSION"].split(".")[0]
    thunderbird_release_versions = [f"{release}.0"] + [
        f"{release}.0.{index}" for index in range(1, 4)
    ]
    if INCLUDE_PREVIOUS_RELEASES > 0:
        # Include the specified number of previous releases
        for i in range(1, INCLUDE_PREVIOUS_RELEASES + 1):
            previous_release = int(release) - i
            thunderbird_release_versions.extend([
                f"{previous_release}.0",
                f"{previous_release}.0.1",
                f"{previous_release}.0.2",
                f"{previous_release}.0.3"
            ])
    return thunderbird_release_versions

@lru_cache
def thunderbird_current_daily_version():
    """Get only the current daily version (no previous versions)"""
    thunderbird_versions = current_thunderbird_versions()
    daily = thunderbird_versions["LATEST_THUNDERBIRD_NIGHTLY_VERSION"]
    return [daily]


@lru_cache
def thunderbird_current_beta_versions():
    """Get only the current beta versions (no previous versions)"""
    thunderbird_versions = current_thunderbird_versions()
    beta = thunderbird_versions["LATEST_THUNDERBIRD_DEVEL_VERSION"].split(".")[0]
    return [f"{beta}.0b{index}" for index in range(1, 7)]


@lru_cache
def thunderbird_current_release_versions():
    """Get only the current release versions (no previous versions)"""
    thunderbird_versions = current_thunderbird_versions()
    release = thunderbird_versions["LATEST_THUNDERBIRD_VERSION"].split(".")[0]
    return [f"{release}.0"] + [f"{release}.0.{index}" for index in range(1, 4)]

@lru_cache
def thunderbird_current_esr140_versions():
    """Get only the current ESR 140 minor version (latest major.minor.x)"""
    all_esr140_versions = thunderbird_esr_versions("140")
    if not all_esr140_versions:
        return []

    # Group versions by minor version (major.minor)
    minor_versions = {}
    for version in all_esr140_versions:
        parts = version.split(".")
        if len(parts) >= 2:
            minor_key = f"{parts[0]}.{parts[1]}"
            if minor_key not in minor_versions:
                minor_versions[minor_key] = []
            minor_versions[minor_key].append(version)

    # Get the latest minor version
    if minor_versions:
        latest_minor = max(minor_versions.keys(), key=lambda x: tuple(map(int, x.split("."))))
        return sorted(minor_versions[latest_minor])

    return []

@lru_cache
def thunderbird_esr_versions(major_version):
    """Get all ESR versions for a given major version from thunderbird_adi.json"""
    r = scraper.get("https://stats.thunderbird.net/thunderbird_adi.json")
    data = json.loads(r.text)

    latest_date = max(data.keys())
    versions_data = data[latest_date].get("versions", {})

    esr_versions = []
    major_str = str(major_version)

    for version in versions_data.keys():
        if version.startswith(f"{major_str}."):
            # Ensure it's an ESR-like version (x.y.z format)
            parts = version.split(".")
            if len(parts) >= 2 and parts[0] == major_str:
                # Filter out alpha/beta versions (those containing 'a' or 'b')
                if 'a' not in version and 'b' not in version:
                    esr_versions.append(version)

    return sorted(esr_versions)


@lru_cache
def thunderbird_esr_count(major_version):
    """Get count of users for all ESR versions of a given major version

    Special case: When major_version is "115", includes all ESR versions <= 115
    """
    r = scraper.get("https://stats.thunderbird.net/thunderbird_adi.json")
    data = json.loads(r.text)

    yesterday_data = data[yesterday()]
    versions_data = yesterday_data.get("versions", {})

    total_count = 0

    if major_version == "115":
        # Special case: include all ESR versions <= 115
        target_major = int(major_version)

        for version in versions_data.keys():
            # Check if it's a stable release (no 'a' or 'b')
            if 'a' not in version and 'b' not in version:
                try:
                    version_parts = version.split(".")
                    if len(version_parts) >= 2:
                        version_major = int(version_parts[0])
                        # Include all versions with major version <= 115
                        if version_major <= target_major:
                            total_count += versions_data[version]
                except (ValueError, IndexError):
                    # Skip versions that don't parse correctly
                    continue
    else:
        # Normal case: only include versions matching the exact major version
        esr_versions = thunderbird_esr_versions(major_version)

        for version in esr_versions:
            if version in versions_data:
                total_count += versions_data[version]

    return total_count


def stn_current_query(query_type):
    """Query stats.thunderbird.net for current versions only"""
    r = scraper.get("https://stats.thunderbird.net/thunderbird_adi.json")
    count = 0

    if query_type == "current-daily-adi":
        for version in thunderbird_current_daily_version():
            if version in json.loads(r.text)[yesterday()]["versions"]:
                count += json.loads(r.text)[yesterday()]["versions"][version]
    elif query_type == "current-beta-adi":
        for version in thunderbird_current_beta_versions():
            if "0b1" not in version:
                continue
            version = thunderbird_current_beta_versions()[0].split("b")[0]
            if version in json.loads(r.text)[yesterday()]["versions"]:
                count += json.loads(r.text)[yesterday()]["versions"][version]
    elif query_type == "current-release-adi":
        for version in thunderbird_current_release_versions():
            if version in json.loads(r.text)[yesterday()]["versions"]:
                count += json.loads(r.text)[yesterday()]["versions"][version]

    return count


def csmo_current_query(query_type):
    """Query crash-stats.mozilla.org for current versions only"""
    versions = ""

    if query_type == "current-daily-crashes":
        for version in thunderbird_current_daily_version():
            versions = f"{versions}version={version}&"
    elif query_type == "current-beta-crashes":
        for version in thunderbird_current_beta_versions():
            versions = f"{versions}version={version}&"
    elif query_type == "current-release-crashes":
        for version in thunderbird_current_release_versions():
            versions = f"{versions}version={version}&"

    today_formatted = today()
    yesterday_formatted = yesterday()
    start_date = f"date=>={yesterday_formatted}T00:00:00.000Z&"
    end_date = f"date=<{today_formatted}T00:00:00.000Z&"

    url_base = "https://crash-stats.mozilla.org/api/SuperSearch/?product=Thunderbird&"
    url = f"{url_base}{versions}{start_date}{end_date}_facets=platform&_facets=release_channel"
    r = scraper.get(url)
    count = json.loads(r.text)["total"]
    return count


def csmo_current_esr140_query():
    """Query crash-stats.mozilla.org for current ESR 140 minor version only"""
    versions = ""

    for version in thunderbird_current_esr140_versions():
        version_with_esr = f"{version}esr"
        versions = f"{versions}version={version_with_esr}&"

    today_formatted = today()
    yesterday_formatted = yesterday()
    start_date = f"date=>={yesterday_formatted}T00:00:00.000Z&"
    end_date = f"date=<{today_formatted}T00:00:00.000Z&"

    url_base = "https://crash-stats.mozilla.org/api/SuperSearch/?product=Thunderbird&"
    url = f"{url_base}{versions}{start_date}{end_date}_facets=platform&_facets=release_channel"

    r = scraper.get(url)
    count = json.loads(r.text)["total"]
    return count


@lru_cache
def today():
    """Get today's date in YYYY-MM-DD format"""
    today = datetime.now()
    return today.strftime("%Y-%m-%d")


@lru_cache
def yesterday():
    """Get yesterday's date in YYYY-MM-DD format"""
    today = datetime.now()
    yesterday = today - timedelta(days=1)
    return yesterday.strftime("%Y-%m-%d")


def bmo_url(query_type, rest_url=True):
    """Get bugzilla.mozilla.org URL"""
    index = 4
    f_version = ""
    o_comparison = ""
    v_status = ""
    for version in thunderbird_status_versions():
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
        case "dataloss":
            url = (
                f"{url}"
                "keywords=dataloss&"
                "keywords_type=allwords&"
                "o1=nowordssubstr&"
                "o2=nowordssubstr&"
            )
        case _:
            sys.exit(f"Unknown query type: {query_type}")

    return url


def csmo_url(query_type, rest_url=True):
    """Get crash-stats.mozilla.org URL"""
    versions = ""
    match query_type:
        case "daily-crashes":
            for version in thunderbird_current_daily_version():
                versions = f"{versions}version={version}&"
        case "beta-crashes":
            for version in thunderbird_current_beta_versions():
                versions = f"{versions}version={version}&"
        case "release-crashes":
            for version in thunderbird_current_release_versions():
                versions = f"{versions}version={version}&"
        case "esr140-crashes":
            for version in thunderbird_current_esr140_versions():
                version = f"{version}esr"
                versions = f"{versions}version={version}&"

    today_formatted = today()
    yesterday_formatted = yesterday()
    start_date = f"date=>={yesterday_formatted}T00:00:00.000Z&"
    end_date = f"date=<{today_formatted}T00:00:00.000Z&"

    if rest_url:
        url_base = (
            "https://crash-stats.mozilla.org/api/SuperSearch/?product=Thunderbird&"
        )
    else:
        url_base = "https://crash-stats.mozilla.org/search/?product=Thunderbird&"

    url = f"{url_base}" f"{versions}" f"{start_date}" f"{end_date}" "_facets=platform&"

    if rest_url:
        url = f"{url}" "_facets=release_channel"
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

    r = scraper.get(bmo_url(query_type), headers=headers, params=params)
    count = len(json.loads(r.text)["bugs"])
    return count


def stn_query(query_type):
    """Query stats.thunderbird.net"""
    r = scraper.get("https://stats.thunderbird.net/thunderbird_adi.json")
    match query_type:
        case "total-adi":
            count = json.loads(r.text)[yesterday()]["count"]
        case "daily-adi":
            count = 0
            for version in thunderbird_daily_versions():
                if version in json.loads(r.text)[yesterday()]["versions"]:
                    count += json.loads(r.text)[yesterday()]["versions"][version]
        case "beta-adi":
            count = 0
            for version in thunderbird_beta_versions():
                if "0b1" not in version:
                    continue
                version = thunderbird_beta_versions()[0].split("b")[0]
                if version in json.loads(r.text)[yesterday()]["versions"]:
                    count += json.loads(r.text)[yesterday()]["versions"][version]
        case "release-adi":
            count = 0
            for version in thunderbird_release_versions():
                if version in json.loads(r.text)[yesterday()]["versions"]:
                    count += json.loads(r.text)[yesterday()]["versions"][version]
        case "esr140-adi":
            count = thunderbird_esr_count("140")
        case "esr128-adi":
            count = thunderbird_esr_count("128")
        case "esr115-adi":
            count = thunderbird_esr_count("115")
        case _:
            sys.exit(f"Unknown query type: {query_type}")
    return count


def csmo_query(query_type):
    """Query crash-stats.mozilla.org"""
    r = scraper.get(csmo_url(query_type))
    count = json.loads(r.text)["total"]
    return count


def print_versions():
    affected_versions = re.sub(
        "cf_status_thunderbird_", "", ", ".join(thunderbird_status_versions())
    )
    daily_versions = f"{', '.join(thunderbird_daily_versions())}"
    beta_versions = f"{', '.join(thunderbird_beta_versions())}"
    release_versions = f"{', '.join(thunderbird_release_versions())}"
    esr140_versions = f"{', '.join(thunderbird_esr_versions(140))}"
    table_data = {
        "Category": [
            "bugzilla affected versions",
            "daily versions",
            "beta versions",
            "release versions",
            "esr140 versions"
        ],
        "Variable": [
            affected_versions,
            daily_versions,
            beta_versions,
            release_versions,
            esr140_versions,
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
        metrics_df.insert(0, "Date", [today()])
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

            # formatting for sheet 1
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

            # write sheet 1
            sheet1 = workbook.add_worksheet("Release Metrics Charts")
            sheet1.write(0, 0, "Query URLs", header_format)
            for row_num, (
                description,
                url,
            ) in enumerate(url_df.itertuples(index=False), start=1):
                sheet1.write_url(row_num, 0, url, link_format, description)
            sheet1.set_column(0, 0, column_width)

            # formatting for sheet 2
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

            percentage_columns = ["L", "O", "R", "T", "U", "V", "Y", "AA"]
            other_columns = list(
                filter(
                    lambda x: x not in percentage_columns, list(string.ascii_uppercase)
                )
            )
            last_other_column = "Z"
            column_width = 12

            # write sheet 2
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
                other_columns.index("B") : other_columns.index(last_other_column) + 1
            ]:
                sheet2.set_column(f"{column}:{column}", column_width, other_format)

        try:
            subprocess.run(["xdg-open", temp_file.name], check=True)
        except Exception as e:
            print(f"Failed to open the file: {e}")


def create_metrics_dict(keys_with_texts):
    return {key: {"text": text} for key, text in keys_with_texts}


def main():
    global INCLUDE_PREVIOUS_DAILIES
    global INCLUDE_PREVIOUS_BETA
    global INCLUDE_PREVIOUS_RELEASES
    global ESR_NEXT
    global INCLUDE_115

    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    logger = logging.getLogger(__name__)

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-id",
        "--include-previous-dailies",
        type=int,
        default=0,
        help="Include the specified number of previous Daily channel versions in addition to the current version",
    )
    parser.add_argument(
        "-ib",
        "--include-previous-beta",
        action="store_true",
        help="Include the previous Beta channel version in addition to the current version",
    )
    parser.add_argument(
        "-ir",
        "--include-previous-releases",
        type=int,
        default=0,
        help="Include the specified number of previous Release channel versions in addition to the current version",
    )
    parser.add_argument(
        "-en",
        "--esr-next",
        action="store_true",
        help="Gather data since THUNDERBIRD_ESR_NEXT instead of THUNDERBIRD_ESR",
    )
    parser.add_argument(
        "-in115",
        "--include-115",
        action="store_true",
        help="Include ESR 115 and older users in the total ADI count (excluded by default)",
    )
    args = parser.parse_args()
    INCLUDE_PREVIOUS_DAILIES = args.include_previous_dailies
    INCLUDE_PREVIOUS_BETA = args.include_previous_beta
    INCLUDE_PREVIOUS_RELEASES = args.include_previous_releases
    ESR_NEXT = args.esr_next
    INCLUDE_115 = args.include_115

    esr_major_version = thunderbird_esr_major_version()
    # fmt: off
    release_readiness_metrics = create_metrics_dict([
        ("regression-all", f"# of regressions (affecting {esr_major_version}+)"),
        ("regression-severe", f"# of severe (S1/S2) regressions (affecting {esr_major_version}+)"),
        ("non-regression-all", f"# of non-regressions (affecting {esr_major_version}+)"),
        ("non-regression-severe", f"# of severe (S1/S2) non-regressions (affecting {esr_major_version}+)"),
        ("topcrash", f"# of topcrash bugs (affecting {esr_major_version}+)"),
        ("perf", f"# of perf bugs (affecting {esr_major_version}+)"),
        ("sec-crit-high", f"# of sec-crit, sec-high bugs (affecting {esr_major_version}+)"),
        ("sec-moderate-low", f"# of sec-moderate, sec-low (affecting {esr_major_version}+)"),
        ("daily-adi", None),
        ("daily-crashes", "Daily crashes (last 24 hours)"),
        ("daily-crash-rate", None),
        ("beta-adi", None),
        ("beta-crashes", "Beta crashes (last 24 hours)"),
        ("beta-crash-rate", None),
        ("release-adi", None),
        ("release-crashes", "Release crashes (last 24 hours)"),
        ("release-crash-rate", None),
        ("total-adi", None),
        ("daily-adi-%", None),
        ("beta-adi-%", None),
        ("release-adi-%", None),
        ("dataloss", f"# of dataloss bugs (affecting {esr_major_version}+)"),
        ("esr140-adi", None),
        ("esr140-adi-%", None),
        ("esr140-crashes", "ESR 140 crashes (last 24 hours)"),
        ("esr140-crash-rate", None),
    ])
    # fmt: on

    print_versions()
    logger = logging.getLogger(__name__)

    logger.info("\n\nca=== Starting BMO (Bugzilla) Queries ===")
    for query_type in BMO_QUERY_TYPES:
        logger.info(f"Querying BMO for {query_type}...")
        count = bmo_query(query_type)
        release_readiness_metrics[query_type]["count"] = count
        logger.info(f"BMO {query_type}: {count} bugs found")
        logger.info(f"BMO {query_type} URL: {bmo_url(query_type, rest_url=False)}")

    logger.info("\n\n=== Starting STN (Stats) Queries ===")
    for query_type in STN_QUERY_TYPES:
        logger.info(f"Querying STN for {query_type}...")
        count = stn_query(query_type)
        release_readiness_metrics[query_type]["count"] = count

        if query_type == "daily-adi":
            daily_versions = thunderbird_daily_versions()
            logger.info(f"STN {query_type}: {count:,} users across {len(daily_versions)} daily versions: {daily_versions}")
        elif query_type == "beta-adi":
            beta_versions = thunderbird_beta_versions()
            logger.info(f"STN {query_type}: {count:,} users across {len(beta_versions)} beta versions: {beta_versions}")
        elif query_type == "release-adi":
            release_versions = thunderbird_release_versions()
            logger.info(f"STN {query_type}: {count:,} users across {len(release_versions)} release versions: {release_versions}")
        elif query_type == "esr140-adi":
            esr140_versions = thunderbird_esr_versions("140")
            logger.info(f"STN {query_type}: {count:,} users across {len(esr140_versions)} ESR 140 versions: {esr140_versions}")
        else:
            logger.info(f"STN {query_type}: {count:,} users")

    logger.info("\n\n=== Starting CSMO (Crash Stats) Queries ===")
    for query_type in CSMO_QUERY_TYPES:
        logger.info(f"Querying CSMO for {query_type}...")
        count = csmo_query(query_type)
        release_readiness_metrics[query_type]["count"] = count

        if query_type == "esr140-crashes":
            esr140_versions = thunderbird_current_esr140_versions()
            logger.info(f"CSMO {query_type}: {count} crashes from current ESR 140 versions: {esr140_versions}")
        else:
            logger.info(f"CSMO {query_type}: {count} crashes")
        logger.info(f"CSMO {query_type} URL: {csmo_url(query_type, rest_url=False)}")

    # Exclude ESR 115 and older users from total ADI by default
    logger.info("\n\n=== ESR 115 Processing ===")
    if not INCLUDE_115:
        esr_115_count = thunderbird_esr_count("115")
        original_total = release_readiness_metrics["total-adi"]["count"]
        release_readiness_metrics["total-adi"]["count"] -= esr_115_count
        logger.info(f"ESR 115 and older users: {esr_115_count:,} (excluded from total)")
        logger.info(f"Total ADI: {original_total:,} -> {release_readiness_metrics['total-adi']['count']:,} (after excluding ESR 115)")
    else:
        esr_115_count = thunderbird_esr_count("115")
        logger.info(f"ESR 115 and older users: {esr_115_count:,} (included in total)")
        logger.info(f"Total ADI: {release_readiness_metrics['total-adi']['count']:,}")

    # Calculate crash rates based on current versions only
    logger.info("\n\n=== Current Version Crash Rate Calculations ===")

    # Daily crash rate
    current_daily_versions = thunderbird_current_daily_version()
    current_daily_adi = stn_current_query("current-daily-adi")
    current_daily_crashes = csmo_current_query("current-daily-crashes")
    daily_crash_rate = current_daily_crashes / current_daily_adi if current_daily_adi > 0 else 0
    logger.info(f"Daily versions: {current_daily_versions}")
    logger.info(f"Daily ADI (current): {current_daily_adi:,}")
    logger.info(f"Daily crashes (current): {current_daily_crashes:,}")
    logger.info(f"Daily crash rate (current): {daily_crash_rate:.6f} ({daily_crash_rate*100:.4f}%)")

    # Beta crash rate
    current_beta_versions = thunderbird_current_beta_versions()
    current_beta_adi = stn_current_query("current-beta-adi")
    current_beta_crashes = csmo_current_query("current-beta-crashes")
    beta_crash_rate = current_beta_crashes / current_beta_adi if current_beta_adi > 0 else 0
    logger.info(f"Beta versions: {current_beta_versions}")
    logger.info(f"Beta ADI (current): {current_beta_adi:,}")
    logger.info(f"Beta crashes (current): {current_beta_crashes:,}")
    logger.info(f"Beta crash rate (current): {beta_crash_rate:.6f} ({beta_crash_rate*100:.4f}%)")

    # Release crash rate
    current_release_versions = thunderbird_current_release_versions()
    current_release_adi = stn_current_query("current-release-adi")
    current_release_crashes = csmo_current_query("current-release-crashes")
    release_crash_rate = current_release_crashes / current_release_adi if current_release_adi > 0 else 0
    logger.info(f"Release versions: {current_release_versions}")
    logger.info(f"Release ADI (current): {current_release_adi:,}")
    logger.info(f"Release crashes (current): {current_release_crashes:,}")
    logger.info(f"Release crash rate (current): {release_crash_rate:.6f} ({release_crash_rate*100:.4f}%)")

    # Update crash rates to use current version data
    release_readiness_metrics["daily-crash-rate"]["count"] = daily_crash_rate
    release_readiness_metrics["beta-crash-rate"]["count"] = beta_crash_rate
    release_readiness_metrics["release-crash-rate"]["count"] = release_crash_rate

    # Keep the original ADI percentages calculation
    logger.info("\n\n=== ADI Percentage Calculations ===")
    total_adi = release_readiness_metrics["total-adi"]["count"]
    for channel in ["daily", "beta", "release"]:
        channel_adi = release_readiness_metrics[f"{channel}-adi"]["count"]
        percentage = channel_adi / total_adi
        release_readiness_metrics[f"{channel}-adi-%"]["count"] = percentage
        logger.info(f"{channel.capitalize()} ADI: {channel_adi:,} / {total_adi:,} = {percentage:.6f} ({percentage*100:.4f}%)")

    logger.info("\n\n=== ESR Crash Rate and Percentage Calculations ===")
    for esr_version in ["140"]:
        # Use current minor version crashes but full ESR ADI count
        current_esr_versions = thunderbird_current_esr140_versions()
        all_esr_versions = thunderbird_esr_versions(esr_version)
        current_esr140_crashes = csmo_current_esr140_query()
        esr_adi = release_readiness_metrics[f"esr{esr_version}-adi"]["count"]

        esr_crash_rate = (
            current_esr140_crashes / esr_adi
            if esr_adi > 0 else 0
        )
        esr_percentage = esr_adi / total_adi

        logger.info(f"ESR {esr_version} all versions: {all_esr_versions}")
        logger.info(f"ESR {esr_version} current minor versions: {current_esr_versions}")
        logger.info(f"ESR {esr_version} ADI (all versions): {esr_adi:,}")
        logger.info(f"ESR {esr_version} ADI percentage: {esr_percentage:.6f} ({esr_percentage*100:.4f}%)")
        logger.info(f"ESR {esr_version} crashes (current minor only): {current_esr140_crashes:,}")
        logger.info(f"ESR {esr_version} crash rate (current minor only): {esr_crash_rate:.6f} ({esr_crash_rate*100:.4f}%)")

        release_readiness_metrics[f"esr{esr_version}-crash-rate"]["count"] = esr_crash_rate
        release_readiness_metrics[f"esr{esr_version}-adi-%"]["count"] = esr_percentage

    for query_type in BMO_QUERY_TYPES:
        release_readiness_metrics[query_type]["url"] = bmo_url(
            query_type, rest_url=False
        )

    for query_type in CSMO_QUERY_TYPES:
        release_readiness_metrics[query_type]["url"] = csmo_url(
            query_type, rest_url=False
        )

    logger.info("\n\n=== Summary ===")
    logger.info(f"Total metrics collected: {len(release_readiness_metrics)}")
    logger.info("Exporting metrics to spreadsheet...")
    export_metrics_to_spreadsheet(release_readiness_metrics)
    logger.info("Metrics collection and export completed successfully!")


if __name__ == "__main__":
    main()
