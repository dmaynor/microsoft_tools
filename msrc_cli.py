#!/usr/bin/env python3
"""
CLI Tool: MSRC Security Updates API Client
Fetch and display CVE details and enumerate updates from Microsoft Security Response Center (MSRC).
Aligned with official MSRC GitHub examples. Includes logging capabilities.
"""
import requests
import argparse
import json
import sys
import csv

MSRC_API_BASE = "https://api.msrc.microsoft.com/cvrf/v3.0/updates"

headers = {
    'Accept': 'application/json',
    'User-Agent': 'MSRC-API-Client/1.0'
}

def fetch_cve_details(cve_id: str) -> dict:
    url = f"{MSRC_API_BASE}/{cve_id}"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

def enumerate_monthly_updates(month: str) -> dict:
    url = f"{MSRC_API_BASE}/{month}"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

def fetch_monthly_cvrf_details(cvrf_url: str) -> dict:
    response = requests.get(cvrf_url, headers=headers)
    response.raise_for_status()
    return response.json()

def log_to_file(data: dict, file_path: str, file_format: str):
    if file_format == "json":
        with open(file_path, 'w') as file:
            json.dump(data, file, indent=2)
    elif file_format == "csv":
        if 'Vulnerability' in data:
            keys = data['Vulnerability'][0].keys()
            with open(file_path, 'w', newline='') as file:
                writer = csv.DictWriter(file, fieldnames=keys)
                writer.writeheader()
                writer.writerows(data['Vulnerability'])
        else:
            raise ValueError("No 'Vulnerability' key in data for CSV export.")
    else:
        raise ValueError("Unsupported file format. Use 'json' or 'csv'.")

def main():
    parser = argparse.ArgumentParser(description="Fetch CVE details from MSRC")
    subparsers = parser.add_subparsers(dest="command", required=True)

    cve_parser = subparsers.add_parser("cve", help="Fetch details for a specific CVE")
    cve_parser.add_argument("cve_id", help="The CVE identifier (e.g., CVE-2025-24043)")

    enum_parser = subparsers.add_parser("enumerate", help="Enumerate CVEs for a specific month")
    enum_parser.add_argument("month", help="Month identifier in yyyy-mmm format (e.g., 2025-mar)")

    parser.add_argument("-o", "--output", help="File path to log output")
    parser.add_argument("-f", "--format", choices=["json", "csv"], default="json", help="Output file format (default: json)")

    args = parser.parse_args()

    try:
        if args.command == "cve":
            details = fetch_cve_details(args.cve_id)
        elif args.command == "enumerate":
            monthly_summary = enumerate_monthly_updates(args.month)
            cvrf_url = monthly_summary['value'][0]['CvrfUrl']
            details = fetch_monthly_cvrf_details(cvrf_url)

        print(json.dumps(details, indent=2))

        if args.output:
            log_to_file(details, args.output, args.format)

    except requests.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}", file=sys.stderr)
        sys.exit(1)
    except Exception as err:
        print(f"An unexpected error occurred: {err}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
