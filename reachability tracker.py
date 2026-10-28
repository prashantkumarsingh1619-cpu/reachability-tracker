#!/usr/bin/env python3
"""
Reachability Tracker CLI
Generates the Reachability Tracker workbook from the site's CIQ + IP-Plan Excel files.

Usage:
    python "reachability tracker.py" --input-dir "Data 2"
    python "reachability tracker.py" --input-dir "Data 1" --out "Reachability_tracker_of_Karnal_2.xlsx"
"""
import argparse
import glob
import os
import sys

import tracker_engine

FILE_PATTERNS = {
    "ciq": "*CIQ*.xlsx",
    "data": "*UP-DATA-IP-Plan*.xlsx",
    "ims": "*UP-IMS-IP-Plan*.xlsx"
}


def find_file(folder: str, pattern: str) -> str:
    matches = glob.glob(os.path.join(folder, pattern))
    if not matches:
        # Also check subdirectories or loose pattern
        matches = [f for f in glob.glob(os.path.join(folder, "**", pattern), recursive=True)
                   if not os.path.basename(f).startswith("~$")]
    if not matches:
        raise FileNotFoundError(f"No file matching pattern '{pattern}' found in '{folder}'")
    matches = [f for f in matches if not os.path.basename(f).startswith("~$")]
    return matches[0]


def detect_site(folder: str, filepaths: dict) -> str:
    combined_names = (folder + " " + " ".join(os.path.basename(p) for p in filepaths.values())).lower()
    if "mohali" in combined_names:
        return "Mohali-5"
    if "karnal" in combined_names:
        return "Karnal-2"
    return "Tracker-Site"


def main():
    parser = argparse.ArgumentParser(
        description="Build Reachability Tracker workbook from CIQ and IP-Plan files.",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "--input-dir", "-i",
        default=".",
        help="Directory containing the 3 source Excel files (*CIQ*, *UP-DATA-IP-Plan*, *UP-IMS-IP-Plan*). Default: ."
    )
    parser.add_argument(
        "--site", "-s",
        default=None,
        help="Site name (e.g. 'Mohali-5' or 'Karnal-2'). Auto-detected if omitted."
    )
    parser.add_argument(
        "--out", "-o",
        default=None,
        help="Output Excel file path. Auto-generated if omitted."
    )

    args = parser.parse_args()

    # If default "." doesn't have CIQ files, check Data 1 and Data 2
    input_dir = args.input_dir
    if input_dir == "." and not glob.glob("*CIQ*.xlsx"):
        if os.path.isdir("Data 1"):
            input_dir = "Data 1"
        elif os.path.isdir("Data 2"):
            input_dir = "Data 2"

    print(f"Scanning for input Excel files in: {os.path.abspath(input_dir)}")
    try:
        files = {k: find_file(input_dir, p) for k, p in FILE_PATTERNS.items()}
    except FileNotFoundError as err:
        print(f"ERROR: {err}", file=sys.stderr)
        sys.exit(1)

    for k, path in files.items():
        print(f"  [+] {k.upper():5s} : {os.path.basename(path)}")

    site_name = args.site or detect_site(input_dir, files)
    out_filename = args.out or f"Reachability_tracker_of_{site_name}.xlsx"

    print(f"Site name: {site_name}")
    print("Extracting parameters and building Reachability Tracker...")

    out_buf, stats, warnings_list, summary_preview, detail_preview = tracker_engine.generate_reachability_tracker(
        ciq_source=files["ciq"],
        data_source=files["data"],
        ims_source=files["ims"],
        site_name=site_name,
        filenames={k: os.path.basename(v) for k, v in files.items()}
    )

    with open(out_filename, "wb") as f:
        f.write(out_buf.getvalue())

    print(f"\nSuccessfully generated: {os.path.abspath(out_filename)}")
    print(f"  Total Blocks      : {stats['total_blocks']}")
    print(f"  Total Detail Rows : {stats['total_detail_rows']} (Rows {stats['start_row']}-{stats['end_row']})")
    print(f"  Nodes Included    : {', '.join(stats['nodes'])}")
    print(f"  File Size         : {stats['workbook_size_bytes'] / 1024:.1f} KB")

    if warnings_list:
        print("\nWarnings:")
        for w in warnings_list:
            print(f"  [!] {w}")
    else:
        print("\nExtraction completed with 0 warnings.")


if __name__ == "__main__":
    main()