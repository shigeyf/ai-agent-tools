# cli.py
# -*- coding: utf-8 -*-

"""
CLI for filling JSON data into PowerPoint slides.
"""

import json
import os
import sys
import argparse
from pptx_fill_data_into_template import pptx_fill_data_from_json

# ---------------------------
# CLI
# ---------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Fill JSON data into placeholders in PowerPoint slides."
    )
    parser.add_argument(
        "-i",
        "--input",
        dest="input_json",
        help="Path to input JSON file",
        required=True,
    )

    args = parser.parse_args()
    json_path = args.input_json

    if not os.path.exists(json_path):
        print(f"[ERROR] JSON file not found: {json_path}")
        sys.exit(1)
    with open(json_path, "r", encoding="utf-8") as f:
        payload = json.load(f)

    print(f"[INFO] Loaded JSON data from {json_path}")

    try:
        pptx_fill_data_from_json(payload)
        print("[INFO] JSON data filling completed.")
    # pylint: disable=broad-except
    except Exception as e:
        print(f"[ERROR] Error at filling JSON data due to: {e}")
        sys.exit(1)
