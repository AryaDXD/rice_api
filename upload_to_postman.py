#!/usr/bin/env python3
"""Upload Student Management API collection to Postman Cloud Workspace via Postman REST API."""

import argparse
import json
import os
import sys
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent
COLLECTION_PATH = ROOT / "postman" / "Student-Management-API.postman_collection.json"
POSTMAN_RESOURCES = ROOT / ".postman" / "resources.yaml"
DEFAULT_WORKSPACE = "409b9255-4de7-47a9-8042-112bd0ef4aab"


def get_workspace_id():
    if POSTMAN_RESOURCES.exists():
        for line in POSTMAN_RESOURCES.read_text(encoding="utf-8").splitlines():
            if "id:" in line and not line.strip().startswith("#"):
                return line.split(":", 1)[1].strip()
    return DEFAULT_WORKSPACE


def main():
    parser = argparse.ArgumentParser(description="Upload Student Management collection to Postman Cloud.")
    parser.add_argument("--key", default=os.environ.get("POSTMAN_API_KEY", ""), help="Postman API Key (PMAK-...)")
    parser.add_argument("--workspace", default=get_workspace_id(), help="Target Postman Workspace ID")
    parser.add_argument("--file", default=str(COLLECTION_PATH), help="Path to collection JSON file")
    args = parser.parse_args()

    api_key = args.key.strip()
    if not api_key:
        print("=================================================================")
        print("  Postman Cloud API Key Required")
        print("=================================================================")
        print("To upload directly to your Postman Cloud workspace, you need a")
        print("Postman API Key (starts with 'PMAK-').")
        print()
        print("You can generate one in 10 seconds at:")
        print("  https://web.postman.co/settings/me/api-keys")
        print()
        api_key = input("Enter your Postman API Key (or press Enter to cancel): ").strip()
        if not api_key:
            print("\nUpload canceled. You can also import the collection manually:")
            print(f"  File -> Import in Postman -> select '{COLLECTION_PATH}'")
            sys.exit(0)

    col_file = Path(args.file)
    if not col_file.exists():
        print(f"Error: Collection file not found at {col_file}")
        sys.exit(1)

    col_json = json.loads(col_file.read_text(encoding="utf-8"))
    col_name = col_json.get("info", {}).get("name", "Student Management API")

    headers = {
        "X-Api-Key": api_key,
        "Content-Type": "application/json",
        "User-Agent": "PostmanUploader/1.0",
    }

    print(f"Connecting to Postman Cloud (Workspace: {args.workspace})...")

    # 1. Check existing collections in workspace
    existing_uid = None
    try:
        req_list = Request(
            f"https://api.getpostman.com/collections?workspace={args.workspace}",
            headers=headers,
            method="GET",
        )
        with urlopen(req_list, timeout=20) as resp:
            data = json.loads(resp.read().decode())
            for col in data.get("collections", []):
                if col.get("name") == col_name:
                    existing_uid = col.get("uid")
                    break
    except Exception as exc:
        pass  # Proceed to attempt create if list query fails

    # 2. Upload (create or update)
    if existing_uid:
        print(f"Found existing collection '{col_name}' (UID: {existing_uid}). Updating...")
        url = f"https://api.getpostman.com/collections/{existing_uid}"
        method = "PUT"
    else:
        print(f"Creating new collection '{col_name}' in workspace...")
        url = f"https://api.getpostman.com/collections?workspace={args.workspace}"
        method = "POST"

    payload = json.dumps({"collection": col_json}).encode("utf-8")
    req = Request(url, data=payload, headers=headers, method=method)

    try:
        with urlopen(req, timeout=30) as resp:
            res_data = json.loads(resp.read().decode())
            res_col = res_data.get("collection", {})
            print()
            print("=================================================================")
            print("  SUCCESSFULLY UPLOADED TO POSTMAN!")
            print("=================================================================")
            print(f"  Collection Name : {res_col.get('name')}")
            print(f"  Collection ID   : {res_col.get('id')}")
            print(f"  Collection UID  : {res_col.get('uid')}")
            print(f"  Workspace ID    : {args.workspace}")
            print()
            print("Open Postman to view your updated workspace!")
    except HTTPError as err:
        err_body = err.read().decode("utf-8", errors="replace")
        print()
        print(f"[ERROR] Postman API responded with HTTP {err.code}:")
        try:
            err_json = json.loads(err_body)
            print(f"  {err_json.get('error', {}).get('message', err_body)}")
        except ValueError:
            print(f"  {err_body}")
        sys.exit(1)
    except URLError as exc:
        print(f"[ERROR] Network connection failed: {exc.reason}")
        sys.exit(1)


if __name__ == "__main__":
    main()

