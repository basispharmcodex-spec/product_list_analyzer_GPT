#!/usr/bin/env python3
"""Run a read-only smoke test against a product list and the active GRLS archive."""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path


PLUGIN_ROOT = Path(__file__).resolve().parents[1]
SERVER_PATH = PLUGIN_ROOT / "mcp-server" / "server.py"

spec = importlib.util.spec_from_file_location("grls_mcp_server", SERVER_PATH)
server = importlib.util.module_from_spec(spec)
assert spec and spec.loader
spec.loader.exec_module(server)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--product-list", required=True)
    parser.add_argument("--grls")
    args = parser.parse_args()

    product = server.inspect_product_list(
        {"file_path": args.product_list, "start_row": 1, "max_rows": 12, "max_columns": 10}
    )
    archive_arguments = {"archive_path": args.grls} if args.grls else {}
    archive = server.inspect_grls_archive(archive_arguments)
    distinct = server.distinct_grls_values(
        {
            **archive_arguments,
            "field": "inn",
            "contains_terms": ["ацикловир"],
            "max_values": 20,
        }
    )

    result = {
        "product_list": {
            "format": product["format"],
            "reader": product.get("reader"),
            "sheets": product["sheets"],
            "selected_sheet": product["selected_sheet"],
            "last_row": product["last_row"],
            "last_column": product["last_column"],
            "preview_rows": len(product["rows"]),
        },
        "grls": {
            "status": server.get_grls_status({}),
            "sections": len(archive["sections"]),
            "statuses": [section["status"] for section in archive["sections"]],
        },
        "literal_candidate_check": distinct,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
