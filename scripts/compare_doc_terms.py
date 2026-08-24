#!/usr/bin/env python3
"""对照招标原件与质疑函，报告质疑函独有的产品名和专有词。"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from scan_tender_text import build_inventory


def term_map(inventory: dict) -> dict[str, dict]:
    """按候选词建立位置索引，便于回填原文定位。"""
    return {item["term"]: item for item in inventory["candidate_terms"]}


def main() -> None:
    parser = argparse.ArgumentParser(description="核对质疑函引用的专有词是否出现在招标原件")
    parser.add_argument("tender", type=Path, help="招标原件 .doc 或 .docx")
    parser.add_argument("challenge", type=Path, help="质疑函 .docx")
    parser.add_argument("--out", type=Path, help="JSON输出路径")
    args = parser.parse_args()

    tender = term_map(build_inventory(args.tender))
    challenge = term_map(build_inventory(args.challenge))
    tender_terms = set(tender)
    challenge_terms = set(challenge)
    report = {
        "tender": str(args.tender),
        "challenge": str(args.challenge),
        "scan_complete": True,
        "draft_only_terms": [
            {
                "term": term,
                "challenge_locations": challenge[term]["locations"],
                "challenge_contexts": challenge[term]["contexts"],
                "status": "招标原件候选词清单未检出，核对文件版本、页码或图片/OCR来源",
            }
            for term in sorted(challenge_terms - tender_terms)
        ],
        "shared_terms": sorted(challenge_terms & tender_terms),
    }
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(payload + "\n", encoding="utf-8")
        print(args.out)
    else:
        print(payload)


if __name__ == "__main__":
    main()
