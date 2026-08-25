import sys
import json
import collections
from usdm4 import USDM4

if __name__ == "__main__":
    json_path = sys.argv[1] if len(sys.argv) > 1 else input("Path to USDM JSON file: ").strip()

    print(f"Validating {json_path} with usdm4's built-in (free, no API key) validator...")
    result = USDM4().validate(json_path)
    print(f"passed_or_not_implemented(): {result.passed_or_not_implemented()}")

    findings = result.to_dict()
    print(f"Total rule entries: {len(findings)}")

    by_status = collections.Counter(f.get("status") for f in findings)
    print("\n=== Counts by status ===")
    for status, count in by_status.most_common():
        print(f"  {status}: {count}")

    real_problems = [
        f for f in findings
        if f.get("status") not in ("Not Implemented", "Success")
    ]
    print(f"\n=== {len(real_problems)} entries that are NOT 'Not Implemented' or 'Success' ===")
    for f in real_problems:
        print(json.dumps(f, indent=2, default=str))

    out_path = "validation_report.json"
    with open(out_path, "w") as fh:
        json.dump(findings, fh, indent=2, default=str)
    print(f"\nFull report (all {len(findings)} entries) saved to {out_path}")