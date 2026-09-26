import json


def generate_audit_summary(audit_results):

    total = len(audit_results)

    supported = 0
    partially_supported = 0
    unsupported = 0
    not_found = 0

    for result in audit_results:

        status = result["status"]

        if status == "SUPPORTED":
            supported += 1

        elif status == "PARTIALLY_SUPPORTED":
            partially_supported += 1

        elif status == "UNSUPPORTED":
            unsupported += 1

        elif status == "NOT_FOUND":
            not_found += 1

    if total > 0:
        pass_rate = round((supported / total) * 100, 2)
    else:
        pass_rate = 0

    if unsupported > 0:
        overall_status = "REVIEW_REQUIRED"
    elif partially_supported > 0:
        overall_status = "REVIEW_REQUIRED"
    else:
        overall_status = "PASSED"

    return {
        "total_claims": total,
        "supported": supported,
        "partially_supported": partially_supported,
        "unsupported": unsupported,
        "not_found": not_found,
        "pass_rate": pass_rate,
        "overall_status": overall_status,
        "claim_results": audit_results
    }


def save_audit_report(report, output_path):

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(
            report,
            f,
            indent=4,
            ensure_ascii=False
        )

    print(f"Audit report saved: {output_path}")