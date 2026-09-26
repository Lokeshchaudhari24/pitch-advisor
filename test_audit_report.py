from src.audit import audit_pitch_claims
from src.audit_report import generate_audit_summary
from src.audit_report import save_audit_report


claims = [
    "HDFC provides a lump-sum payout in case of accidental death or permanent disability.",

    "HDFC per-day cash allowance directly supports employees traveling internationally.",

    "HDFC provides coverage for mental health conditions and burnout.",

    "HDFC provides outpatient benefits."
]


audit_results = audit_pitch_claims(claims)

report = generate_audit_summary(audit_results)

save_audit_report(
    report,
    "outputs/audit_report.json"
)


print("\n==============================")
print("AUDIT SUMMARY")
print("==============================")

print("Total Claims:", report["total_claims"])
print("Supported:", report["supported"])
print("Partially Supported:", report["partially_supported"])
print("Unsupported:", report["unsupported"])
print("Not Found:", report["not_found"])
print("Pass Rate:", report["pass_rate"], "%")
print("Overall Status:", report["overall_status"])