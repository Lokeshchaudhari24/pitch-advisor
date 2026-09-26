from src.company_profile import generate_company_profile
from src.risk_mapping import map_risks_to_policies
from src.recommendation import recommend_policy
from src.audit import audit_pitch_claims
from src.audit_report import generate_audit_summary
from src.pitch import build_pitch


# 1. Company profile
profile = generate_company_profile("Infosys")


# 2. Risk → policy mapping
mappings = map_risks_to_policies(
    profile["key_risks"],
    top_k=5
)


# 3. Policy recommendation
recommendation = recommend_policy(
    profile,
    mappings
)


# 4. Claims to audit
claims = []

for benefit in recommendation.get(
    "supporting_benefits", []
):

    claims.append(
        benefit["benefit"]
    )


# 5. Audit claims
audit_results = audit_pitch_claims(claims)

# 6. Audit summary
audit_report = generate_audit_summary(
    audit_results
)


# 7. Keep ONLY supported claims
supported_claims = []

for result in audit_results:

    if result["status"] == "SUPPORTED":
        supported_claims.append(result)


# 8. Build final recommendation from audited claims
verified_recommendation = recommendation.copy()

verified_recommendation["supporting_benefits"] = []

for result in supported_claims:

    verified_recommendation["supporting_benefits"].append({
        "benefit": result["claim"],
        "source": result["source"],
        "page": result["page"]
    })


# If recommendation contains unsupported claims,
# mark it for advisor review.
if len(supported_claims) < len(claims):

    verified_recommendation["confidence"] = "Review Required"


# 9. Build final pitch
pitch = build_pitch(
    profile,
    verified_recommendation,
    audit_report
)


# 10. Print final pitch
print("\n==============================")
print("FINAL VERIFIED PITCH")
print("==============================")

for i, slide in enumerate(
    pitch,
    start=1
):

    print(
        f"\nSLIDE {i}: {slide['title']}"
    )

    print(slide["content"])


print("\n==============================")
print("AUDIT STATUS")
print("==============================")

print(
    "Overall:",
    audit_report["overall_status"]
)

print(
    "Pass Rate:",
    audit_report["pass_rate"],
    "%"
)

print(
    "Verified Claims:",
    len(supported_claims)
)

print(
    "Total Claims:",
    len(claims)
)