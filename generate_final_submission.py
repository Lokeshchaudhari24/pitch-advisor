from src.company_profile import generate_company_profile
from src.risk_mapping import map_risks_to_policies
from src.audit import audit_pitch_claims
from src.audit_report import generate_audit_summary
from src.audit_report import save_audit_report
from src.pitch import build_pitch
from src.ppt_generator import create_pitch_ppt


# ==========================================
# 1. COMPANY PROFILE
# ==========================================

profile = generate_company_profile("Infosys")


# ==========================================
# 2. RISK → POLICY MAPPING
# ==========================================

mappings = map_risks_to_policies(
    profile["key_risks"],
    top_k=5
)


# ==========================================
# 3. COLLECT VERIFIED CANDIDATE BENEFITS
# ==========================================

candidate_benefits = []

for mapping in mappings:

    for benefit in mapping.get(
        "relevant_benefits",
        []
    ):

        candidate_benefits.append({
            "risk": mapping["risk"],
            "benefit": benefit["benefit"],
            "policy": benefit["policy"],
            "source": benefit["source"],
            "page": benefit["page"]
        })


# ==========================================
# 4. AUDIT EVERY CANDIDATE CLAIM
# ==========================================

claims = [
    item["benefit"]
    for item in candidate_benefits
]

audit_results = audit_pitch_claims(
    claims
)


# ==========================================
# 5. KEEP ONLY SUPPORTED CLAIMS
# ==========================================

verified_benefits = []

for benefit, audit in zip(
    candidate_benefits,
    audit_results
):

    if audit["status"] == "SUPPORTED":

        verified_benefits.append({
            "risk": benefit["risk"],
            "benefit": audit["claim"],
            "policy": benefit["policy"],
            "source": audit["source"],
            "page": audit["page"]
        })


# ==========================================
# 6. BUILD EVIDENCE-BASED RECOMMENDATION
# ==========================================

policy_scores = {}

for benefit in verified_benefits:

    policy = benefit["policy"]

    if policy not in policy_scores:
        policy_scores[policy] = []

    policy_scores[policy].append(
        benefit
    )


if policy_scores:

    # Select policy with the largest number
    # of independently supported benefits.
    recommended_policy = max(
        policy_scores,
        key=lambda p: len(
            policy_scores[p]
        )
    )

    supporting_benefits = policy_scores[
        recommended_policy
    ]

    confidence = "High"

    reasons = [
        (
            f"{item['benefit']} "
            f"(Source: {item['source']}, "
            f"Page: {item['page']})"
        )
        for item in supporting_benefits
    ]

else:

    recommended_policy = (
        "No sufficiently supported policy recommendation"
    )

    supporting_benefits = []

    confidence = "Insufficient"

    reasons = [
        "No candidate policy benefit passed "
        "the claim-level audit."
    ]


recommendation = {
    "recommended_policy": recommended_policy,
    "confidence": confidence,
    "reasons": reasons,
    "supporting_benefits": supporting_benefits,
    "limitations": [
        "Only claims directly supported by "
        "the policy documents were considered."
    ]
}


# ==========================================
# 7. FINAL AUDIT SUMMARY
# ==========================================

audit_report = generate_audit_summary(
    audit_results
)


# ==========================================
# 8. SAVE AUDIT REPORT
# ==========================================

save_audit_report(
    audit_report,
    "outputs/audit_report.json"
)


# ==========================================
# 9. BUILD FINAL PITCH
# ==========================================

pitch = build_pitch(
    profile,
    recommendation,
    audit_report
)


# ==========================================
# 10. GENERATE PPT
# ==========================================

create_pitch_ppt(
    pitch,
    "outputs/infosys_final_pitch.pptx"
)


# ==========================================
# 11. SUMMARY
# ==========================================

print("\n================================")
print("FINAL SUBMISSION GENERATED")
print("================================")

print(
    "Recommended Policy:",
    recommended_policy
)

print(
    "Confidence:",
    confidence
)

print(
    "Supported Candidate Claims:",
    len(verified_benefits)
)

print(
    "Total Candidate Claims:",
    len(candidate_benefits)
)

print(
    "Audit Status:",
    audit_report["overall_status"]
)

print(
    "Pass Rate:",
    audit_report["pass_rate"],
    "%"
)

print(
    "PPT:",
    "outputs/infosys_final_pitch.pptx"
)

print(
    "Audit:",
    "outputs/audit_report.json"
)