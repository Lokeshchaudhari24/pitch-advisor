from src.company_profile import generate_company_profile
from src.risk_mapping import map_risks_to_policies
from src.recommendation import recommend_policy

profile = generate_company_profile("Infosys")

mappings = map_risks_to_policies(
    profile["key_risks"],
    top_k=5
)

recommendation = recommend_policy(
    profile,
    mappings
)

print("\n==============================")
print("RECOMMENDATION")
print("==============================")

print("Policy:", recommendation["recommended_policy"])
print("Confidence:", recommendation["confidence"])

print("\nReasons:")
for reason in recommendation["reasons"]:
    print("-", reason)

print("\nSupporting Benefits:")
for benefit in recommendation["supporting_benefits"]:
    print("-", benefit)

print("\nLimitations:")
for limitation in recommendation["limitations"]:
    print("-", limitation)