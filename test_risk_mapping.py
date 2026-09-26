from src.company_profile import generate_company_profile
from src.risk_mapping import map_risks_to_policies


profile = generate_company_profile("Infosys")

print("\nCOMPANY:", profile["company_name"])

mappings = map_risks_to_policies(
    profile["key_risks"],
    top_k=5
)

for mapping in mappings:

    print("\n================================")
    print("RISK:", mapping["risk"])

    if not mapping["relevant_benefits"]:
        print("No directly supported policy benefit found.")
        continue

    for benefit in mapping["relevant_benefits"]:

        print("\nPOLICY:", benefit["policy"])
        print("BENEFIT:", benefit["benefit"])
        print("REASON:", benefit["reason"])
        print("SOURCE:", benefit["source"])
        print("PAGE:", benefit["page"])