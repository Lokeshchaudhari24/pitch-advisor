from src.company_profile import generate_company_profile
from src.risk_mapping import map_risks_to_policies
from src.pitch import build_pitch

profile = generate_company_profile("Infosys")

mappings = map_risks_to_policies(
    profile["key_risks"],
    top_k=5
)

pitch = build_pitch(profile, mappings)

for i, slide in enumerate(pitch, start=1):
    print("\n==============================")
    print(f"SLIDE {i}: {slide['title']}")
    print("==============================")

    print(slide["content"])