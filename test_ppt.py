from src.risk_mapping import map_risks_to_policies
from src.pitch import build_pitch
from src.ppt_generator import create_pitch_ppt


profile = {
    "company_name": "Infosys",
    "industry": "Information Technology Services and Consulting",
    "company_size": "Large Enterprise",
    "key_risks": [
        "Sedentary lifestyle and ergonomic-related health risks",
        "Mental health strain and burnout",
        "Cross-border medical exposure",
        "Medical cost inflation"
    ],
    "assumptions": []
}


mappings = map_risks_to_policies(
    profile["key_risks"]
)

pitch = build_pitch(
    profile,
    mappings
)

create_pitch_ppt(
    pitch,
    "outputs/infosys_pitch.pptx"
)