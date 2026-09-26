import os
import json
from dotenv import load_dotenv
from groq import Groq

load_dotenv()

client = Groq(
    api_key=os.getenv("GROQ_API_KEY")
)


def recommend_policy(company_profile, risk_mappings):

    verified_benefits = []

    for mapping in risk_mappings:
        for benefit in mapping.get("relevant_benefits", []):
            verified_benefits.append({
                "risk": mapping["risk"],
                "policy": benefit["policy"],
                "benefit": benefit["benefit"],
                "reason": benefit["reason"],
                "source": benefit["source"],
                "page": benefit["page"]
            })

    context = json.dumps(
        verified_benefits,
        indent=2
    )

    prompt = f"""
You are an insurance advisor assisting Marsh.

Company:
{company_profile["company_name"]}

Industry:
{company_profile["industry"]}

Company risks:
{json.dumps(company_profile["key_risks"], indent=2)}

VERIFIED POLICY BENEFITS:
{context}

TASK:
Evaluate the verified benefits and determine whether there is
enough direct evidence to recommend a policy.

STRICT RULES:

1. A policy can only be recommended if its verified benefits
   directly address one or more of the company's stated risks.

2. Do NOT treat generic insurance features such as portability,
   renewal, waiting-period continuity, or migration as a direct
   solution to unrelated employee health risks.

3. Do NOT infer that a policy covers a risk merely because the
   benefit could be indirectly useful.

4. Do NOT invent coverage.

5. If no policy has a sufficiently direct evidence-based match,
   return:
   "No sufficiently supported policy recommendation."

6. Prefer policies with multiple directly supported benefits
   addressing multiple company risks.

7. Every reason must reference the actual verified evidence.

8. State important gaps or limitations.

Return ONLY valid JSON:

{{
    "recommended_policy": "...",
    "confidence": "High/Medium/Low/Insufficient",
    "reasons": [
        "..."
    ],
    "supporting_benefits": [
        {{
            "risk": "...",
            "benefit": "...",
            "source": "...",
            "page": 0
        }}
    ],
    "limitations": [
        "..."
    ]
}}
"""

    response = client.chat.completions.create(
        model="openai/gpt-oss-120b",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )

    text = response.choices[0].message.content.strip()

    text = text.replace("```json", "")
    text = text.replace("```", "")
    text = text.strip()

    return json.loads(text)