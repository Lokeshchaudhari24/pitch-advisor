import os
import json
from dotenv import load_dotenv
from groq import Groq
from src.retrieval import search_policy

load_dotenv()

client = Groq(api_key=os.getenv("GROQ_API_KEY"))


def map_risks_to_policies(key_risks, top_k=8):

    mappings = []

    for risk in key_risks:

        if isinstance(risk, dict):
            risk_title = str(
                risk.get("title") or risk.get("risk") or risk.get("name") or ""
            ).strip()
            risk_description = str(
                risk.get("description") or risk.get("short_description") or ""
            ).strip()
            risk_detail = str(
                risk.get("detail") or risk.get("reason") or ""
            ).strip()
        else:
            risk_title = str(risk).strip()
            risk_description = ""
            risk_detail = ""

        if not risk_title:
            continue

        query = ". ".join(
            value for value in (risk_title, risk_description, risk_detail) if value
        )

        results = search_policy(query, top_k=top_k)

        evidence = []

        for i, document in enumerate(results["documents"][0]):

            source = results["metadatas"][0][i]

            evidence.append({
                "text": document,
                "source": source
            })

        context = "\n\n".join(
            f"""
SOURCE: {item['source']}
EVIDENCE:
{item['text']}
"""
            for item in evidence
        )

        prompt = f"""
You are an insurance analyst helping a Marsh advisor.

CLIENT RISK:
{risk_title}
SHORT DESCRIPTION: {risk_description or "Not provided"}
CONTEXT: {risk_detail or "Not provided"}

Below are retrieved excerpts from the available insurance
policy documents.

POLICY EVIDENCE:
{context}

TASK:

Identify ALL DISTINCT insurance benefits in the evidence that
are DIRECTLY relevant to the client risk.

Important rules:

1. Use ONLY the supplied policy evidence.

2. Do NOT use general insurance knowledge.

3. Do NOT invent benefits, coverage, limits, exclusions,
   waiting periods, or conditions.

4. A benefit should be included only when the evidence
   directly supports its relevance to the stated client risk.

5. Do NOT treat generic health-insurance features as relevant
   unless the evidence clearly connects them to the client risk.

6. Look through ALL retrieved evidence before deciding.

7. Return MULTIPLE benefits when multiple DISTINCT benefits
   are directly supported.

8. Return between 0 and 3 benefits for this risk.

9. Do not repeat the same benefit using different wording.

10. If no benefit is directly relevant, return an empty list.

11. For every benefit, provide the exact policy source and page
    from the retrieved evidence.

Return ONLY valid JSON in this format:

{{
    "risk": "{risk_title}",
    "relevant_benefits": [
        {{
            "policy": "...",
            "benefit": "...",
            "reason": "...",
            "source": "...",
            "page": 0
        }}
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

        text = (
            text
            .replace("```json", "")
            .replace("```", "")
            .strip()
        )

        try:
            result = json.loads(text)

        except json.JSONDecodeError:
            result = {
                "risk": risk_title,
                "relevant_benefits": []
            }

        result["risk"] = risk_title

        mappings.append(result)

    return mappings