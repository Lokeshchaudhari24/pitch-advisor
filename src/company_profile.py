import os
import json
from dotenv import load_dotenv
from groq import Groq

load_dotenv()

client = Groq(
    api_key=os.getenv("GROQ_API_KEY")
)


def generate_company_profile(company_name):

    prompt = f"""
You are helping a Marsh insurance advisor prepare a client profile.

Company: {company_name}

Provide:

1. Industry
2. Company size
3. Business overview
4. Key business risks relevant to employee/health insurance
5. Assumptions where exact information is unavailable

BUSINESS OVERVIEW RULES:
- Give a concise 2-3 sentence description of the company's
  primary business activities, operating model, and major
  business areas.
- Use only stable, high-level information.
- Do not invent precise revenue, employee counts, locations,
  market share, or other changing facts.
- If information is uncertain, mention it under assumptions.

KEY RISK RULES:
- Identify 3-5 broad risks relevant to employee health,
  workplace wellbeing, employee benefits, workplace safety,
  or related insurance.
- Do not invent highly specific company facts.
- Keep the risks practical for an insurance advisor.
- Each risk MUST have:
  1. A short title of 3-6 words.
  2. A short description of 5-10 words.
  3. A detailed explanation that gives the full context.

The short title and description are specifically designed
for presentation slides, so keep them concise.
The detailed explanation can contain the complete risk context.

Return ONLY valid JSON:

{{
    "company_name": "{company_name}",
    "industry": "...",
    "company_size": "...",
    "business_overview": "...",
    "key_risks": [
        {{
            "title": "...",
            "description": "...",
            "detail": "..."
        }},
        {{
            "title": "...",
            "description": "...",
            "detail": "..."
        }},
        {{
            "title": "...",
            "description": "...",
            "detail": "..."
        }}
    ],
    "assumptions": [
        "...",
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
    text = text.replace("```", "").strip()

    return json.loads(text)