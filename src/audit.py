import os
import json
import time
from dotenv import load_dotenv
from groq import Groq, RateLimitError
from src.retrieval import search_policy

load_dotenv()

client = Groq(
    api_key=os.getenv("GROQ_API_KEY"),
    max_retries=0,
)

RATE_LIMIT_RETRIES = 3


def _create_audit_completion(prompt):
    for attempt in range(RATE_LIMIT_RETRIES + 1):
        try:
            return client.chat.completions.create(
                model="openai/gpt-oss-120b",
                messages=[{"role": "user", "content": prompt}],
            )
        except RateLimitError as error:
            if attempt == RATE_LIMIT_RETRIES:
                raise

            retry_after = error.response.headers.get("retry-after")
            try:
                delay = max(float(retry_after), 0.5)
            except (TypeError, ValueError):
                delay = min(2 ** attempt, 8)
            time.sleep(delay)


def audit_claim(claim, top_k=5):

    results = search_policy(claim, top_k=top_k)

    evidence = []

    for i, document in enumerate(results["documents"][0]):
        evidence.append({
            "text": document,
            "source": results["metadatas"][0][i],
            "distance": results["distances"][0][i]
        })

    context = "\n\n".join(
        f"""
SOURCE: {item['source']}
DISTANCE: {item['distance']}
EVIDENCE:
{item['text']}
"""
        for item in evidence
    )

    prompt = f"""
You are auditing an AI-generated insurance statement.

AI-GENERATED CLAIM:
{claim}

POLICY DOCUMENT EVIDENCE:
{context}

Determine whether the claim is actually supported by the
provided policy evidence.

STRICT RULES:

1. Do not rely on general insurance knowledge.
2. Use ONLY the supplied evidence.
3. A benefit existing in the document does not automatically
   mean the specific interpretation in the claim is supported.
4. Distinguish direct evidence from inference.
5. If only part of the claim is supported, mark it
   PARTIALLY_SUPPORTED.
6. If the evidence does not support the claim, mark it
   UNSUPPORTED.
7. Identify the exact source and page where possible.
8. Never invent a policy clause.

Return ONLY valid JSON:

{{
    "claim": "...",
    "status": "SUPPORTED/PARTIALLY_SUPPORTED/UNSUPPORTED/NOT_FOUND",
    "confidence": "High/Medium/Low",
    "explanation": "...",
    "source": "...",
    "page": 0,
    "evidence": "..."
}}
"""

    response = _create_audit_completion(prompt)

    text = response.choices[0].message.content.strip()

    text = text.replace("```json", "")
    text = text.replace("```", "")
    text = text.strip()

    return json.loads(text)


def audit_pitch_claims(claims):

    results = []

    for claim in claims:
        result = audit_claim(claim)
        results.append(result)

    return results