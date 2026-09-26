def build_pitch(company_profile, recommendation, audit_report):
    """Build client-facing pitch data for the fixed Marsh PPT template."""

    slides = []

    slides.append({
        "title": "Company Overview",
        "content": {
            "company": company_profile.get("company_name", ""),
            "industry": company_profile.get("industry", ""),
            "company_size": company_profile.get("company_size", ""),
            "business_overview": company_profile.get("business_overview", ""),
            "key_risks": company_profile.get("key_risks", []),
            "assumptions": company_profile.get("assumptions", []),
        },
    })

    slides.append({
        "title": "Why Choose Marsh",
        "content": {
            "points": [
                "Client-specific insurance advisory based on business risks",
                "Evidence-backed policy comparison",
                "AI-assisted policy and risk analysis",
                "Claim-level audit to identify unsupported coverage statements",
            ]
        },
    })

    slides.append({
        "title": "Evidence-Based Recommendation",
        "content": {
            "recommended_policy": recommendation.get("recommended_policy", ""),
            "confidence": recommendation.get("confidence", ""),
            "reasons": recommendation.get("reasons", []),
            "supporting_benefits": recommendation.get("supporting_benefits", []),
        },
    })

    slides.append({
        "title": "AI Content Audit",
        "content": {
            "total_claims": audit_report.get("total_claims", 0),
            "supported": audit_report.get("supported", 0),
            "partially_supported": audit_report.get("partially_supported", 0),
            "unsupported": audit_report.get("unsupported", 0),
            "not_found": audit_report.get("not_found", 0),
            "pass_rate": audit_report.get("pass_rate", 0),
            "overall_status": audit_report.get("overall_status", "REVIEW_REQUIRED"),
            # Critical: preserve claim-level audit details for Slide 4.
            "claim_results": audit_report.get("claim_results", []),
        },
    })

    slides.append({
        "title": "Advisor Review",
        "content": {
            "workflow": [
                "Review claims marked REVIEW_REQUIRED",
                "Verify source policy clause and page",
                "Edit or remove unsupported statements",
                "Approve final client-facing pitch",
            ]
        },
    })

    return slides
