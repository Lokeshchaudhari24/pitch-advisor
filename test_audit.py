from src.audit import audit_pitch_claims


claims = [
    "HDFC provides a lump-sum payout in case of accidental death or permanent disability.",
    
    "HDFC per-day cash allowance directly supports employees traveling internationally.",
    
    "HDFC provides coverage for mental health conditions and burnout.",
    
    "HDFC provides outpatient benefits."
]


results = audit_pitch_claims(claims)


print("\n==============================")
print("AUDIT RESULTS")
print("==============================")


for result in results:

    print("\n--------------------------------")
    print("CLAIM:", result["claim"])
    print("STATUS:", result["status"])
    print("CONFIDENCE:", result["confidence"])
    print("SOURCE:", result["source"])
    print("PAGE:", result["page"])
    print("EXPLANATION:", result["explanation"])
    print("EVIDENCE:", result["evidence"])