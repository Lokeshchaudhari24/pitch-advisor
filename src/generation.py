import os
from dotenv import load_dotenv
from groq import Groq

# Load API key from .env
load_dotenv()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

def generate_answer(question, context):
    prompt = f"""
    You are an insurance policy assistant.

    Answer the question ONLY using the policy information provided below.
    If the information is not present in the context, say:
    "Not supported by the provided policy documents."
    Do not invent coverage, limits, exclusions, waiting periods, or benefits.

    POLICY CONTEXT:
    {context}
    
    QUESTION: {question}
    
    Provide a concise factual answer.
    """
    
    # Llama 3 model use kar rahe hain jo fast aur completely free hai
    response = client.chat.completions.create(
        messages=[{"role": "user", "content": prompt}],
        model="openai/gpt-oss-120b", 
        temperature=0.1
    )
    
    return response.choices[0].message.content

