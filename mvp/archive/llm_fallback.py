import os
import json
from google import genai
from google.genai import types

_SYSTEM = (
    "You are a document matching assistant. Given a folder or file identifier from a client ZIP "
    "and a list of candidates from an Excel sheet, identify which candidate this identifier most "
    "likely belongs to. Folder names may use employee IDs, abbreviated names, name-ID combinations, "
    "or partial names. Return JSON only — no markdown, no explanation outside the JSON."
)


def llm_match(identifier, excel_rows):
    api_key = os.environ.get('GEMINI_API_KEY', '').strip()
    if not api_key:
        return None

    try:
        client = genai.Client(api_key=api_key)

        candidates = [
            {'ars': r.get('ars'), 'name': r.get('name'), 'emp_id': r.get('emp_id')}
            for r in excel_rows
        ]
        prompt = (
            f"Identifier: {identifier}\n"
            f"Candidates: {json.dumps(candidates)}\n\n"
            "Return this JSON:\n"
            '{"matched_ars": "<ARS value or null>", "confidence": <0-100>, "reasoning": "<one line>"}'
        )

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(system_instruction=_SYSTEM),
        )
        text = response.text.strip()

        # Strip markdown code fences if model wraps output
        if text.startswith('```'):
            parts = text.split('```')
            text = parts[1] if len(parts) > 1 else text
            if text.startswith('json'):
                text = text[4:]

        return json.loads(text.strip())

    except Exception as e:
        print(f'[LLM] Error for "{identifier}": {e}')
        return None
