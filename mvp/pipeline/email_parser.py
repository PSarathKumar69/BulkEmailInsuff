import os
import json
import time
from google import genai
from google.genai import types

_SYSTEM = (
    "You are an email classification assistant for a background verification company.\n\n"
    "You receive:\n"
    "1. Email subject\n"
    "2. Email body/thread text\n"
    "3. A list of attachment filenames\n"
    "4. (Sometimes) the body text of a forwarded/attached .eml or .msg email found among "
    "the attachments — read this as part of the context, the same as the main body, before "
    "classifying anything. It often explains what the other attachments actually are, or "
    "contains its own candidate-level data.\n\n"
    "Your job:\n"
    "1. Classify each attachment filename into one of:\n"
    "   - excel_attachments : .xlsx, .xls, .xlsb, .csv, .ods files\n"
    "   - archive_attachments : .zip, .rar, .7z, .tar, .tar.gz, .tgz, .tar.bz2, .tbz2, .tar.xz files\n"
    "   - email_container_attachments : .eml, .msg files\n"
    "   - loose_file_attachments : PDF, images, DOCX, TIFF, HEIC, or any other document\n"
    "   Use the filename extension to decide. Every filename must appear in exactly one list.\n\n"
    "2. Extract any URLs from the body text and classify each as:\n"
    "   - drive      : Google Drive links (drive.google.com, docs.google.com)\n"
    "   - sharepoint : SharePoint / OneDrive links (sharepoint.com, onedrive.com, 1drv.ms)\n"
    "   - zip_link   : Direct download links — URL path ends in .zip, .rar, or .7z\n"
    "   - unknown    : Any other URL\n"
    "   Only extract URLs that look like file-sharing or download links — ignore tracking pixels, "
    "unsubscribe links, and signature images.\n\n"
    "3. body_candidate_data: If the body text OR the forwarded-email content (if given) CLEARLY "
    "states candidate-level data "
    "(a person's name paired with an identifier such as UAN, Employee ID, ARS number, or similar), "
    "extract each candidate as an object. Otherwise return null.\n"
    "   - identifier: choose the SINGLE best primary identifier in this priority order: "
    "ARS number first, then check_id, then emp_id, then UAN.\n"
    "   - identifier_type must be one of: uan | emp_id | ars | other\n"
    "   - extra_fields: a flat dict of EVERY piece of candidate data from the body — "
    "including the candidate's name and the primary identifier, plus any other fields "
    "(UAN, Aadhaar, PAN, Employee ID, ARS, phone, address, remarks, pending items, dates, S.No). "
    "Use the EXACT label or column name from the email as the key (e.g. 'Candidate', 'UAN', "
    "'Employee ID', 'Pending', 'ARS No') and the candidate's value as the value. "
    "Do NOT rename, normalise, or omit any field. If a field is blank for this candidate, omit that key.\n\n"
    "4. has_files_or_links: true if there are any attachments OR any links detected; "
    "false only when the email body alone is the complete data source (no attachments, no links).\n\n"
    "Return ONLY valid JSON. No explanation, no markdown, no text outside the JSON."
)

_OUTPUT_SCHEMA = """{
  "excel_attachments": ["mapping.xlsx"],
  "archive_attachments": ["docs.zip"],
  "email_container_attachments": ["forwarded.eml"],
  "loose_file_attachments": ["aadhaar.pdf"],
  "links": [{"url": "https://...", "type": "drive|sharepoint|zip_link|unknown"}],
  "body_candidate_data": [
    {
      "name": "Arun Kumar",
      "identifier": "6197-001269",
      "identifier_type": "ars",
      "extra_fields": {
        "Employee ID": "EMP10421",
        "UAN Number": "100412356789",
        "Aadhaar Number": "2341 8765 4321",
        "PAN Number": "ARSPK1234A",
        "Phone Number": "9876543210",
        "Address": "12, Anna Nagar, Chennai - 600040",
        "Remarks": "Aadhaar and PAN copy missing"
      }
    }
  ],
  "has_files_or_links": true
}"""

_MODELS      = ['gemini-2.5-flash', 'gemini-2.5-flash-lite']
_RETRY_WAITS = [0, 5, 15, 30]


def _strip_fences(text):
    text = text.strip()
    if text.startswith('```'):
        parts = text.split('```')
        text = parts[1] if len(parts) > 1 else text
        if text.startswith('json'):
            text = text[4:]
    return text.strip()


def _is_transient(e):
    s = str(e)
    return any(x in s for x in ('503', '429', 'UNAVAILABLE', 'RESOURCE_EXHAUSTED'))


def _empty_result():
    return {
        'excel_attachments':           [],
        'archive_attachments':         [],
        'email_container_attachments': [],
        'loose_file_attachments':      [],
        'links':                       [],
        'body_candidate_data':         None,
        'has_files_or_links':          False,
    }


def parse_email(subject, body, attachment_filenames, inner_context=None):
    """
    One Gemini call to classify email content.

    Args:
        subject             : email subject line (str)
        body                : full email body / thread text (str)
        attachment_filenames: list of filenames for uploaded attachments
        inner_context       : optional text from any top-level .eml/.msg attachment(s),
                              pre-extracted and read BEFORE this call so the forwarded
                              email's content can inform classification from the start
                              (e.g. what a generically-named ZIP/Excel attachment actually
                              is, or candidate data stated only in the forwarded email)

    Returns structured dict:
        excel_attachments, archive_attachments, email_container_attachments,
        loose_file_attachments, links, body_candidate_data, has_files_or_links

    Raises on second consecutive JSON parse failure or full retry exhaustion.
    Falls back gracefully if there are no attachments and no body — returns empty result.
    """
    api_key = os.environ.get('GEMINI_API_KEY', '').strip()
    if not api_key:
        raise Exception(
            'GEMINI_API_KEY is not configured. '
            'Email parsing requires the API key to be set in .env'
        )

    # Nothing to parse
    if not subject.strip() and not body.strip() and not attachment_filenames:
        return _empty_result()

    client = genai.Client(api_key=api_key)

    inner_section = ''
    if inner_context and inner_context.strip():
        inner_section = (
            f"\nForwarded/attached email content (.eml/.msg found in attachments) — "
            f"read this as part of the context before classifying:\n{inner_context.strip()}\n"
        )

    base_prompt = (
        f"Subject: {subject}\n\n"
        f"Body:\n{body}\n"
        f"{inner_section}\n"
        f"Attachment filenames:\n{json.dumps(attachment_filenames, ensure_ascii=False)}\n\n"
        f"Return JSON in exactly this structure:\n{_OUTPUT_SCHEMA}"
    )
    retry_prompt = (
        base_prompt
        + '\n\nCRITICAL: Your previous response was not valid JSON. '
        'Return ONLY the JSON object — no markdown, no explanation, no text outside the braces.'
    )

    json_fail_count = 0
    last_error      = None

    for model in _MODELS:
        for attempt, wait in enumerate(_RETRY_WAITS):
            if wait:
                time.sleep(wait)
            prompt = retry_prompt if json_fail_count > 0 else base_prompt
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=prompt,
                    config=types.GenerateContentConfig(system_instruction=_SYSTEM),
                )
                result = json.loads(_strip_fences(response.text))

                # Normalise — ensure all expected keys exist
                result.setdefault('excel_attachments',           [])
                result.setdefault('archive_attachments',         [])
                result.setdefault('email_container_attachments', [])
                result.setdefault('loose_file_attachments',      [])
                result.setdefault('links',                       [])
                result.setdefault('body_candidate_data',         None)

                # Recompute has_files_or_links from actual content if AI got it wrong
                has_content = bool(
                    result['excel_attachments']
                    or result['archive_attachments']
                    or result['email_container_attachments']
                    or result['loose_file_attachments']
                    or result['links']
                )
                result['has_files_or_links'] = has_content or bool(attachment_filenames)

                print(f'[EmailParser] Success with {model} on attempt {attempt + 1}')
                return result

            except json.JSONDecodeError as e:
                last_error      = e
                json_fail_count += 1
                if json_fail_count >= 2:
                    raise Exception(
                        f'Email parsing returned invalid JSON twice: {e}. '
                        'Check that the Gemini API key is valid and the model is responding correctly.'
                    )

            except Exception as e:
                last_error = e
                if _is_transient(e):
                    print(f'[EmailParser] {model} attempt {attempt + 1} — transient error: {e}')
                    if attempt == len(_RETRY_WAITS) - 1:
                        print(f'[EmailParser] All retries for {model} exhausted, trying next model…')
                else:
                    print(f'[EmailParser] {model} attempt {attempt + 1} — non-transient error: {e}')
                    break  # skip remaining retries for this model
        else:
            continue
        break

    raise Exception(f'Email parsing failed after all retries: {last_error}')
