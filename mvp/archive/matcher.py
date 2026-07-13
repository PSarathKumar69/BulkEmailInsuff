import re
from rapidfuzz import fuzz
from pipeline.llm_fallback import llm_match


def _norm(s):
    """Lowercase, strip separators, remove leading zeros."""
    s = re.sub(r'[\s\-_\.]', '', str(s).lower())
    return s.lstrip('0') or '0'


def _extract_numerics(identifier):
    """Extract standalone digit sequences from identifier (min 3 digits to avoid noise)."""
    return [n for n in re.findall(r'\d+', identifier) if len(n) >= 3]


def _exact_match(identifier, excel_rows):
    norm_id = _norm(identifier)
    numerics = _extract_numerics(identifier)

    for row in excel_rows:
        # Full identifier vs each field
        for field in ('emp_id', 'ars', 'check_id'):
            val = row.get(field, '')
            if val and _norm(val) == norm_id:
                return row

        # Full identifier vs name (spaces removed)
        name = row.get('name', '')
        if name and _norm(name) == norm_id:
            return row

        # Embedded numeric sequences vs ID fields
        # Handles identifiers like "Dr. Anupama- 1287438" → try "1287438" against emp_id
        for num in numerics:
            norm_num = _norm(num)
            for field in ('emp_id', 'ars', 'check_id'):
                val = row.get(field, '')
                if val and _norm(val) == norm_num:
                    return row

    return None


def _fuzzy_match(identifier, excel_rows):
    best_score, best_row = 0, None
    for row in excel_rows:
        name = row.get('name', '')
        if not name:
            continue
        score = fuzz.token_sort_ratio(identifier, name)
        if score > best_score:
            best_score, best_row = score, row
    return best_row, best_score


def _build_result(row, unit, method, confidence, status, notes=''):
    return {
        'identifier':   unit['identifier'],
        'ars':          row.get('ars', '') if row else '',
        'name':         row.get('name', '') if row else '',
        'emp_id':       row.get('emp_id', '') if row else '',
        'matched_path': unit['path'],
        'match_method': method,
        'confidence':   confidence,
        'status':       status,
        'notes':        notes or unit.get('notes', ''),
    }


_FLAG_MAP = {
    'non_identifiable':   ('Flagged — CST', 'None', 'Generic/non-identifiable filename'),
    'empty_folder':       ('Flagged — CST', 'None', 'Empty folder — no documents'),
    'password_protected': ('Flagged — CST', 'None', 'Password protected ZIP'),
    'extraction_failed':  ('Flagged — CST', 'None', 'ZIP extraction failed'),
    'duplicate_id':       ('Flagged — CST', 'None', 'Duplicate identifier'),
}


def match_unit(unit, excel_rows):
    flag = unit.get('flag')
    if flag in _FLAG_MAP:
        status, method, default_notes = _FLAG_MAP[flag]
        return _build_result(None, unit, method, 0, status,
                             unit.get('notes', default_notes))

    identifier = unit['identifier']

    # Layer 1 — Exact (full identifier + embedded numeric IDs)
    row = _exact_match(identifier, excel_rows)
    if row:
        return _build_result(row, unit, 'Exact', 100, 'Auto Matched')

    # Layer 2 — Fuzzy name match
    row, score = _fuzzy_match(identifier, excel_rows)
    score = round(score)
    if score >= 85:
        return _build_result(row, unit, 'Fuzzy', score, 'Auto Matched')
    if score >= 70:
        return _build_result(row, unit, 'Fuzzy', score, 'Needs CST Review')

    # Layer 3 — LLM fallback
    llm_result = llm_match(identifier, excel_rows)
    if llm_result and llm_result.get('matched_ars'):
        matched_row = next(
            (r for r in excel_rows if r.get('ars') == llm_result['matched_ars']), None
        )
        confidence = llm_result.get('confidence', 0)
        if matched_row and confidence >= 80:
            return _build_result(
                matched_row, unit, 'LLM', confidence, 'Needs CST Review',
                llm_result.get('reasoning', '')
            )

    return _build_result(None, unit, 'None', 0, 'No Match — CST')


def match_all(units, excel_rows):
    results = [match_unit(unit, excel_rows) for unit in units]

    # Post-match duplicate detection: multiple units resolved to the same ARS row
    ars_map = {}
    for i, r in enumerate(results):
        if r['ars'] and r['status'] == 'Auto Matched':
            ars_map.setdefault(r['ars'], []).append(i)

    for ars, indices in ars_map.items():
        if len(indices) > 1:
            for i in indices:
                results[i]['status'] = 'Needs CST Review'
                results[i]['notes'] = (
                    f'Duplicate: {len(indices)} units matched ARS {ars} — manual review needed'
                )

    return results
