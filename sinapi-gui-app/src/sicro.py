import os
import re
import sys
from typing import Dict, List

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # not in bundle, so we can get it from the script's directory
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)

MONTH_NAME_TO_NUMBER = {
    'janeiro': '01', 'fevereiro': '02', 'marco': '03', 'abril': '04', 
    'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08', 
    'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12'
}

STATE_NAME_TO_ACRONYM = {
    'acre': 'AC', 'alagoas': 'AL', 'amapa': 'AP', 'amazonas': 'AM', 
    'bahia': 'BA', 'ceara': 'CE', 'distrito-federal': 'DF', 'espirito-santo': 'ES', 
    'goias': 'GO', 'maranhao': 'MA', 'mato-grosso': 'MT', 'mato-grosso-do-sul': 'MS', 
    'minas-gerais': 'MG', 'para': 'PA', 'paraiba': 'PB', 'parana': 'PR', 
    'pernambuco': 'PE', 'piaui': 'PI', 'rio-de-janeiro': 'RJ', 
    'rio-grande-do-norte': 'RN', 'rio-grande-do-sul': 'RS', 'rondonia': 'RO', 
    'roraima': 'RR', 'santa-catarina': 'SC', 'sao-paulo': 'SP', 'sergipe': 'SE', 
    'tocantins': 'TO'
}

LINKS_FILE_PATH = resource_path('LINKS_SICRO.txt')

def parse_sicro_links() -> Dict[str, Dict[str, Dict[str, Dict[str, str]]]]:
    """
    Parses LINKS_SICRO.txt and returns a nested dictionary.
    Format: {STATE: {YEAR: {MONTH: {"regular": URL, "revisado": URL}}}}
    """
    links = {}
    if not os.path.exists(LINKS_FILE_PATH):
        print(f"Warning: Links file not found at {LINKS_FILE_PATH}")
        return {}

    # Read all URLs first so we can do a two-pass parsing (helps when year is missing in some URLs)
    with open(LINKS_FILE_PATH, 'r') as f:
        raw_urls = [u.strip() for u in f if u.strip()]

    # Prepare regex and helper
    month_pattern = re.compile(r"^(?:" + "|".join(MONTH_NAME_TO_NUMBER.keys()) + r")(?:-\d+)?$", re.IGNORECASE)

    records = []
    # First pass: extract state, month (base), and year when available
    for url in raw_urls:
        file_name = url.split('/')[-1].lower()
        # Filter out noisy links, but keep 'revisado'
        if 'nota' in file_name or 'obra' in file_name:
            continue

        path_parts = url.split('/')
        year = None
        month = None
        state_acronym = None

        for part in path_parts:
            if re.fullmatch(r'\d{4}(?:-\d)?', part):
                year = part.split('-')[0]

            m = month_pattern.match(part)
            if m:
                base_month = re.split(r'-', part, maxsplit=1)[0].lower()
                if base_month in MONTH_NAME_TO_NUMBER:
                    month = MONTH_NAME_TO_NUMBER[base_month]

            if part.lower() in STATE_NAME_TO_ACRONYM:
                state_acronym = STATE_NAME_TO_ACRONYM[part.lower()]

        # Also try to extract year from filename
        if not year:
            m_year = re.search(r"(\d{4})", file_name)
            if m_year:
                year = m_year.group(1)

        records.append({
            'url': url,
            'file_name': file_name,
            'state': state_acronym,
            'month': month,
            'year': year,
            'is_revisado': 'revisado' in file_name,
            'index': len(records)
        })

    # Build candidate years mapping for (state, month)
    candidate_years = {}
    for r in records:
        if r['state'] and r['month'] and r['year']:
            key = (r['state'], r['month'])
            candidate_years.setdefault(key, set()).add(r['year'])

    # Second pass: create links, filling missing years when possible
    for idx, r in enumerate(records):
        state_acronym = r['state']
        month = r['month']
        year = r['year']
        url = r['url']
        link_type = 'revisado' if r['is_revisado'] else 'regular'

        if not (state_acronym and month):
            continue

        if not year:
            key = (state_acronym, month)
            years = candidate_years.get(key)
            if years and len(years) == 1:
                year = next(iter(years))
            else:
                # Try to find nearest record (by index) with same state/month that has a year
                found_year = None
                max_search = max(100, len(records))
                for offset in range(1, len(records)):
                    prev_idx = idx - offset
                    next_idx = idx + offset
                    cand = None
                    if prev_idx >= 0:
                        cand = records[prev_idx]
                        if cand['state'] == state_acronym and cand['month'] == month and cand['year']:
                            found_year = cand['year']
                            break
                    if next_idx < len(records):
                        cand = records[next_idx]
                        if cand['state'] == state_acronym and cand['month'] == month and cand['year']:
                            found_year = cand['year']
                            break
                if found_year:
                    year = found_year
                else:
                    # unable to determine year
                    continue

        if state_acronym not in links:
            links[state_acronym] = {}
        if year not in links[state_acronym]:
            links[state_acronym][year] = {}
        if month not in links[state_acronym][year]:
            links[state_acronym][year][month] = {}

        links[state_acronym][year][month][link_type] = url

    return links

def get_available_months(all_links: Dict, state: str, year: str) -> List[str]:
    """
    Returns a sorted list of available months for a given state and year.
    Includes "(revisado)" for revised versions.
    """
    months = []
    if state in all_links and year in all_links[state]:
        for month_num, types in sorted(all_links[state][year].items()):
            if "regular" in types:
                months.append(month_num)
            if "revisado" in types:
                months.append(f"{month_num} (revisado)")
    return months

def get_sicro_link(all_links: Dict, state: str, year: str, month_str: str) -> str:
    """
    Returns the download URL for the given state, year, and month string.
    The month_str can be e.g. "10" or "10 (revisado)".
    """
    if not month_str:
        return ""
        
    month_num = month_str.split(' ')[0]
    link_type = "revisado" if "revisado" in month_str else "regular"
    
    return all_links.get(state, {}).get(year, {}).get(month_num, {}).get(link_type, "")

if __name__ == '__main__':
    # Example usage
    sicro_links_data = parse_sicro_links()
    print("Parsed SICRO Links:")
    import json
    print(json.dumps(sicro_links_data, indent=2))
    
    # Test getting available months
    test_state = 'PR'
    test_year = '2022'
    available_months = get_available_months(sicro_links_data, test_state, test_year)
    print(f"Available months for {test_state}/{test_year}: {available_months}")
    
    # Test getting a specific link for regular and revised
    test_month_regular = '10'
    link_regular = get_sicro_link(sicro_links_data, test_state, test_year, test_month_regular)
    print(f"Link for {test_state}/{test_year}/{test_month_regular}: {link_regular}")
    
    test_month_revised = '10 (revisado)'
    link_revised = get_sicro_link(sicro_links_data, test_state, test_year, test_month_revised)
    print(f"Link for {test_state}/{test_year}/{test_month_revised}: {link_revised}")
