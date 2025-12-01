import os
import re
from typing import Dict, List

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

LINKS_FILE_PATH = os.path.join(os.path.dirname(__file__), 'LINKS_SICRO.txt')

def parse_sicro_links() -> Dict[str, Dict[str, Dict[str, Dict[str, str]]]]:
    """
    Parses LINKS_SICRO.txt and returns a nested dictionary.
    Format: {STATE: {YEAR: {MONTH: {"regular": URL, "revisado": URL}}}}
    """
    links = {}
    if not os.path.exists(LINKS_FILE_PATH):
        print(f"Warning: Links file not found at {LINKS_FILE_PATH}")
        return {}

    with open(LINKS_FILE_PATH, 'r') as f:
        for url in f:
            url = url.strip()
            if not url:
                continue

            file_name = url.split('/')[-1].lower()
            # Filter out noisy links, but keep 'revisado'
            if 'nota' in file_name or 'obra' in file_name:
                continue

            path_parts = url.split('/')
            
            year = None
            month = None
            state_acronym = None

            # Extract info from URL path
            for i, part in enumerate(path_parts):
                # Year can be like '2019' or '2019-1'
                if re.fullmatch(r'\d{4}(?:-\d)?', part):
                    year = part.split('-')[0] # Normalize '2019-1' to '2019'
                
                # Month from month name
                if part.lower() in MONTH_NAME_TO_NUMBER:
                    month = MONTH_NAME_TO_NUMBER[part.lower()]
                
                # State from state name
                if part.lower() in STATE_NAME_TO_ACRONYM:
                    state_acronym = STATE_NAME_TO_ACRONYM[part.lower()]

            if year and month and state_acronym:
                link_type = "revisado" if "revisado" in file_name else "regular"

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
