import importlib
try:
    import sicro
except Exception:
    sicro = importlib.import_module("sicro")


def parse_sicro_links():
    return sicro.parse_sicro_links()


def get_available_months(sicro_links_data, state, year):
    return sicro.get_available_months(sicro_links_data, state, year)


def get_sicro_link(sicro_links_data, state, year, month):
    return sicro.get_sicro_link(sicro_links_data, state, year, month)
