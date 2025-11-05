import unittest
from src.sinapi import gerar_links_sinapi

class TestSinapi(unittest.TestCase):

    def test_gerar_links_sinapi_valid(self):
        links = gerar_links_sinapi(2024, 11, "Ambos")
        self.assertEqual(len(links), 54)  # 2 types * 27 states

    def test_gerar_links_sinapi_invalid_tipo(self):
        with self.assertRaises(ValueError):
            gerar_links_sinapi(2024, 11, "InvalidTipo")

    def test_gerar_links_sinapi_single_type(self):
        links_desonerado = gerar_links_sinapi(2024, 11, "Desonerado")
        links_nao_desonerado = gerar_links_sinapi(2024, 11, "NaoDesonerado")
        self.assertEqual(len(links_desonerado), 27)  # 1 type * 27 states
        self.assertEqual(len(links_nao_desonerado), 27)  # 1 type * 27 states

    def test_gerar_links_sinapi_edge_case(self):
        links = gerar_links_sinapi(2024, 0, "Ambos")
        self.assertEqual(links, [])  # No valid month should return empty list

if __name__ == '__main__':
    unittest.main()