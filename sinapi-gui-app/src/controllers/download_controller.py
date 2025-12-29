import threading
import webbrowser
try:
    from sinapi import gerar_links_sinapi, abrir_links_no_navegador
    from aninhar import aninhar_arquivos, apagar_dados_sinapi
    from formatar_aninhados import format_excel_files
    from orse import funcao_orse
except Exception:
    import importlib
    sinapi_mod = importlib.import_module("sinapi")
    aninhar_mod = importlib.import_module("aninhar")
    formatar_mod = importlib.import_module("formatar_aninhados")
    orse_mod = importlib.import_module("orse")

    gerar_links_sinapi = getattr(sinapi_mod, "gerar_links_sinapi")
    abrir_links_no_navegador = getattr(sinapi_mod, "abrir_links_no_navegador")
    aninhar_arquivos = getattr(aninhar_mod, "aninhar_arquivos")
    apagar_dados_sinapi = getattr(aninhar_mod, "apagar_dados_sinapi")
    format_excel_files = getattr(formatar_mod, "format_excel_files")
    funcao_orse = getattr(orse_mod, "funcao_orse")


def start_sinapi(ano, mes, tipo, estados_list):
    """Gera links e abre no navegador em background, preservando comportamento atual."""
    links = gerar_links_sinapi(ano, mes, tipo, estados_list=estados_list)
    if links:
        threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()
    return links


def gerar_links_sync(ano, mes, tipo, estados_list=None):
    return gerar_links_sinapi(ano, mes, tipo, estados_list=estados_list)


def start_orse(ano, mes, tipo):
    """Executa `funcao_orse` em thread de background."""
    threading.Thread(target=funcao_orse, args=(ano, mes, tipo), daemon=True).start()


def funcao_orse_sync(ano, mes, tipo):
    return funcao_orse(ano, mes, tipo)


def start_aninhar(tipo_arquivo, sicro_composicoes, sicro_equipamentos_desonerado, sicro_equipamentos, sicro_materiais):
    threading.Thread(
        target=aninhar_arquivos,
        args=(None, tipo_arquivo, sicro_composicoes, sicro_equipamentos_desonerado, sicro_equipamentos, sicro_materiais),
        daemon=True,
    ).start()


def aninhar_arquivos_sync(tipo_arquivo, sicro_composicoes, sicro_equipamentos_desonerado, sicro_equipamentos, sicro_materiais):
    return aninhar_arquivos(None, tipo_arquivo, sicro_composicoes, sicro_equipamentos_desonerado, sicro_equipamentos, sicro_materiais)


def start_formatar_aninhados():
    threading.Thread(target=format_excel_files, daemon=True).start()


def format_excel_files_sync():
    return format_excel_files()


def abrir_links(urls):
    for u in urls:
        webbrowser.open_new_tab(u)


def abrir_links_sync(urls):
    return abrir_links_no_navegador(urls)


def apagar_dados_sync():
    return apagar_dados_sinapi()
