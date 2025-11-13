import os
import shutil
import zipfile
import re
from datetime import datetime
from typing import List, Tuple, Optional
import pandas as pd
import sinapi

print("VALOR ATUAL! ", sinapi.obter_valor_janeiro_2021())
print("VALOR ALTERADO! ", sinapi.definir_valor_janeiro_2021(True))
print("VALOR ATUAL! ", sinapi.obter_valor_janeiro_2021())
print("VALOR ALTERADO! ", sinapi.definir_valor_janeiro_2021(False))
print("VALOR ATUAL! ", sinapi.obter_valor_janeiro_2021())


def apagar_dados_sinapi():
    """Apaga todos os arquivos e pastas da pasta SinapiDownloads."""
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    base_dir = os.path.join(desktop_dir, "SinapiDownloads")
    if os.path.isdir(base_dir):
        for item in os.listdir(base_dir):
            item_path = os.path.join(base_dir, item)
            try:
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    os.unlink(item_path)
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                print(f"Item deletado: {item_path}")
            except Exception as e:
                print(f'Falha ao deletar {item_path}. Razão: {e}')
                
                

def aninhar_arquivos(base_dir: Optional[str] = None) -> Tuple[List[str], str]:
    """
    1) Move todos os arquivos da pasta Downloads cujo nome contém 'sinapi'
       (case-insensitive) para a pasta "Sinapi downloads" na Área de Trabalho
       do usuário atual. Cria a pasta de destino se não existir.
    2) Extrai todos os .zip encontrados diretamente dentro da pasta "Sinapi downloads".
    3) Procura recursivamente por arquivos Excel cujo nome contenha 'sintetico' ou 'insumos'
       (case-insensitive) dentro de base_dir (inclui conteúdo extraído).
    4) Junta todas as abas encontradas em um único arquivo Excel em <base_dir>/aninhar/.
    Retorna (moved_paths, out_path) - moved_paths: lista de arquivos movidos; out_path: caminho do arquivo aninhado
    (out_path será "" se não houve arquivos Excel encontrados).
    """
    # determina base_dir (Sinapi downloads na Área de Trabalho)
    if base_dir is None:
        desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        base_dir = os.path.join(desktop_dir, "SinapiDownloads")
    os.makedirs(base_dir, exist_ok=True)
    base_dir = os.path.abspath(base_dir)

    # 1) mover arquivos da pasta Downloads que contenham 'sinapi'
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    moved: List[str] = []
    token = "sinapi"

    if os.path.isdir(downloads_dir):
        for root, _, files in os.walk(downloads_dir):
            for fname in files:
                if token in fname.lower():
                    src = os.path.join(root, fname)
                    dst = os.path.join(base_dir, fname)
                    base, ext = os.path.splitext(fname)
                    counter = 1
                    while os.path.exists(dst):
                        dst = os.path.join(base_dir, f"{base}_{counter}{ext}")
                        counter += 1
                    try:
                        shutil.move(src, dst)
                        moved.append(dst)
                    except Exception as e:
                        print(f"Falha ao mover {src} -> {dst}: {e}")

    print(f"Arquivos movidos ({len(moved)}):")
    for p in moved:
        print(p)

    # 2) extrair todos os .zip no diretório base (não recursivo)
    extracted_root = os.path.join(base_dir, "__extracted_zips__")
    os.makedirs(extracted_root, exist_ok=True)
    
    # Primeira extração - zips da pasta base para __extracted_zips__
    zip_files = [f for f in os.listdir(base_dir) if f.lower().endswith(".zip")]
    for zname in zip_files:
        zpath = os.path.join(base_dir, zname)
        try:
            subdir = os.path.join(extracted_root, os.path.splitext(zname)[0])
            os.makedirs(subdir, exist_ok=True)
            with zipfile.ZipFile(zpath, "r") as zf:
                zf.extractall(subdir)
        except zipfile.BadZipFile:
            print(f"Arquivo zip inválido, pulando: {zpath}")
        except Exception as e:
            print(f"Falha ao extrair {zpath}: {e}")

    # 2b) Extrair recursivamente todos os arquivos .zip dentro de __extracted_zips__
    while True:
        # Encontra todos os arquivos .zip em todas as subpastas de __extracted_zips__
        nested_zips = []
        for root, _, files in os.walk(extracted_root):
            for file in files:
                if file.lower().endswith(".zip"):
                    nested_zips.append(os.path.join(root, file))

        # Se não houver mais zips para extrair, o processo está completo.
        if not nested_zips:
            break

        # Itera sobre os zips encontrados para extraí-los
        for zip_path in nested_zips:
            # O destino da extração é a pasta onde o zip se encontra.
            extract_dir = os.path.dirname(zip_path)
            try:
                with zipfile.ZipFile(zip_path, "r") as zf:
                    # Usar extractall para extrair na pasta pai do zip.
                    zf.extractall(extract_dir)
                
                # Remove o arquivo zip após a extração bem-sucedida para evitar reprocessamento.
                os.remove(zip_path)

            except zipfile.BadZipFile:
                print(f"Zip interno inválido, pulando: {zip_path}")
            except Exception as e:
                print(f"Falha ao extrair zip interno {zip_path}: {e}")
                
    # filtrar arquivos de 2021 (1a4)
    lista_manter = []
    
    # Condicional, caso a variável seja verdadeira
    if sinapi.obter_valor_janeiro_2021() == True:
        lista_manter.append("202101")
        lista_manter.append("012021")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_janeiro_2021(False)
    
    if sinapi.obter_valor_fevereiro_2021() == True:
        lista_manter.append("202102")
        lista_manter.append("022021")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_fevereiro_2021(False)
    
    if sinapi.obter_valor_marco_2021() == True:
        lista_manter.append("202103")
        lista_manter.append("032021")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_marco_2021(False)
        
    if sinapi.obter_valor_abril_2021() == True:
        lista_manter.append("202104")
        lista_manter.append("042021")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_abril_2021(False)

    if sinapi.obter_valor_setembro_2020() == True:
        lista_manter.append("202009")
        lista_manter.append("092020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_setembro_2020(False)
    
    if sinapi.obter_valor_outubro_2020() == True:
        lista_manter.append("202010")
        lista_manter.append("102020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_outubro_2020(False)

    if sinapi.obter_valor_novembro_2020() == True:
        lista_manter.append("202011")
        lista_manter.append("112020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_novembro_2020(False)

    if sinapi.obter_valor_dezembro_2020() == True:
        lista_manter.append("202012")
        lista_manter.append("122020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_dezembro_2020(False)

    if sinapi.obter_valor_maio_2020() == True:
        lista_manter.append("202005")
        lista_manter.append("052020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_maio_2020(False)
    
    if sinapi.obter_valor_junho_2020() == True:
        lista_manter.append("202006")
        lista_manter.append("062020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_junho_2020(False)

    if sinapi.obter_valor_julho_2020() == True:
        lista_manter.append("202007")
        lista_manter.append("072020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_julho_2020(False)

    if sinapi.obter_valor_agosto_2020() == True:
        lista_manter.append("202008")
        lista_manter.append("082020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_agosto_2020(False)

    if sinapi.obter_valor_janeiro_2020() == True:
        lista_manter.append("202001")
        lista_manter.append("012020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_janeiro_2020(False)
    
    if sinapi.obter_valor_fevereiro_2020() == True:
        lista_manter.append("202002")
        lista_manter.append("022020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_fevereiro_2020(False)

    if sinapi.obter_valor_marco_2020() == True:
        lista_manter.append("202003")
        lista_manter.append("032020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_marco_2020(False)

    if sinapi.obter_valor_abril_2020() == True:
        lista_manter.append("202004")
        lista_manter.append("042020")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_abril_2020(False)

    if sinapi.obter_valor_julho_2018() == True:
        lista_manter.append("201807")
        lista_manter.append("072018")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_julho_2018(False)
    
    if sinapi.obter_valor_agosto_2018() == True:
        lista_manter.append("201808")
        lista_manter.append("082018")
        print("Valor atual da lista: ", lista_manter)
        sinapi.definir_valor_agosto_2018(False)
    
    # Deletar arquivos que não estão na lista_manter
    specific_folder_path = os.path.join(extracted_root, "SINAPI_ref_Insumos_Composicoes_AC_2021_01a04_Retificacao01")
    if os.path.isdir(specific_folder_path):
        if lista_manter: # Apenas executa se a lista não estiver vazia
            print(f"Itens a manter: {lista_manter}")
            for root, _, files in os.walk(specific_folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    manter_arquivo = False
                    for item in lista_manter:
                        if item in file:
                            manter_arquivo = True
                            break
                    
                    # Condição adicional para deletar arquivos com "Analitico" no nome ou que sejam .pdf
                    deletar_por_regra_adicional = "Analitico" in file or file.lower().endswith('.pdf')

                    if not manter_arquivo or deletar_por_regra_adicional:
                        try:
                            os.remove(file_path)
                            print(f"Arquivo deletado: {file_path}")
                        except OSError as e:
                            print(f"Erro ao deletar o arquivo {file_path}: {e}")
    
    specific_folder_path_2020_09a12 = os.path.join(extracted_root, "SINAPI_ref_Insumos_Composicoes_AC_2020_09a12_Retificacao01")
    if os.path.isdir(specific_folder_path_2020_09a12):
        if lista_manter: # Apenas executa se a lista não estiver vazia
            print(f"Itens a manter: {lista_manter}")
            for root, _, files in os.walk(specific_folder_path_2020_09a12):
                for file in files:
                    file_path = os.path.join(root, file)
                    manter_arquivo = False
                    for item in lista_manter:
                        if item in file:
                            manter_arquivo = True
                            break
                    
                    deletar_por_regra_adicional = "Analitico" in file or file.lower().endswith('.pdf')

                    if not manter_arquivo or deletar_por_regra_adicional:
                        try:
                            os.remove(file_path)
                            print(f"Arquivo deletado: {file_path}")
                        except OSError as e:
                            print(f"Erro ao deletar o arquivo {file_path}: {e}")
    
    specific_folder_path_2020_05a08 = os.path.join(extracted_root, "SINAPI_ref_Insumos_Composicoes_AC_2020_05a08")
    if os.path.isdir(specific_folder_path_2020_05a08):
        if lista_manter: # Apenas executa se a lista não estiver vazia
            print(f"Itens a manter: {lista_manter}")
            for root, _, files in os.walk(specific_folder_path_2020_05a08):
                for file in files:
                    file_path = os.path.join(root, file)
                    manter_arquivo = False
                    for item in lista_manter:
                        if item in file:
                            manter_arquivo = True
                            break
                    
                    deletar_por_regra_adicional = "Analitico" in file or file.lower().endswith('.pdf')

                    if not manter_arquivo or deletar_por_regra_adicional:
                        try:
                            os.remove(file_path)
                            print(f"Arquivo deletado: {file_path}")
                        except OSError as e:
                            print(f"Erro ao deletar o arquivo {file_path}: {e}")
    
    specific_folder_path_2020_01a04 = os.path.join(extracted_root, "SINAPI_ref_Insumos_Composicoes_AC_2020_01a04")
    if os.path.isdir(specific_folder_path_2020_01a04):
        if lista_manter: # Apenas executa se a lista não estiver vazia
            print(f"Itens a manter: {lista_manter}")
            for root, _, files in os.walk(specific_folder_path_2020_01a04):
                for file in files:
                    file_path = os.path.join(root, file)
                    manter_arquivo = False
                    for item in lista_manter:
                        if item in file:
                            manter_arquivo = True
                            break
                    
                    deletar_por_regra_adicional = "Analitico" in file or file.lower().endswith('.pdf')

                    if not manter_arquivo or deletar_por_regra_adicional:
                        try:
                            os.remove(file_path)
                            print(f"Arquivo deletado: {file_path}")
                        except OSError as e:
                            print(f"Erro ao deletar o arquivo {file_path}: {e}")

    specific_folder_path_2018_07e08 = os.path.join(extracted_root, "SINAPI_ref_Insumos_Composicoes_AC_2018_07e08_Retificacao")
    if os.path.isdir(specific_folder_path_2018_07e08):
        if lista_manter: # Apenas executa se a lista não estiver vazia
            print(f"Itens a manter: {lista_manter}")
            for root, _, files in os.walk(specific_folder_path_2018_07e08):
                for file in files:
                    file_path = os.path.join(root, file)
                    manter_arquivo = False
                    for item in lista_manter:
                        if item in file:
                            manter_arquivo = True
                            break
                    
                    deletar_por_regra_adicional = "Analitico" in file or file.lower().endswith('.pdf')

                    if not manter_arquivo or deletar_por_regra_adicional:
                        try:
                            os.remove(file_path)
                            print(f"Arquivo deletado: {file_path}")
                        except OSError as e:
                            print(f"Erro ao deletar o arquivo {file_path}: {e}")

    # 3) localizar arquivos Excel relevantes (recursivo, inclui conteúdo extraído)
    exts = {'.xls', '.xlsx', '.xlsm', '.xlsb'}
    tokens = ("Sintetico", "Insumos")
    matches: List[str] = []
    output_dir = os.path.join(base_dir, "aninhar")
    os.makedirs(output_dir, exist_ok=True)

    for root, _, files in os.walk(base_dir):
        # evita procurar dentro da própria pasta de saída
        if os.path.abspath(root).startswith(os.path.abspath(output_dir)):
            continue
        for fname in files:
            low = fname.lower()
            ext = os.path.splitext(low)[1]
            # só considerar arquivos Excel
            if ext not in exts:
                continue
            # condições solicitadas:
            # - inclui se o nome contém "sintetico"
            # - inclui se o nome contém "insumos" desde que NÃO contenha "família" (ou "familia")
            has_sint = "sintetico" in low
            has_insumos = "insumos" in low
            has_familia = "família" in low or "familia" in low
            if has_sint or (has_insumos and not has_familia):
                matches.append(os.path.join(root, fname))

    if not matches:
        print("Nenhum arquivo Excel com 'sintetico' ou 'insumos' encontrado após extração.")
        return moved, ""

    # 4) unir abas em um único arquivo Excel de saída
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(output_dir, f"aninhado_{timestamp}.xlsx")

    def _safe_sheet_name(name: str) -> str:
        invalid = ['\\','/','*','[',']',':','?']
        s = "".join("_" if c in invalid else c for c in name)
        s = s.strip()
        if len(s) == 0:
            s = "sheet"
        return s[:31]

    def _format_tab_name_from_filename(fname: str) -> str:
        """
        Constrói o prefixo do nome da aba a partir do nome do arquivo:
          - Sintético -> 'SINA-SIN'
          - Insumos   -> 'SINA-INS'
        Extrai sigla do estado (ex: 'AC') e data no formato YYYYMM (ex: '202401').
        Sufixo de tipo: 'DES' para Desonerado, 'NDS' para Não Desonerado.
        Retorna string já curta; caller deve garantir unicidade.
        """
        name = os.path.splitext(os.path.basename(fname))[0]
        low = name.lower()

        print(low)
        
        # prefixo
        if "sintetico" in low:
            pref = "SINA-SIN"
        elif "insumos" in low:
            pref = "SINA-INS"
        else:
            pref = "SINA-UNK"

        # estado: procurar token de duas letras entre underscores ou token standalone uppercase
        estado = ""
        m = re.search(r"_([A-Z]{2})_", name)
        if m:
            estado = m.group(1)
        else:
            # tentar tokens passando por '_' separador
            parts = name.split("_")
            for p in parts:
                if re.fullmatch(r"[A-Z]{2}", p):
                    estado = p
                    break

        # data: procurar YYYYMM ou MMYYYY ou YYYY_MM
        date_token = ""
        m = re.search(r"(\d{4}_\d{2})", name)
        if m:
            date_token = m.group(1).replace("_", "")
        else:
            m = re.search(r"(\d{6})", name)
            if m:
                date_token = m.group(1)
            else:
                m = re.search(r"_(\d{2})(\d{4})_", "_" + name + "_")
                if m:
                    # found _MMYYYY_
                    date_token = f"{m.group(2)}{m.group(1)}"

        # tipo abreviado
        if "naodesonerado" in low:
            tipo_abbr = "NDS"
        # elif "NaoDesonerado" in low or "naoDesonerado" in low or "nao" in low or "não" in low:
        #     tipo_abbr = "NDS"
        elif "desonerado" in low:
            tipo_abbr = "DES"
        else:
            tipo_abbr = "UNK"

        # montar partes (omitir vazios)
        parts = [pref]
        if estado:
            parts.append(estado)
        if date_token:
            parts.append(date_token)
        parts.append(tipo_abbr)
        tab = "-".join(parts)
        # garantir limite de 31 caracteres
        return _safe_sheet_name(tab)

    used_sheet_names = set()
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for fpath in matches:
            try:
                sheets = pd.read_excel(fpath, sheet_name=None)
            except Exception as e:
                print(f"Falha ao ler '{fpath}': {e}. Pulando.")
                continue

            base_name = os.path.splitext(os.path.basename(fpath))[0]
            base_name = base_name.replace(" ", "_")
            # nome base sugerido a partir do nome do arquivo
            file_based_tab = _format_tab_name_from_filename(fpath)

            if isinstance(sheets, dict):
                for sname, df in sheets.items():
                    # preferir nome baseado em arquivo; adicionar sufixo da aba original se necessário para distinguir
                    candidate = file_based_tab
                    if len(sheets) > 1:
                        # anexar parte da aba original curta para evitar colisão quando múltiplas abas por arquivo
                        candidate = _safe_sheet_name(f"{candidate}_{sname}")[:31]
                    orig = candidate
                    i = 1
                    while candidate in used_sheet_names:
                        suffix = f"_{i}"
                        candidate = (orig[:31-len(suffix)]) + suffix
                        i += 1
                    used_sheet_names.add(candidate)
                    try:
                        df.to_excel(writer, sheet_name=candidate, index=False)
                    except Exception as e:
                        print(f"Erro ao escrever aba '{candidate}' de '{fpath}': {e}")
            else:
                candidate = file_based_tab
                orig = candidate
                i = 1
                while candidate in used_sheet_names:
                    suffix = f"_{i}"
                    candidate = (orig[:31-len(suffix)]) + suffix
                    i += 1
                used_sheet_names.add(candidate)
                try:
                    sheets.to_excel(writer, sheet_name=candidate, index=False)
                except Exception as e:
                    print(f"Erro ao escrever aba '{candidate}' de '{fpath}': {e}")

    print(f"Arquivo aninhado criado em: {out_path}")
    return moved, out_path
