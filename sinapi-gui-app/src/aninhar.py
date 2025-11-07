import os
import shutil
import zipfile
import re
from datetime import datetime
from typing import List, Tuple, Optional
import pandas as pd

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
        base_dir = os.path.join(desktop_dir, "Sinapi downloads")
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

    # 3) localizar arquivos Excel relevantes (recursivo, inclui conteúdo extraído)
    exts = {'.xls', '.xlsx', '.xlsm', '.xlsb'}
    tokens = ("sintetico", "insumos")
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
        if "desonerado" in low:
            tipo_abbr = "DES"
        elif "naodesonerado" in low or "nao desonerado" in low or "naodesonerado" in low or "não desonerado" in low:
            tipo_abbr = "NDS"
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

