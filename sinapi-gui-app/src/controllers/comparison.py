import os
from copy import copy
import openpyxl
from openpyxl.styles import PatternFill


def col_to_idx(col_letter):
    """Converte letra de coluna (e.g., 'A', 'B') para índice base-zero."""
    if not col_letter or not col_letter.isalpha():
        return None
    return ord(col_letter.upper()) - ord('A')

def compare_workbooks(
    project_path: str, 
    database_path: str, 
    output_dir: str = None,
    project_code_col: str = None,
    project_value_col: str = None,
    db_code_col: str = None,
    db_value_col: str = None
) -> str:
    """Compara 'Curva ABC' do projeto com as bases e gera arquivo de resultado.
    Retorna o caminho do arquivo salvo.
    Replica comportamento atual presente em `ui.py`.
    """
    proj_wb = openpyxl.load_workbook(project_path)
    db_wb = openpyxl.load_workbook(database_path, read_only=True)

    result_wb = openpyxl.Workbook()
    result_wb.remove(result_wb.active)

    proj_ws = proj_wb["Curva ABC"]
    result_ws = result_wb.create_sheet(title="Resultado da Comparação")

    header = [cell.value for cell in proj_ws[1]]
    
    # Determine project column indices
    proj_code_idx = col_to_idx(project_code_col)
    proj_price_idx = col_to_idx(project_value_col)

    if proj_code_idx is None:
        try:
            # proj_code_idx = header.index("Descrição")
            proj_code_idx = header.index("Código")
        except ValueError:
            proj_code_idx = 1  # Fallback to column B
    
    if proj_price_idx is None:
        try:
            # proj_price_idx = header.index("Preço Unitário")
            proj_price_idx = header.index("Valor Unit")
        except ValueError:
            proj_price_idx = 3  # Fallback to column D

    # Determine database column indices
    db_code_idx = col_to_idx(db_code_col)
    db_price_idx = col_to_idx(db_value_col)

    if db_code_idx is None:
        db_code_idx = 1 # Fallback to column B
    
    if db_price_idx is None:
        db_price_idx = 3 # Fallback to column D


    db_data = {}
    for sheet_name in db_wb.sheetnames:
        sheet = db_wb[sheet_name]
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # Ensure row has enough columns
            if len(row) > max(db_code_idx, db_price_idx):
                desc = row[db_code_idx]
                price = row[db_price_idx]
                if isinstance(desc, str):
                    db_data[desc.strip()] = (price, sheet_name)

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    blue_fill = PatternFill(start_color="BDE0FE", end_color="BDE0FE", fill_type="solid")
    gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    for col_idx, cell in enumerate(proj_ws[1], 1):
        result_ws.cell(row=1, column=col_idx, value=cell.value)

    base_header_col = len(header) + 1
    result_ws.cell(row=1, column=base_header_col, value="Planilha da Base")
    result_ws.cell(row=1, column=base_header_col + 1, value="Valor Curva ABC")
    result_ws.cell(row=1, column=base_header_col + 2, value="Valor Base de Dados")
    result_ws.cell(row=1, column=base_header_col + 3, value="Diferença")

    for row_idx, proj_row_data in enumerate(proj_ws.iter_rows(min_row=2, values_only=True), 2):
        for col_idx, cell_value in enumerate(proj_row_data, 1):
            result_ws.cell(row=row_idx, column=col_idx, value=cell_value)

        proj_desc = proj_row_data[proj_code_idx]
        proj_price = proj_row_data[proj_price_idx]

        fill_color = None
        if isinstance(proj_desc, str) and proj_desc.strip() in db_data:
            db_price, db_sheet_name = db_data[proj_desc.strip()]
            result_ws.cell(row=row_idx, column=base_header_col, value=db_sheet_name)
            try:
                proj_price_float = float(proj_price)
                db_price_float = float(db_price)

                result_ws.cell(row=row_idx, column=base_header_col + 1, value=proj_price_float)
                result_ws.cell(row=row_idx, column=base_header_col + 2, value=db_price_float)
                result_ws.cell(row=row_idx, column=base_header_col + 3, value=proj_price_float - db_price_float)

                if proj_price_float < db_price_float:
                    fill_color = green_fill
                elif proj_price_float > db_price_float:
                    fill_color = red_fill
                else:
                    fill_color = blue_fill
            except (ValueError, TypeError):
                fill_color = gray_fill
        else:
            fill_color = gray_fill

        if fill_color:
            for col in range(1, len(proj_row_data) + 1):
                result_ws.cell(row=row_idx, column=col).fill = fill_color

    original_abc_ws = result_wb.create_sheet(title="Curva ABC (Original)")
    for r_idx, row in enumerate(proj_ws.iter_rows(), 1):
        for c_idx, cell in enumerate(row, 1):
            new_cell = original_abc_ws.cell(row=r_idx, column=c_idx, value=cell.value)
            if hasattr(cell, 'has_style') and cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    for db_sheet_name in db_wb.sheetnames:
        db_ws = db_wb[db_sheet_name]
        new_db_ws = result_wb.create_sheet(title=db_sheet_name)
        for r_idx, row in enumerate(db_ws.iter_rows(), 1):
            for c_idx, cell in enumerate(row, 1):
                new_cell = new_db_ws.cell(row=r_idx, column=c_idx, value=cell.value)
                if hasattr(cell, 'has_style') and cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = cell.number_format

    filename = "comparacao_resultados.xlsx"
    if output_dir:
        output_filename = os.path.join(output_dir, filename)
    else:
        output_filename = filename
    result_wb.save(output_filename)
    return os.path.abspath(output_filename)


def validate_project_has_curva(project_path: str) -> bool:
    """Retorna True se o workbook contém a aba 'Curva ABC'."""
    try:
        wb = openpyxl.load_workbook(project_path, read_only=True)
        return "Curva ABC" in wb.sheetnames
    except Exception:
        return False
