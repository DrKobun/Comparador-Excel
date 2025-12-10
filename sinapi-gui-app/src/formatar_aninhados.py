import os
import xlwings as xw

def format_excel_files():
    """
    Automatiza a formatação de planilhas Excel específicas em um diretório.

    Esta função itera sobre todos os arquivos Excel na pasta 
    'Arquivos-SINAPI-SICRO-ORSE/aninhar' na área de trabalho do usuário. 
    Para cada pasta de trabalho, ele encontra planilhas que começam com
    o nome 'SINA-SIN' e executa as seguintes operações:
    1. Exclui as colunas A, B, C, D, E, F, J e L.
    2. Auto-ajusta a largura de todas as colunas restantes para garantir que o conteúdo seja visível.
    
    As alterações são salvas nos arquivos originais.
    
    """
    # Constrói o caminho para o diretório de destino na área de trabalho
    home_dir = os.path.expanduser('~')
    target_dir = os.path.join(home_dir, 'Desktop', 'Arquivos-SINAPI-SICRO-ORSE', 'aninhar')

    if not os.path.isdir(target_dir):
        print(f"Erro: O diretório não foi encontrado em '{target_dir}'")
        print("Por favor, verifique se o caminho está correto.")
        return

    print(f"Procurando por arquivos Excel em: {target_dir}")
    excel_files = [f for f in os.listdir(target_dir) if f.lower().endswith(('.xlsx', '.xls'))]

    if not excel_files:
        print("Nenhum arquivo Excel encontrado no diretório.")
        return
        
    # Usar visible=False faz o Excel rodar em segundo plano
    with xw.App(visible=False) as app:
        for filename in excel_files:
            file_path = os.path.join(target_dir, filename)
            wb = None  # Inicializa wb como None
            try:
                print(f"\nAbrindo arquivo: {filename}...")
                wb = app.books.open(file_path)

                if wb.api.ReadOnly:
                    print(f"  - AVISO: O arquivo '{filename}' está no modo somente leitura. As alterações não serão salvas.")
                
                sheet_formatted = False
                # Itera sobre todas as planilhas
                for sheet in wb.sheets:
                    # Condição para planilhas "SINA-SIN"
                    if sheet.name.startswith('SINA-SIN'):
                        sheet_formatted = True
                        print(f"  - Formatando planilha (SINA-SIN): '{sheet.name}'")
                        
                        # --- Etapa de Limpeza ---
                        try:
                            print("    - Passo 1a: Removendo painéis congelados...")
                            sheet.activate()
                            wb.app.api.ActiveWindow.FreezePanes = False
                        except Exception as e:
                            print(f"    - AVISO ao remover painéis congelados: {e}")
                        try:
                            print("    - Passo 1b: Removendo filtros...")
                            sheet.api.AutoFilterMode = False
                        except Exception as e:
                            print(f"    - AVISO ao remover filtros: {e}")
                        try:
                            print("    - Passo 1c: Desfazendo células mescladas...")
                            sheet.used_range.unmerge()
                        except Exception as e:
                            print(f"    - AVISO ao desfazer mesclagem de células: {e}")
                        try:
                            print("    - Passo 1d: Desprotegendo planilha...")
                            sheet.api.Unprotect()
                        except Exception as e:
                            print(f"    - AVISO ao desproteger: A operação falhou. Erro: {e}")

                        # --- Etapa de Modificação ---
                        try:
                            print("    - Passo 2: Excluindo colunas...")
                            sheet.range('L:L').delete()
                            sheet.range('J:J').delete()
                            sheet.range('F:F').delete()
                            sheet.range('E:E').delete()
                            sheet.range('D:D').delete()
                            sheet.range('C:C').delete()
                            sheet.range('B:B').delete()
                            sheet.range('A:A').delete()
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO na exclusão de colunas.")
                            raise e
                        try:
                            print("    - Passo 3: Excluindo linhas 1-6...")
                            
                            sheet.api.Rows(7).Delete()
                            sheet.api.Rows(6).Delete()
                            sheet.api.Rows(5).Delete()
                            sheet.api.Rows(4).Delete()
                            sheet.api.Rows(3).Delete()
                            sheet.api.Rows(2).Delete()
                            sheet.api.Rows(1).Delete()
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO na exclusão de linhas.")
                            raise e
                        try:
                            print("    - Passo 4: Ajustando a largura das colunas...")
                            sheet.used_range.columns.autofit()
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO no ajuste de colunas.")
                            raise e
                        try:
                            print("    - Passo 5: Ajustando a largura da coluna B para 180...")
                            sheet.range('B:B').column_width = 180
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO ao ajustar a largura da coluna B.")
                            raise e
                        
                        # Passo 6: Converter valores da coluna D para número
                        try:
                            print("    - Passo 6: Convertendo coluna D para número...")
                            last_row = sheet.range('D' + str(sheet.cells.last_cell.row)).end('up').row
                            range_d = sheet.range(f'D1:D{last_row}')
                            
                            values = range_d.value
                            new_values = []
                            if isinstance(values, list): # Se houver mais de uma célula
                                for value in values:
                                    if value is None:
                                        new_values.append(None)
                                        continue
                                    try:
                                        # Trata valores como "R$ 1.234,56"
                                        s_value = str(value).replace('R$ ', '').replace('.', '').replace(',', '.')
                                        new_values.append(float(s_value))
                                    except (ValueError, TypeError):
                                        new_values.append(value) # Mantém o valor se não for conversível
                                range_d.value = [[v] for v in new_values]
                            elif values is not None: # Se for apenas uma célula
                                try:
                                    s_value = str(values).replace('R$ ', '').replace('.', '').replace(',', '.')
                                    range_d.value = float(s_value)
                                except (ValueError, TypeError):
                                    pass # Mantém o valor original se não for conversível
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO ao converter a coluna D para número: {e}")
                            raise e

                        print(f"    - Planilha '{sheet.name}' formatada com sucesso.")

                    elif "EQP" in sheet.name:
                        sheet_formatted = True
                        print(f"  - Formatando planilha (EQP): '{sheet.name}'")
                        
                        # Exclui as colunas C, D, E, F, G, H, I
                        try:
                            print("    - Excluindo colunas C a I...")
                            sheet.range('I:I').delete()
                            sheet.range('H:H').delete()
                            sheet.range('G:G').delete()
                            sheet.range('F:F').delete()
                            sheet.range('E:E').delete()
                            sheet.range('D:D').delete()
                            sheet.range('C:C').delete()
                            
                            sheet.api.Rows(3).Delete()
                            sheet.api.Rows(2).Delete()
                            sheet.api.Rows(1).Delete()
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO na exclusão de colunas.")
                            raise e
                        
                        print(f"    - Planilha '{sheet.name}' formatada com sucesso.")

                    elif sheet.name.startswith('SICRO'):
                        sheet_formatted = True
                        print(f"  - Formatando planilha (SICRO): '{sheet.name}'")
                        
                        # Exclui as linhas 1, 2 e 3
                        try:
                            print("    - Excluindo linhas 1 a 3...")
                            sheet.range('3:3').delete()
                            sheet.range('2:2').delete()
                            sheet.range('1:1').delete()
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO na exclusão de linhas.")
                            raise e
                        
                        print(f"    - Planilha '{sheet.name}' formatada com sucesso.")

                    # Condição para planilhas "SINA-INS"
                    elif sheet.name.startswith('SINA-INS'):
                        sheet_formatted = True
                        print(f"  - Formatando planilha (SINA-INS): '{sheet.name}'")
                        
                        # Exclui as linhas 1-6
                        try:
                            print("    - Excluindo linhas 1-6...")
                            sheet.api.Rows(7).Delete()
                            sheet.api.Rows(6).Delete()
                            sheet.api.Rows(5).Delete()
                            sheet.api.Rows(4).Delete()
                            sheet.api.Rows(3).Delete()
                            sheet.api.Rows(2).Delete()
                            sheet.api.Rows(1).Delete()
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO na exclusão de linhas.")
                            raise e
                        # Remove a linha D
                        try:
                            sheet.range('D:D').delete()
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO na exclusão da linha D.")
                            raise e
                        # Ajusta a largura das colunas B e E
                        try:
                            print("    - Ajustando largura das colunas B (95) e E (18)...")
                            sheet.range('B:B').column_width = 95
                            sheet.range('E:E').column_width = 18
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO ao ajustar a largura das colunas.")
                            raise e

                        # Convertendo a coluna D (originalmente E) para valores numéricos
                        try:
                            print("    - Convertendo coluna D para número...")
                            last_row = sheet.range('D' + str(sheet.cells.last_cell.row)).end('up').row
                            range_d = sheet.range(f'D1:D{last_row}')
                            
                            values = range_d.value
                            new_values = []
                            if isinstance(values, list): # Se houver mais de uma célula
                                for value in values:
                                    if value is None:
                                        new_values.append(None)
                                        continue
                                    try:
                                        # Trata valores como "R$ 1.234,56"
                                        s_value = str(value).replace('R$ ', '').replace('.', '').replace(',', '.')
                                        new_values.append(float(s_value))
                                    except (ValueError, TypeError):
                                        new_values.append(value) # Mantém o valor se não for conversível
                                range_d.value = [[v] for v in new_values]
                            elif values is not None: # Se for apenas uma célula
                                try:
                                    s_value = str(values).replace('R$ ', '').replace('.', '').replace(',', '.')
                                    range_d.value = float(s_value)
                                except (ValueError, TypeError):
                                    pass # Mantém o valor original se não for conversível
                        except Exception as e:
                            print(f"    - ERRO CRÍTICO ao converter a coluna D para número: {e}")
                            raise e

                        print(f"    - Planilha '{sheet.name}' formatada com sucesso.")

                if sheet_formatted:
                    print(f"Salvando alterações em {filename}...")
                    wb.save()
                else:
                    print(f"  - Nenhuma planilha com o prefixo 'SINA-SIN' encontrada em {filename}.")

            except Exception as e:
                print(f"  ERRO: Não foi possível processar o arquivo {filename}. Motivo: {e}")
            finally:
                if wb:
                    wb.close()

    print("\nProcesso concluído.")

if __name__ == '__main__':
    print("------------------------------------------------------------------")
    print("Este script irá modificar arquivos Excel para formatá-los.")
    print("Requisitos:")
    print("1. A biblioteca 'xlwings' deve ser instalada (pip install xlwings).")
    print("2. Microsoft Excel deve estar instalado neste computador.")
    print("3. Os arquivos a serem formatados devem estar em:")
    print(f"   {os.path.join(os.path.expanduser('~'), 'Desktop', 'Arquivos-SINAPI-SICRO-ORSE', 'aninhar')}")
    print("------------------------------------------------------------------\n")
    
    try:
        input("Pressione Enter para iniciar o processo de formatação...")
        format_excel_files()
    except KeyboardInterrupt:
        print("\nProcesso cancelado pelo usuário.")
    
