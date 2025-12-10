from tkinter import Tk, StringVar, IntVar, BooleanVar, Label, OptionMenu, Radiobutton, Checkbutton, Button, Toplevel, messagebox, Frame, font
from datetime import datetime
import threading
import importlib
import sys
import os
import webbrowser

# Adiciona o diretório raiz ao sys.path para encontrar o módulo orse
src_dir = os.path.dirname(__file__)
root_dir = os.path.abspath(os.path.join(src_dir, '..', '..'))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

try:
    from orse import funcao_orse
except ImportError:
    # Fallback para quando executado em um contexto diferente
    orse_mod = importlib.import_module("orse")
    funcao_orse = getattr(orse_mod, "funcao_orse")

# tenta import direto; se houver conflito entre dois módulos 'sinapi' no sys.path,
# força carregar o módulo a partir da pasta do src (garante consistência ao rodar main.py)
try:
    from sinapi import (
        gerar_links_sinapi, 
        abrir_links_no_navegador,
        definir_valor_janeiro_2021,
        definir_valor_fevereiro_2021,
        definir_valor_marco_2021,
        definir_valor_abril_2021,
        definir_valor_setembro_2020,
        definir_valor_outubro_2020,
        definir_valor_novembro_2020,
        definir_valor_dezembro_2020,
        definir_valor_maio_2020,
        definir_valor_junho_2020,
        definir_valor_julho_2020,
        definir_valor_agosto_2020,
        definir_valor_janeiro_2020,
        definir_valor_fevereiro_2020,
        definir_valor_marco_2020,
        definir_valor_abril_2020,
        definir_valor_julho_2018,
        definir_valor_agosto_2018,
        definir_valor_janeiro_2019,
        definir_valor_fevereiro_2019,
        definir_valor_marco_2019,
        definir_valor_abril_2019,
        definir_valor_maio_2019,
        definir_valor_junho_2019,
        definir_valor_julho_2019,
        definir_valor_agosto_2019,
        definir_valor_setembro_2019,
        definir_valor_outubro_2019,
        definir_valor_novembro_2019,
        definir_valor_dezembro_2019,
        definir_valor_janeiro_2018,
        definir_valor_fevereiro_2018,
        definir_valor_marco_2018,
        definir_valor_abril_2018,
        definir_valor_maio_2018,
        definir_valor_junho_2018,
        definir_valor_setembro_2018,
        definir_valor_outubro_2018,
        definir_valor_novembro_2018,
        definir_valor_dezembro_2018,
        definir_valor_janeiro_2017,
        definir_valor_fevereiro_2017,
        definir_valor_marco_2017,
        definir_valor_abril_2017,
        definir_valor_maio_2017,
        definir_valor_junho_2017,
        definir_valor_julho_2017,
        definir_valor_agosto_2017,
        definir_valor_setembro_2017,
        definir_valor_outubro_2017,
        definir_valor_novembro_2017,
        definir_valor_dezembro_2017
    )
    from aninhar import aninhar_arquivos, apagar_dados_sinapi
    import sicro
except Exception:
    src_dir = os.path.dirname(__file__)  # pasta src
    if src_dir not in sys.path:
        sys.path.insert(0, src_dir)
    sinapi_mod = importlib.import_module("sinapi")
    aninhar_mod = importlib.import_module("aninhar")
    sicro = importlib.import_module("sicro")
    
    gerar_links_sinapi = getattr(sinapi_mod, "gerar_links_sinapi")
    abrir_links_no_navegador = getattr(sinapi_mod, "abrir_links_no_navegador")
    aninhar_arquivos = getattr(aninhar_mod, "aninhar_arquivos")
    apagar_dados_sinapi = getattr(aninhar_mod, "apagar_dados_sinapi")


from resources.states import estados

class SinapiApp:
    def __init__(self, master):
        self.master = master
        master.title("SINAPI, ORSE e SICRO downloads")
        master.geometry("800x450")
        master.resizable(False, False)

        self.selected_service = StringVar(value="SINAPI")
        current_year = datetime.now().year
        default_year = str(current_year if 2017 <= current_year <= 2024 else 2024)
        self.selected_year = StringVar(value=default_year)
        self.selected_month = StringVar(value=datetime.now().strftime("%m"))
        self.selected_type = IntVar(value=0)
        self.selected_file_type = StringVar(value="Ambos")
        self.selected_states = {estado: IntVar(value=0) for estado in estados}
        self.selected_orse_type = StringVar(value="ambos")
        
        # SICRO related
        self.sicro_links_data = sicro.parse_sicro_links()
        self.selected_sicro_state = StringVar(value=estados[0] if estados else "")
        self.sicro_composicoes = BooleanVar(value=False)
        self.sicro_equipamentos_desonerado = BooleanVar(value=False)
        self.sicro_equipamentos = BooleanVar(value=False)
        self.sicro_materiais = BooleanVar(value=False)


        # Vars for months 1-4 checkboxes
        self.jan_2021 = BooleanVar(value=False)
        self.feb_2021 = BooleanVar(value=False)
        self.mar_2021 = BooleanVar(value=False)
        self.apr_2021 = BooleanVar(value=False)

        self.sep_2020 = BooleanVar(value=False)
        self.oct_2020 = BooleanVar(value=False)
        self.nov_2020 = BooleanVar(value=False)
        self.dec_2020 = BooleanVar(value=False)

        self.may_2020 = BooleanVar(value=False)
        self.jun_2020 = BooleanVar(value=False)
        self.jul_2020 = BooleanVar(value=False)
        self.aug_2020 = BooleanVar(value=False)

        self.jan_2020 = BooleanVar(value=False)
        self.feb_2020 = BooleanVar(value=False)
        self.mar_2020 = BooleanVar(value=False)
        self.apr_2020 = BooleanVar(value=False)

        self.jul_2018 = BooleanVar(value=False)
        self.aug_2018 = BooleanVar(value=False)

        self.jan_2019 = BooleanVar(value=False)
        self.feb_2019 = BooleanVar(value=False)
        self.mar_2019 = BooleanVar(value=False)
        self.apr_2019 = BooleanVar(value=False)
        self.may_2019 = BooleanVar(value=False)
        self.jun_2019 = BooleanVar(value=False)
        self.jul_2019 = BooleanVar(value=False)
        self.aug_2019 = BooleanVar(value=False)
        self.sep_2019 = BooleanVar(value=False)
        self.oct_2019 = BooleanVar(value=False)
        self.nov_2019 = BooleanVar(value=False)
        self.dec_2019 = BooleanVar(value=False)

        self.jan_2018 = BooleanVar(value=False)
        self.feb_2018 = BooleanVar(value=False)
        self.mar_2018 = BooleanVar(value=False)
        self.apr_2018 = BooleanVar(value=False)
        self.may_2018 = BooleanVar(value=False)
        self.jun_2018 = BooleanVar(value=False)
        self.sep_2018 = BooleanVar(value=False)
        self.oct_2018 = BooleanVar(value=False)
        self.nov_2018 = BooleanVar(value=False)
        self.dec_2018 = BooleanVar(value=False)

        self.jan_2017 = BooleanVar(value=False)
        self.feb_2017 = BooleanVar(value=False)
        self.mar_2017 = BooleanVar(value=False)
        self.apr_2017 = BooleanVar(value=False)
        self.may_2017 = BooleanVar(value=False)
        self.jun_2017 = BooleanVar(value=False)
        self.jul_2017 = BooleanVar(value=False)
        self.aug_2017 = BooleanVar(value=False)
        self.sep_2017 = BooleanVar(value=False)
        self.oct_2017 = BooleanVar(value=False)
        self.nov_2017 = BooleanVar(value=False)
        self.dec_2017 = BooleanVar(value=False)

        self.months_1_to_4_frame = None
        self.month_menu = None
        self.year_menu = None
        self.orse_widgets = None
        self.sinapi_widgets = None
        self.sicro_widgets = None
        self.baixar_button = None
        self.aninhar_button = None
        self.add_state_button = None
        self.apagar_button = None

        self.create_widgets()

        self.selected_year.trace_add("write", self._on_year_change)
        self.selected_month.trace_add("write", self._on_month_change)
        self.selected_service.trace_add("write", self._on_service_change)
        self.selected_sicro_state.trace_add("write", self._on_sicro_params_change)
        self.selected_year.trace_add("write", self._on_sicro_params_change)
        
        self._on_year_change()
        self._on_service_change()

    def _on_service_change(self, *args):
        service = self.selected_service.get()
        
        # Hide all service-specific widgets initially
        if self.sinapi_widgets: self.sinapi_widgets.pack_forget()
        if self.orse_widgets: self.orse_widgets.pack_forget()
        if self.sicro_widgets: self.sicro_widgets.pack_forget()
        
        # Configure UI based on selected service
        if service == "ORSE":
            years = [str(y) for y in range(2021, 2025)]
            if self.orse_widgets: self.orse_widgets.pack(pady=2)
            if self.baixar_button: self.baixar_button.config(command=self.execute_orse)
            if self.aninhar_button: self.aninhar_button.pack_forget()
            if self.add_state_button: self.add_state_button.pack_forget()
            if self.apagar_button: self.apagar_button.pack_forget()

        elif service == "SICRO":
            years = [str(y) for y in range(2017, datetime.now().year + 1)]
            if self.sicro_widgets: self.sicro_widgets.pack(pady=2)
            if self.baixar_button: self.baixar_button.config(command=self.execute_sicro)
            if self.aninhar_button: self.aninhar_button.pack(side='left', padx=5)
            if self.add_state_button: self.add_state_button.pack_forget()
            if self.apagar_button: self.apagar_button.pack(side='right', padx=5)
            self._on_sicro_params_change()

        else:  # SINAPI
            years = [str(y) for y in range(2017, 2025)]
            if self.sinapi_widgets: self.sinapi_widgets.pack(pady=2)
            if self.baixar_button: self.baixar_button.config(command=self.execute_sinapi)
            if self.aninhar_button: self.aninhar_button.pack(side='left', padx=5)
            if self.add_state_button: self.add_state_button.pack(side='left', padx=5)
            if self.apagar_button: self.apagar_button.pack(side='right', padx=5)

        # Update year dropdown
        menu = self.year_menu["menu"]
        menu.delete(0, "end")
        for year in years:
            menu.add_command(label=year, command=lambda value=year: self.selected_year.set(value))
        
        if self.selected_year.get() not in years:
            self.selected_year.set(years[-1] if years else "")
        
        # Atualiza os meses com base no serviço selecionado
        self._on_year_change()

    def _on_year_change(self, *args):
        year = self.selected_year.get()
        service = self.selected_service.get()

        if service == "SICRO":
            self._on_sicro_params_change()
            return
        
        new_months = []
        if service == "ORSE":
            new_months = [f"{m:02d}" for m in range(1, 13)]
        else:  # Lógica para SINAPI/outros
            months_2021 = ["1 a 4"] + [f"{m:02d}" for m in range(5, 13)]
            months_2020 = ["1 a 4", "5 a 8", "9 a 12"]
            months_2019 = ["1 a 4", "5 a 8", "9 a 12"]
            months_2018 = ["1 a 6", "7 e 8", "9 a 12"]
            months_2017 = ["1 a 6", "7 a 12"]
            
            if year == "2021": new_months = months_2021
            elif year == "2020": new_months = months_2020
            elif year == "2019": new_months = months_2019
            elif year == "2018": new_months = months_2018
            elif year == "2017": new_months = months_2017
            else: new_months = [f"{m:02d}" for m in range(1, 13)]

        current_month = self.selected_month.get()

        if self.month_menu:
            menu = self.month_menu["menu"]
            menu.delete(0, "end")
            for month in new_months:
                menu.add_command(label=month, command=lambda value=month: self.selected_month.set(value))
                
            if current_month not in new_months:
                self.selected_month.set(new_months[0])
            else:
                self.selected_month.set(current_month)

    def _on_month_change(self, *args):
        year = self.selected_year.get()
        month = self.selected_month.get()
        service = self.selected_service.get()

        # Limpa os checkboxes de meses específicos
        if self.months_1_to_4_frame:
            for widget in self.months_1_to_4_frame.winfo_children():
                widget.destroy()

        # A lógica de checkboxes só se aplica a SINAPI
        if service != "SINAPI":
            return

        if year == "2021" and month == "1 a 4":
            Checkbutton(self.months_1_to_4_frame, text="Mês 1", variable=self.jan_2021, command=lambda: definir_valor_janeiro_2021(self.jan_2021.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 2", variable=self.feb_2021, command=lambda: definir_valor_fevereiro_2021(self.feb_2021.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 3", variable=self.mar_2021, command=lambda: definir_valor_marco_2021(self.mar_2021.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 4", variable=self.apr_2021, command=lambda: definir_valor_abril_2021(self.apr_2021.get())).pack(side='left')
        
        elif year == "2020" and month == "9 a 12":
            Checkbutton(self.months_1_to_4_frame, text="Mês 9", variable=self.sep_2020, command=lambda: definir_valor_setembro_2020(self.sep_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 10", variable=self.oct_2020, command=lambda: definir_valor_outubro_2020(self.oct_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 11", variable=self.nov_2020, command=lambda: definir_valor_novembro_2020(self.nov_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 12", variable=self.dec_2020, command=lambda: definir_valor_dezembro_2020(self.dec_2020.get())).pack(side='left')

        elif year == "2020" and month == "5 a 8":
            Checkbutton(self.months_1_to_4_frame, text="Mês 5", variable=self.may_2020, command=lambda: definir_valor_maio_2020(self.may_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 6", variable=self.jun_2020, command=lambda: definir_valor_junho_2020(self.jun_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 7", variable=self.jul_2020, command=lambda: definir_valor_julho_2020(self.jul_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 8", variable=self.aug_2020, command=lambda: definir_valor_agosto_2020(self.aug_2020.get())).pack(side='left')

        elif year == "2020" and month == "1 a 4":
            Checkbutton(self.months_1_to_4_frame, text="Mês 1", variable=self.jan_2020, command=lambda: definir_valor_janeiro_2020(self.jan_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 2", variable=self.feb_2020, command=lambda: definir_valor_fevereiro_2020(self.feb_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 3", variable=self.mar_2020, command=lambda: definir_valor_marco_2020(self.mar_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 4", variable=self.apr_2020, command=lambda: definir_valor_abril_2020(self.apr_2020.get())).pack(side='left')

        elif year == "2018" and month == "7 e 8":
            Checkbutton(self.months_1_to_4_frame, text="Mês 7", variable=self.jul_2018, command=lambda: definir_valor_julho_2018(self.jul_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 8", variable=self.aug_2018, command=lambda: definir_valor_agosto_2018(self.aug_2018.get())).pack(side='left')

        elif year == "2019" and month == "1 a 4":
            Checkbutton(self.months_1_to_4_frame, text="Mês 1", variable=self.jan_2019, command=lambda: definir_valor_janeiro_2019(self.jan_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 2", variable=self.feb_2019, command=lambda: definir_valor_fevereiro_2019(self.feb_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 3", variable=self.mar_2019, command=lambda: definir_valor_marco_2019(self.mar_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 4", variable=self.apr_2019, command=lambda: definir_valor_abril_2019(self.apr_2019.get())).pack(side='left')
        
        elif year == "2019" and month == "5 a 8":
            Checkbutton(self.months_1_to_4_frame, text="Mês 5", variable=self.may_2019, command=lambda: definir_valor_maio_2019(self.may_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 6", variable=self.jun_2019, command=lambda: definir_valor_junho_2019(self.jun_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 7", variable=self.jul_2019, command=lambda: definir_valor_julho_2019(self.jul_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 8", variable=self.aug_2019, command=lambda: definir_valor_agosto_2019(self.aug_2019.get())).pack(side='left')

        elif year == "2019" and month == "9 a 12":
            Checkbutton(self.months_1_to_4_frame, text="Mês 9", variable=self.sep_2019, command=lambda: definir_valor_setembro_2019(self.sep_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 10", variable=self.oct_2019, command=lambda: definir_valor_outubro_2019(self.oct_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 11", variable=self.nov_2019, command=lambda: definir_valor_novembro_2019(self.nov_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 12", variable=self.dec_2019, command=lambda: definir_valor_dezembro_2019(self.dec_2019.get())).pack(side='left')

        elif year == "2018" and month == "1 a 6":
            Checkbutton(self.months_1_to_4_frame, text="Mês 1", variable=self.jan_2018, command=lambda: definir_valor_janeiro_2018(self.jan_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 2", variable=self.feb_2018, command=lambda: definir_valor_fevereiro_2018(self.feb_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 3", variable=self.mar_2018, command=lambda: definir_valor_marco_2018(self.mar_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 4", variable=self.apr_2018, command=lambda: definir_valor_abril_2018(self.apr_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 5", variable=self.may_2018, command=lambda: definir_valor_maio_2018(self.may_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 6", variable=self.jun_2018, command=lambda: definir_valor_junho_2018(self.jun_2018.get())).pack(side='left')

        elif year == "2018" and month == "9 a 12":
            Checkbutton(self.months_1_to_4_frame, text="Mês 9", variable=self.sep_2018, command=lambda: definir_valor_setembro_2018(self.sep_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 10", variable=self.oct_2018, command=lambda: definir_valor_outubro_2018(self.oct_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 11", variable=self.nov_2018, command=lambda: definir_valor_novembro_2018(self.nov_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 12", variable=self.dec_2018, command=lambda: definir_valor_dezembro_2018(self.dec_2018.get())).pack(side='left')

        elif year == "2017" and month == "1 a 6":
            Checkbutton(self.months_1_to_4_frame, text="Mês 1", variable=self.jan_2017, command=lambda: definir_valor_janeiro_2017(self.jan_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 2", variable=self.feb_2017, command=lambda: definir_valor_fevereiro_2017(self.feb_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 3", variable=self.mar_2017, command=lambda: definir_valor_marco_2017(self.mar_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 4", variable=self.apr_2017, command=lambda: definir_valor_abril_2017(self.apr_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 5", variable=self.may_2017, command=lambda: definir_valor_maio_2017(self.may_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 6", variable=self.jun_2017, command=lambda: definir_valor_junho_2017(self.jun_2017.get())).pack(side='left')

        elif year == "2017" and month == "7 a 12":
            Checkbutton(self.months_1_to_4_frame, text="Mês 7", variable=self.jul_2017, command=lambda: definir_valor_julho_2017(self.jul_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 8", variable=self.aug_2017, command=lambda: definir_valor_agosto_2017(self.aug_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 9", variable=self.sep_2017, command=lambda: definir_valor_setembro_2017(self.sep_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 10", variable=self.oct_2017, command=lambda: definir_valor_outubro_2017(self.oct_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 11", variable=self.nov_2017, command=lambda: definir_valor_novembro_2017(self.nov_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 12", variable=self.dec_2017, command=lambda: definir_valor_dezembro_2017(self.dec_2017.get())).pack(side='left')


    def create_widgets(self):
        main_frame = Frame(self.master)
        main_frame.pack(fill='both', expand=True)

        # --- Botões ---
        bottom_frame = Frame(main_frame)
        bottom_frame.pack(side='bottom', fill='x', pady=10)

        self.baixar_button = Button(bottom_frame, text="Baixar", command=self.execute_sinapi)
        self.baixar_button.pack(side='left', padx=5)
        
        self.aninhar_button = Button(bottom_frame, text="Juntar", command=self.execute_aninhar)
        self.aninhar_button.pack(side='left', padx=5)
        
        self.add_state_button = Button(bottom_frame, text="+1", command=self.add_state)
        self.add_state_button.pack(side='left', padx=5)
        
        self.apagar_button = Button(bottom_frame, text="apagar dados", command=apagar_dados_sinapi)
        self.apagar_button.pack(side='right', padx=5)

        # --- Conteúdo Principal ---
        content_frame = Frame(main_frame)
        content_frame.pack(side='top', fill='x')

        Label(content_frame, text="Selecione a base de dados:").pack()
        OptionMenu(content_frame, self.selected_service, "SINAPI", "SICRO", "ORSE").pack()

        Label(content_frame, text="Selecione a data:").pack()
        date_frame = Frame(content_frame)
        date_frame.pack(pady=2)

        years = [str(y) for y in range(2017, 2025)]
        self.year_menu = OptionMenu(date_frame, self.selected_year, *years)
        self.year_menu.pack(side='left', padx=(0, 6))
        
        self.month_menu = OptionMenu(date_frame, self.selected_month, "")
        self.month_menu.pack(side='left')

        self.months_1_to_4_frame = Frame(content_frame)
        self.months_1_to_4_frame.pack()

        # --- Widgets SINAPI ---
        self.sinapi_widgets = Frame(content_frame)
        # self.sinapi_widgets.pack(pady=2) # Controlado por _on_service_change

        radio_frame = Frame(self.sinapi_widgets)
        radio_frame.pack(pady=2)

        type_frame = Frame(radio_frame)
        type_frame.pack(side='left', padx=10)
        
        Label(type_frame, text="Tipo para baixar?").pack(anchor='w')
        Radiobutton(type_frame, text="Ambos", variable=self.selected_type, value=0).pack(anchor='w')
        Radiobutton(type_frame, text="Desonerado", variable=self.selected_type, value=1).pack(anchor='w')
        Radiobutton(type_frame, text="Não Desonerado", variable=self.selected_type, value=2).pack(anchor='w')

        file_type_frame = Frame(radio_frame)
        file_type_frame.pack(side='left', padx=10)

        Label(file_type_frame, text="Quais arquivos juntar?").pack(anchor='w')
        Radiobutton(file_type_frame, text="Ambos", variable=self.selected_file_type, value="Ambos").pack(anchor='w')
        Radiobutton(file_type_frame, text="Insumos", variable=self.selected_file_type, value="Insumos").pack(anchor='w')
        Radiobutton(file_type_frame, text="Sintéticos", variable=self.selected_file_type, value="Sintetico").pack(anchor='w')

        Label(self.sinapi_widgets, text="Selecione os estados:").pack()

        num_rows = 3
        total = len(estados)
        chunk = (total + num_rows - 1) // num_rows
        rows = [Frame(self.sinapi_widgets) for _ in range(num_rows)]
        for r in rows:
            r.pack(fill='x', padx=8, pady=2)

        for i, estado in enumerate(estados):
            row_idx = min(i // chunk, num_rows - 1)
            cb = Checkbutton(rows[row_idx], text=estado, variable=self.selected_states[estado])
            cb.pack(side='left', anchor='w', padx=4, pady=2)

        # texto de link clicável:
        link_font = font.Font(size=10, underline=True)
        link_label = Label(self.sinapi_widgets, text="site de downloads da CAIXA (SINAPI)", fg="blue", cursor="hand2")
        link_label.pack()
        link_label.bind("<Button-1>", self.open_link)
        link_label.config(font=link_font)

        # --- Widgets ORSE ---
        self.orse_widgets = Frame(content_frame)
        # Não fazer o pack inicial

        orse_type_frame = Frame(self.orse_widgets)
        orse_type_frame.pack(pady=5)
        
        Label(orse_type_frame, text="Qual tipo baixar?").pack(anchor='w')
        Radiobutton(orse_type_frame, text="Ambos", variable=self.selected_orse_type, value="ambos").pack(anchor='w')
        Radiobutton(orse_type_frame, text="Insumos", variable=self.selected_orse_type, value="insumos").pack(anchor='w')
        Radiobutton(orse_type_frame, text="Serviços", variable=self.selected_orse_type, value="servicos").pack(anchor='w')

        # --- Widgets SICRO ---
        self.sicro_widgets = Frame(content_frame)

        sicro_checkbox_frame = Frame(self.sicro_widgets)
        sicro_checkbox_frame.pack(pady=5)
        
        Label(sicro_checkbox_frame, text="Quais planilhas juntar?").pack(anchor='w')
        Checkbutton(sicro_checkbox_frame, text="Composições de Custos", variable=self.sicro_composicoes).pack(anchor='w')
        Checkbutton(sicro_checkbox_frame, text="Equipamentos (com desoneração)", variable=self.sicro_equipamentos_desonerado).pack(anchor='w')
        Checkbutton(sicro_checkbox_frame, text="Equipamentos", variable=self.sicro_equipamentos).pack(anchor='w')
        Checkbutton(sicro_checkbox_frame, text="Materiais", variable=self.sicro_materiais).pack(anchor='w')

        sicro_state_frame = Frame(self.sicro_widgets)
        sicro_state_frame.pack(pady=5)
        Label(sicro_state_frame, text="Selecione o estado:").pack(anchor='w')
        OptionMenu(sicro_state_frame, self.selected_sicro_state, *estados).pack(anchor='w')

    def open_link(self, event):
        webbrowser.open_new_tab("https://www.caixa.gov.br/site/paginas/downloads.aspx")



    def execute_aninhar(self):
        tipo_arquivo = self.selected_file_type.get()
        sicro_composicoes = self.sicro_composicoes.get()
        sicro_equipamentos_desonerado = self.sicro_equipamentos_desonerado.get()
        sicro_equipamentos = self.sicro_equipamentos.get()
        sicro_materiais = self.sicro_materiais.get()
        
        threading.Thread(
            target=aninhar_arquivos, 
            args=(
                None, 
                tipo_arquivo,
                sicro_composicoes,
                sicro_equipamentos_desonerado,
                sicro_equipamentos,
                sicro_materiais
            ), 
            daemon=True
        ).start()

    def execute_sicro(self):
        state = self.selected_sicro_state.get()
        year = self.selected_year.get()
        month = self.selected_month.get()

        if not all([state, year, month]):
            messagebox.showwarning("Aviso", "Por favor, selecione estado, ano e mês.")
            return

        link = sicro.get_sicro_link(self.sicro_links_data, state, year, month)
        
        if link:
            threading.Thread(target=abrir_links_no_navegador, args=([link],), daemon=True).start()
        else:
            messagebox.showerror("Erro", f"Link de download não encontrado para {state}/{year}/{month}.")

    def _on_sicro_params_change(self, *args):
        if self.selected_service.get() != "SICRO":
            return

        state = self.selected_sicro_state.get()
        year = self.selected_year.get()
        
        new_months = sicro.get_available_months(self.sicro_links_data, state, year)
        
        current_month = self.selected_month.get()
        menu = self.month_menu["menu"]
        menu.delete(0, "end")
        
        if new_months:
            for month in new_months:
                menu.add_command(label=month, command=lambda value=month: self.selected_month.set(value))
            
            if current_month not in new_months:
                self.selected_month.set(new_months[0])
        else:
            self.selected_month.set("")

    def execute_orse(self):
        try:
            ano = int(self.selected_year.get())
            mes = int(self.selected_month.get())
            tipo = self.selected_orse_type.get()

            self.baixar_button.config(state="disabled")
            
            threading.Thread(
                target=self._run_orse_and_reenable, 
                args=(ano, mes, tipo), 
                daemon=True
            ).start()

        except ValueError:
            messagebox.showerror("Erro de Entrada", "Ano e Mês devem ser valores numéricos.")
            self.baixar_button.config(state="normal")
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro: {e}")
            self.baixar_button.config(state="normal")

    def _run_orse_and_reenable(self, ano, mes, tipo):
        try:
            funcao_orse(ano, mes, tipo)
            # messagebox.showinfo("Sucesso", "Download do ORSE concluído com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro na Execução", f"Falha ao executar o download do ORSE: {e}")
        finally:
            self.master.after(0, lambda: self.baixar_button.config(state="normal"))

    def execute_sinapi(self, estado: str = None):
        ano_str = self.selected_year.get()
        mes_str = self.selected_month.get()
        
        try:
            ano = int(ano_str)
        except ValueError:
            messagebox.showerror("Erro", "Ano inválido.")
            return

        tipo = ["Ambos", "Desonerado", "NaoDesonerado"][self.selected_type.get()]

        if estado:
            target_states = [estado]
        else:
            target_states = [s for s, v in self.selected_states.items() if v.get() == 1]

        if not target_states:
            messagebox.showwarning("Aviso", "Escolha pelo menos um estado.")
            return

        if ano == 2021 and mes_str == "1 a 4":
            selected_months_2021 = []
            if self.jan_2021.get(): selected_months_2021.append(1)
            if self.feb_2021.get(): selected_months_2021.append(2)
            if self.mar_2021.get(): selected_months_2021.append(3)
            if self.apr_2021.get(): selected_months_2021.append(4)

            if not selected_months_2021:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 1 a 4.")
                return
            
            first_selected_month = selected_months_2021[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2020 and mes_str == "9 a 12":
            selected_months_2020 = []
            if self.sep_2020.get(): selected_months_2020.append(9)
            if self.oct_2020.get(): selected_months_2020.append(10)
            if self.nov_2020.get(): selected_months_2020.append(11)
            if self.dec_2020.get(): selected_months_2020.append(12)

            if not selected_months_2020:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 9 a 12.")
                return

            first_selected_month = selected_months_2020[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2020 and mes_str == "5 a 8":
            selected_months_2020 = []
            if self.may_2020.get(): selected_months_2020.append(5)
            if self.jun_2020.get(): selected_months_2020.append(6)
            if self.jul_2020.get(): selected_months_2020.append(7)
            if self.aug_2020.get(): selected_months_2020.append(8)

            if not selected_months_2020:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 5 a 8.")
                return

            first_selected_month = selected_months_2020[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2020 and mes_str == "1 a 4":
            selected_months_2020 = []
            if self.jan_2020.get(): selected_months_2020.append(1)
            if self.feb_2020.get(): selected_months_2020.append(2)
            if self.mar_2020.get(): selected_months_2020.append(3)
            if self.apr_2020.get(): selected_months_2020.append(4)

            if not selected_months_2020:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 1 a 4.")
                return

            first_selected_month = selected_months_2020[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2018 and mes_str == "7 e 8":
            selected_months_2018 = []
            if self.jul_2018.get(): selected_months_2018.append(7)
            if self.aug_2018.get(): selected_months_2018.append(8)

            if not selected_months_2018:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 7 e 8.")
                return

            first_selected_month = selected_months_2018[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2019 and mes_str == "1 a 4":
            selected_months_2019 = []
            if self.jan_2019.get(): selected_months_2019.append(1)
            if self.feb_2019.get(): selected_months_2019.append(2)
            if self.mar_2019.get(): selected_months_2019.append(3)
            if self.apr_2019.get(): selected_months_2019.append(4)

            if not selected_months_2019:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 1 a 4.")
                return

            first_selected_month = selected_months_2019[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2019 and mes_str == "5 a 8":
            selected_months_2019 = []
            if self.may_2019.get(): selected_months_2019.append(5)
            if self.jun_2019.get(): selected_months_2019.append(6)
            if self.jul_2019.get(): selected_months_2019.append(7)
            if self.aug_2019.get(): selected_months_2019.append(8)

            if not selected_months_2019:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 5 a 8.")
                return

            first_selected_month = selected_months_2019[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2019 and mes_str == "9 a 12":
            selected_months_2019 = []
            if self.sep_2019.get(): selected_months_2019.append(9)
            if self.oct_2019.get(): selected_months_2019.append(10)
            if self.nov_2019.get(): selected_months_2019.append(11)
            if self.dec_2019.get(): selected_months_2019.append(12)

            if not selected_months_2019:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 9 a 12.")
                return

            first_selected_month = selected_months_2019[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2018 and mes_str == "1 a 6":
            selected_months_2018 = []
            if self.jan_2018.get(): selected_months_2018.append(1)
            if self.feb_2018.get(): selected_months_2018.append(2)
            if self.mar_2018.get(): selected_months_2018.append(3)
            if self.apr_2018.get(): selected_months_2018.append(4)
            if self.may_2018.get(): selected_months_2018.append(5)
            if self.jun_2018.get(): selected_months_2018.append(6)

            if not selected_months_2018:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 1 a 6.")
                return

            first_selected_month = selected_months_2018[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2018 and mes_str == "9 a 12":
            selected_months_2018 = []
            if self.sep_2018.get(): selected_months_2018.append(9)
            if self.oct_2018.get(): selected_months_2018.append(10)
            if self.nov_2018.get(): selected_months_2018.append(11)
            if self.dec_2018.get(): selected_months_2018.append(12)

            if not selected_months_2018:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 9 a 12.")
                return

            first_selected_month = selected_months_2018[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2017 and mes_str == "1 a 6":
            selected_months_2017 = []
            if self.jan_2017.get(): selected_months_2017.append(1)
            if self.feb_2017.get(): selected_months_2017.append(2)
            if self.mar_2017.get(): selected_months_2017.append(3)
            if self.apr_2017.get(): selected_months_2017.append(4)
            if self.may_2017.get(): selected_months_2017.append(5)
            if self.jun_2017.get(): selected_months_2017.append(6)

            if not selected_months_2017:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 1 a 6.")
                return

            first_selected_month = selected_months_2017[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        elif ano == 2017 and mes_str == "7 a 12":
            selected_months_2017 = []
            if self.jul_2017.get(): selected_months_2017.append(7)
            if self.aug_2017.get(): selected_months_2017.append(8)
            if self.sep_2017.get(): selected_months_2017.append(9)
            if self.oct_2017.get(): selected_months_2017.append(10)
            if self.nov_2017.get(): selected_months_2017.append(11)
            if self.dec_2017.get(): selected_months_2017.append(12)

            if not selected_months_2017:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 7 a 12.")
                return

            first_selected_month = selected_months_2017[0]
            for st in target_states:
                links = gerar_links_sinapi(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

        else:
            try:
                mes = int(mes_str)
            except ValueError:
                messagebox.showerror("Erro", "Mês inválido.")
                return
            
            for st in target_states:
                links = gerar_links_sinapi(ano, mes, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_no_navegador, args=(links,), daemon=True).start()

    def add_state(self):
        # abre outra instância da mesma janela em Toplevel
        new_win = Toplevel(self.master)
        SinapiApp(new_win)

def run_app():
    root = Tk()
    app = SinapiApp(root)
    root.mainloop()

if __name__ == "__main__":
    run_app()