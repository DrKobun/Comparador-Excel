from tkinter import Tk, StringVar, IntVar, BooleanVar, Label, OptionMenu, Radiobutton, Checkbutton, Button, Toplevel, messagebox, Frame, font, filedialog, Menu, Menubutton, Entry
from tkinter.ttk import Combobox
from datetime import datetime
import threading
import os
import webbrowser

from controllers import (
    parse_sicro_links,
    get_available_months,
    get_sicro_link,
    gerar_links_sync,
    abrir_links_sync,
    funcao_orse_sync,
    start_aninhar,
    format_excel_files_sync,
    compare_workbooks,
    validate_project_has_curva,
    apagar_dados_sync,
)

import importlib
try:
    import sinapi
except Exception:
    sinapi = importlib.import_module("sinapi")

import aninhar
import formatar_aninhados
# Backwards-compat aliases for existing code paths in this file
# (aliases removed) use controller functions directly


from resources.states import estados

class SinapiApp:
    def __init__(self, master):
        self.master = master
        master.title("SINAPI, ORSE e SICRO downloads")
        master.geometry("800x600")
        master.resizable(False, False)

        self.selected_service = StringVar(value="SINAPI")
        current_year = datetime.now().year
        default_year = str(current_year if 2017 <= current_year <= 2024 else 2024)
        self.selected_year = StringVar(value=default_year)
        self.selected_month = StringVar(value=datetime.now().strftime("%m"))
        self.selected_type = IntVar(value=0)
        self.selected_file_type = StringVar(value="Ambos")
        self.selected_states = {estado: IntVar(value=0) for estado in estados}
        self.select_all_states_var = BooleanVar(value=False)
        self.selected_orse_type = StringVar(value="ambos")
        
        # SICRO related
        self.sicro_links_data = parse_sicro_links()
        
        self.selected_sicro_states = {estado: IntVar(value=0) for estado in estados}
        self.select_all_sicro_states_var = BooleanVar(value=False)
        
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
        # self.add_state_button = None
        self.apagar_button = None
        self.formatar_button = None
        self.saida_btn = None
        self.custom_aninhar_path = StringVar()
        self.custom_formatar_path = StringVar()
        self.custom_comparison_path = StringVar()
        
                
        
        # Radio buttons refs for hide/show logic
        self.rb_ambos = None
        self.rb_desonerado = None
        self.rb_nao_desonerado = None
        
        # New variables for file comparison
        self.project_file_path = StringVar()
        self.database_file_path = StringVar()
        self.project_full_path = None
        self.database_full_path = None
        
        # New vars for comparison columns
        self.project_code_col = StringVar()
        self.project_value_col = StringVar()
        self.database_code_col = StringVar()
        self.database_value_col = StringVar()

        self.vcmd = (self.master.register(self.validate_alpha), '%P')
        
        self.create_widgets()
        
        self.selected_year.trace_add("write", self._on_year_change)
        self.selected_month.trace_add("write", self._on_month_change)
        self.selected_service.trace_add("write", self._on_service_change)
        self._on_year_change()
        self._on_service_change()

    def validate_alpha(self, P):
        if P == "" or P.isalpha():
            return True
        return False

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
            if self.formatar_button: self.formatar_button.pack_forget()
            if self.saida_btn: self.saida_btn.pack_forget()
            # if self.add_state_button: self.add_state_button.pack_forget()
            if self.apagar_button: self.apagar_button.pack_forget()

        elif service == "SICRO":
            years = [str(y) for y in range(2017, datetime.now().year + 1)]
            if self.sicro_widgets: self.sicro_widgets.pack(pady=2)
            if self.baixar_button: self.baixar_button.config(command=self.execute_sicro)
            if self.aninhar_button: self.aninhar_button.pack(side='left', padx=5)
            if self.formatar_button: self.formatar_button.pack(side='left', padx=5)
            if self.saida_btn: self.saida_btn.pack(side='left', padx=5)
            # if self.add_state_button: self.add_state_button.pack_forget()
            if self.apagar_button: self.apagar_button.pack(side='right', padx=5)

        else:  # SINAPI
            years = [str(y) for y in range(2017, 2025)]
            if self.sinapi_widgets: self.sinapi_widgets.pack(pady=2)
            if self.baixar_button: self.baixar_button.config(command=self.execute_sinapi)
            if self.aninhar_button: self.aninhar_button.pack(side='left', padx=5)
            if self.formatar_button: self.formatar_button.pack(side='left', padx=5)
            if self.saida_btn: self.saida_btn.pack(side='left', padx=5)
            # if self.add_state_button: self.add_state_button.pack(side='left', padx=5)
            if self.apagar_button: self.apagar_button.pack(side='right', padx=5)

        # Update year combobox values
        try:
            self.year_menu['values'] = years
        except Exception:
            pass

        if self.selected_year.get() not in years:
            self.selected_year.set(years[-1] if years else "")
        try:
            self.year_menu.set(self.selected_year.get())
        except Exception:
            pass
        
        # Atualiza os meses com base no serviço selecionado
        self._on_year_change()

    def _on_year_change(self, *args):
        year = self.selected_year.get()
        service = self.selected_service.get()

        new_months = []
        if service == "SICRO":
            # Prefer months from the selected SICRO states; if none selected, use union across all states
            selected_states = [s for s, v in self.selected_sicro_states.items() if v.get() == 1]
            if selected_states:
                # use months from first selected state
                new_months = get_available_months(self.sicro_links_data, selected_states[0], year)
            else:
                # union months across all states available in sicro_links_data for the year
                months_set = set()
                for st in (self.sicro_links_data.keys() if self.sicro_links_data else []):
                    for m in get_available_months(self.sicro_links_data, st, year):
                        months_set.add(m)
                # sort by month number, keeping regular before revisado
                def _month_sort_key(x):
                    try:
                        num = int(x.split(' ')[0])
                    except Exception:
                        num = 0
                    return (num, 'revisado' in x)
                new_months = sorted(months_set, key=_month_sort_key)
        elif service == "ORSE":
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
            # populate Combobox values (works in dev and in PyInstaller exe)
            try:
                self.month_menu['values'] = new_months
            except Exception:
                pass

            if current_month not in new_months:
                self.selected_month.set(new_months[0] if new_months else "")
            else:
                self.selected_month.set(current_month)

            try:
                # ensure displayed value matches variable
                self.month_menu.set(self.selected_month.get())
            except Exception:
                pass

    def _on_month_change(self, *args):
        year = self.selected_year.get()
        month = self.selected_month.get()
        service = self.selected_service.get()

        # Logic to hide/show "Ambos" radio button
        is_grouped_month = " a " in month or " e " in month
        
        if service == "SINAPI" and hasattr(self, 'rb_ambos') and self.rb_ambos:
            is_ambos_visible = self.rb_ambos in self.rb_ambos.master.pack_slaves()

            if is_grouped_month:
                if is_ambos_visible:
                    self.rb_ambos.pack_forget()
                    # Set default to "Desonerado" if "Ambos" was selected
                    if self.selected_type.get() == 0:
                        self.selected_type.set(1)
            else:  # Not a grouped month
                if not is_ambos_visible:
                    # Re-pack in the correct order to show "Ambos" at the top
                    self.rb_desonerado.pack_forget()
                    self.rb_nao_desonerado.pack_forget()
                    self.rb_ambos.pack(anchor='w')
                    self.rb_desonerado.pack(anchor='w')
                    self.rb_nao_desonerado.pack(anchor='w')

        # Limpa os checkboxes de meses específicos
        if self.months_1_to_4_frame:
            for widget in self.months_1_to_4_frame.winfo_children():
                widget.destroy()

        # A lógica de checkboxes só se aplica a SINAPI
        if service != "SINAPI":
            return

        if year == "2021" and month == "1 a 4":
            Checkbutton(self.months_1_to_4_frame, text="Mês 1", variable=self.jan_2021, command=lambda: sinapi.definir_valor_janeiro_2021(self.jan_2021.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 2", variable=self.feb_2021, command=lambda: sinapi.definir_valor_fevereiro_2021(self.feb_2021.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 3", variable=self.mar_2021, command=lambda: sinapi.definir_valor_marco_2021(self.mar_2021.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 4", variable=self.apr_2021, command=lambda: sinapi.definir_valor_abril_2021(self.apr_2021.get())).pack(side='left')
        
        elif year == "2020" and month == "9 a 12":
            Checkbutton(self.months_1_to_4_frame, text="Mês 9", variable=self.sep_2020, command=lambda: sinapi.definir_valor_setembro_2020(self.sep_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 10", variable=self.oct_2020, command=lambda: sinapi.definir_valor_outubro_2020(self.oct_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 11", variable=self.nov_2020, command=lambda: sinapi.definir_valor_novembro_2020(self.nov_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 12", variable=self.dec_2020, command=lambda: sinapi.definir_valor_dezembro_2020(self.dec_2020.get())).pack(side='left')

        elif year == "2020" and month == "5 a 8":
            Checkbutton(self.months_1_to_4_frame, text="Mês 5", variable=self.may_2020, command=lambda: sinapi.definir_valor_maio_2020(self.may_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 6", variable=self.jun_2020, command=lambda: sinapi.definir_valor_junho_2020(self.jun_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 7", variable=self.jul_2020, command=lambda: sinapi.definir_valor_julho_2020(self.jul_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 8", variable=self.aug_2020, command=lambda: sinapi.definir_valor_agosto_2020(self.aug_2020.get())).pack(side='left')

        elif year == "2020" and month == "1 a 4":
            Checkbutton(self.months_1_to_4_frame, text="Mês 1", variable=self.jan_2020, command=lambda: sinapi.definir_valor_janeiro_2020(self.jan_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 2", variable=self.feb_2020, command=lambda: sinapi.definir_valor_fevereiro_2020(self.feb_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 3", variable=self.mar_2020, command=lambda: sinapi.definir_valor_marco_2020(self.mar_2020.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 4", variable=self.apr_2020, command=lambda: sinapi.definir_valor_abril_2020(self.apr_2020.get())).pack(side='left')

        elif year == "2018" and month == "7 e 8":
            Checkbutton(self.months_1_to_4_frame, text="Mês 7", variable=self.jul_2018, command=lambda: sinapi.definir_valor_julho_2018(self.jul_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 8", variable=self.aug_2018, command=lambda: sinapi.definir_valor_agosto_2018(self.aug_2018.get())).pack(side='left')

        elif year == "2019" and month == "1 a 4":
            Checkbutton(self.months_1_to_4_frame, text="Mês 1", variable=self.jan_2019, command=lambda: sinapi.definir_valor_janeiro_2019(self.jan_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 2", variable=self.feb_2019, command=lambda: sinapi.definir_valor_fevereiro_2019(self.feb_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 3", variable=self.mar_2019, command=lambda: sinapi.definir_valor_marco_2019(self.mar_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 4", variable=self.apr_2019, command=lambda: sinapi.definir_valor_abril_2019(self.apr_2019.get())).pack(side='left')
        
        elif year == "2019" and month == "5 a 8":
            Checkbutton(self.months_1_to_4_frame, text="Mês 5", variable=self.may_2019, command=lambda: sinapi.definir_valor_maio_2019(self.may_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 6", variable=self.jun_2019, command=lambda: sinapi.definir_valor_junho_2019(self.jun_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 7", variable=self.jul_2019, command=lambda: sinapi.definir_valor_julho_2019(self.jul_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 8", variable=self.aug_2019, command=lambda: sinapi.definir_valor_agosto_2019(self.aug_2019.get())).pack(side='left')

        elif year == "2019" and month == "9 a 12":
            Checkbutton(self.months_1_to_4_frame, text="Mês 9", variable=self.sep_2019, command=lambda: sinapi.definir_valor_setembro_2019(self.sep_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 10", variable=self.oct_2019, command=lambda: sinapi.definir_valor_outubro_2019(self.oct_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 11", variable=self.nov_2019, command=lambda: sinapi.definir_valor_novembro_2019(self.nov_2019.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 12", variable=self.dec_2019, command=lambda: sinapi.definir_valor_dezembro_2019(self.dec_2019.get())).pack(side='left')

        elif year == "2018" and month == "1 a 6":
            Checkbutton(self.months_1_to_4_frame, text="Mês 1", variable=self.jan_2018, command=lambda: sinapi.definir_valor_janeiro_2018(self.jan_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 2", variable=self.feb_2018, command=lambda: sinapi.definir_valor_fevereiro_2018(self.feb_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 3", variable=self.mar_2018, command=lambda: sinapi.definir_valor_marco_2018(self.mar_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 4", variable=self.apr_2018, command=lambda: sinapi.definir_valor_abril_2018(self.apr_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 5", variable=self.may_2018, command=lambda: sinapi.definir_valor_maio_2018(self.may_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 6", variable=self.jun_2018, command=lambda: sinapi.definir_valor_junho_2018(self.jun_2018.get())).pack(side='left')

        elif year == "2018" and month == "9 a 12":
            Checkbutton(self.months_1_to_4_frame, text="Mês 9", variable=self.sep_2018, command=lambda: sinapi.definir_valor_setembro_2018(self.sep_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 10", variable=self.oct_2018, command=lambda: sinapi.definir_valor_outubro_2018(self.oct_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 11", variable=self.nov_2018, command=lambda: sinapi.definir_valor_novembro_2018(self.nov_2018.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 12", variable=self.dec_2018, command=lambda: sinapi.definir_valor_dezembro_2018(self.dec_2018.get())).pack(side='left')

        elif year == "2017" and month == "1 a 6":
            Checkbutton(self.months_1_to_4_frame, text="Mês 1", variable=self.jan_2017, command=lambda: sinapi.definir_valor_janeiro_2017(self.jan_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 2", variable=self.feb_2017, command=lambda: sinapi.definir_valor_fevereiro_2017(self.feb_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 3", variable=self.mar_2017, command=lambda: sinapi.definir_valor_marco_2017(self.mar_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 4", variable=self.apr_2017, command=lambda: sinapi.definir_valor_abril_2017(self.apr_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 5", variable=self.may_2017, command=lambda: sinapi.definir_valor_maio_2017(self.may_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 6", variable=self.jun_2017, command=lambda: sinapi.definir_valor_junho_2017(self.jun_2017.get())).pack(side='left')

        elif year == "2017" and month == "7 a 12":
            Checkbutton(self.months_1_to_4_frame, text="Mês 7", variable=self.jul_2017, command=lambda: sinapi.definir_valor_julho_2017(self.jul_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 8", variable=self.aug_2017, command=lambda: sinapi.definir_valor_agosto_2017(self.aug_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 9", variable=self.sep_2017, command=lambda: sinapi.definir_valor_setembro_2017(self.sep_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 10", variable=self.oct_2017, command=lambda: sinapi.definir_valor_outubro_2017(self.oct_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 11", variable=self.nov_2017, command=lambda: sinapi.definir_valor_novembro_2017(self.nov_2017.get())).pack(side='left')
            Checkbutton(self.months_1_to_4_frame, text="Mês 12", variable=self.dec_2017, command=lambda: sinapi.definir_valor_dezembro_2017(self.dec_2017.get())).pack(side='left')


    def create_widgets(self):
        main_frame = Frame(self.master)
        main_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # --- Top section for downloading ---
        download_frame = Frame(main_frame, relief='groove', bd=2)
        download_frame.pack(fill='x', padx=5, pady=5)

        # --- Conteúdo Principal ---
        content_frame = Frame(download_frame)
        content_frame.pack(side='top', fill='x', padx=5, pady=5)

        Label(content_frame, text="Selecione a base de dados:").pack()
        # Combobox for service selection
        self.service_menu = Combobox(content_frame, textvariable=self.selected_service, values=["SINAPI", "SICRO", "ORSE"], state='readonly')
        self.service_menu.pack()
        try:
            self.service_menu.set(self.selected_service.get())
        except Exception:
            pass

        Label(content_frame, text="Selecione a data:").pack()
        date_frame = Frame(content_frame)
        date_frame.pack(pady=2)

        years = [str(y) for y in range(2017, 2025)]
        # Combobox for year selection
        self.year_menu = Combobox(date_frame, textvariable=self.selected_year, values=years, state='readonly')
        self.year_menu.pack(side='left', padx=(0, 6))
        try:
            self.year_menu.set(self.selected_year.get())
        except Exception:
            pass
        
        # Use ttk Combobox for month selection — more consistent in frozen executables
        self.month_menu = Combobox(date_frame, textvariable=self.selected_month, values=[], state='readonly')
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
        self.rb_ambos = Radiobutton(type_frame, text="Ambos", variable=self.selected_type, value=0)
        self.rb_ambos.pack(anchor='w')
        self.rb_desonerado = Radiobutton(type_frame, text="Desonerado", variable=self.selected_type, value=1)
        self.rb_desonerado.pack(anchor='w')
        self.rb_nao_desonerado = Radiobutton(type_frame, text="Não Desonerado", variable=self.selected_type, value=2)
        self.rb_nao_desonerado.pack(anchor='w')

        file_type_frame = Frame(radio_frame)
        file_type_frame.pack(side='left', padx=10)

        Label(file_type_frame, text="Quais arquivos juntar?").pack(anchor='w')
        Radiobutton(file_type_frame, text="Ambos", variable=self.selected_file_type, value="Ambos").pack(anchor='w')
        Radiobutton(file_type_frame, text="Insumos", variable=self.selected_file_type, value="Insumos").pack(anchor='w')
        Radiobutton(file_type_frame, text="Sintéticos", variable=self.selected_file_type, value="Sintetico").pack(anchor='w')

        states_header_frame = Frame(self.sinapi_widgets)
        states_header_frame.pack(pady=2)
        Label(states_header_frame, text="Selecione os estados:").pack(side='left')
        Checkbutton(states_header_frame, text="Selecionar Todos", variable=self.select_all_states_var, command=self.toggle_all_states).pack(side='left', padx=5)

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
        link_label = Label(self.sinapi_widgets, text="Site de donwloads SINAPI", fg="blue", cursor="hand2")
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
        
        orse_links_frame = Frame(self.orse_widgets)
        orse_links_frame.pack()

        link_label = Label(orse_links_frame, text="Site de downloads ORSE", fg="blue", cursor="hand2")
        link_label.pack(side='left', padx=5)
        link_label.bind("<Button-1>", self.open_link_orse)
        link_label.config(font=link_font) 

        drive_label = Label(orse_links_frame, text="Google Drive", fg="blue", cursor="hand2")
        drive_label.pack(side='left', padx=5)
        drive_label.bind("<Button-1>", self.open_link_orse_drive)
        drive_label.config(font=link_font)
        
        # --- Widgets SICRO ---
        self.sicro_widgets = Frame(content_frame)

        sicro_checkbox_frame = Frame(self.sicro_widgets)
        sicro_checkbox_frame.pack(pady=5)
        
        Label(sicro_checkbox_frame, text="Quais planilhas juntar?").pack(anchor='w')
        Checkbutton(sicro_checkbox_frame, text="Composições de Custos", variable=self.sicro_composicoes).pack(anchor='w')
        Checkbutton(sicro_checkbox_frame, text="Equipamentos (com desoneração)", variable=self.sicro_equipamentos_desonerado).pack(anchor='w')
        Checkbutton(sicro_checkbox_frame, text="Equipamentos", variable=self.sicro_equipamentos).pack(anchor='w')
        Checkbutton(sicro_checkbox_frame, text="Materiais", variable=self.sicro_materiais).pack(anchor='w')

        sicro_states_header_frame = Frame(self.sicro_widgets)
        sicro_states_header_frame.pack(pady=2)
        Label(sicro_states_header_frame, text="Selecione os estados:").pack(side='left')
        Checkbutton(sicro_states_header_frame, text="Selecionar Todos", variable=self.select_all_sicro_states_var, command=self.toggle_all_sicro_states).pack(side='left', padx=5)
        
        
        
        
        num_rows_sicro = 3
        total_sicro = len(estados)
        chunk_sicro = (total_sicro + num_rows_sicro - 1) // num_rows_sicro
        sicro_states_rows = [Frame(self.sicro_widgets) for _ in range(num_rows_sicro)]
        for r in sicro_states_rows:
            r.pack(fill='x', padx=8, pady=2)

        for i, estado in enumerate(estados):
            row_idx = min(i // chunk_sicro, num_rows_sicro - 1)
            cb = Checkbutton(sicro_states_rows[row_idx], text=estado, variable=self.selected_sicro_states[estado])
            cb.pack(side='left', anchor='w', padx=4, pady=2)

        # texto de link clicável:
        link_font = font.Font(size=10, underline=True)
        link_label = Label(self.sicro_widgets, text="Site de downloads SICRO", fg="blue", cursor="hand2")
        link_label.pack()
        link_label.bind("<Button-1>", self.open_link_sicro)
        link_label.config(font=link_font)
        
        # --- Botões de download ---
        bottom_frame = Frame(download_frame)
        bottom_frame.pack(side='bottom', fill='x', pady=10)

        self.baixar_button = Button(bottom_frame, text="Baixar", command=self.execute_sinapi)
        self.baixar_button.pack(side='left', padx=5)
        
        self.aninhar_button = Button(bottom_frame, text="Juntar", command=self.execute_aninhar)
        self.aninhar_button.pack(side='left', padx=5)

        self.formatar_button = Button(bottom_frame, text="Formatar", command=self.execute_formatar_aninhados)
        self.formatar_button.pack(side='left', padx=5)
        
        self.saida_btn = Menubutton(bottom_frame, text="Saída...", direction='above', relief='solid', bd=1)
        self.saida_menu = Menu(self.saida_btn, tearoff=0)
        self.saida_menu.add_command(label="Juntar", command=self.select_aninhar_output)
        self.saida_menu.add_command(label="Formatar", command=self.select_formatar_output)
        self.saida_menu.add_command(label="Comparação", command=self.select_comparison_output)
        self.saida_btn.config(menu=self.saida_menu)
        self.saida_btn.pack(side='left', padx=5)
        
        # self.add_state_button = Button(bottom_frame, text="+1", command=self.add_state)
        # self.add_state_button.pack(side='left', padx=5)
        
        self.apagar_button = Button(bottom_frame, text="apagar dados", command=apagar_dados_sync)
        self.apagar_button.pack(side='right', padx=5)

        # --- New Comparison Section ---
        comparison_frame = Frame(main_frame, relief='groove', bd=2)
        comparison_frame.pack(fill='x', padx=5, pady=5, side='bottom')

        Label(comparison_frame, text="Comparar Planilhas", font=font.Font(weight='bold')).pack(pady=5)

        # Project file selection
        proj_frame = Frame(comparison_frame)
        proj_frame.pack(fill='x', padx=10, pady=2)
        Button(proj_frame, text="Arquivo do projeto", command=self.select_project_file).pack(side='left')
        
        Label(proj_frame, text="Coluna do código do INSUMO/SERVIÇO:").pack(side='left', padx=(10, 0))
        Entry(proj_frame, textvariable=self.project_code_col, width=5, validate='key', validatecommand=self.vcmd).pack(side='left')
        Label(proj_frame, text="Coluna do valor do INSUMO/SERVIÇO:").pack(side='left', padx=(10, 0))
        Entry(proj_frame, textvariable=self.project_value_col, width=5, validate='key', validatecommand=self.vcmd).pack(side='left')


        # Database file selection
        db_frame = Frame(comparison_frame)
        db_frame.pack(fill='x', padx=10, pady=2)
        Button(db_frame, text="Arquivo das bases de dados", command=self.select_database_file).pack(side='left')
        
        Label(db_frame, text="Coluna do código do INSUMO/SERVIÇO:").pack(side='left', padx=(10, 0))
        Entry(db_frame, textvariable=self.database_code_col, width=5, validate='key', validatecommand=self.vcmd).pack(side='left')
        Label(db_frame, text="Coluna do valor do INSUMO/SERVIÇO:").pack(side='left', padx=(10, 0))
        Entry(db_frame, textvariable=self.database_value_col, width=5, validate='key', validatecommand=self.vcmd).pack(side='left')
        
        # Start comparison button
        Button(comparison_frame, text="Iniciar Comparação", command=self.start_comparison, width=25, height=3).pack(pady=10)

    def select_aninhar_output(self):
        path = filedialog.askdirectory(title="Selecione pasta de saída para Juntar")
        if path:
            self.custom_aninhar_path.set(path)

    def select_formatar_output(self):
        path = filedialog.askdirectory(title="Selecione pasta de saída para Formatar")
        if path:
            self.custom_formatar_path.set(path)

    def select_comparison_output(self):
        path = filedialog.askdirectory(title="Selecione pasta de saída para Comparação")
        if path:
            self.custom_comparison_path.set(path)

    def execute_formatar_aninhados(self):
        self.formatar_button.config(state="disabled")
        threading.Thread(target=self._run_formatar_and_reenable, daemon=True).start()

    def _run_formatar_and_reenable(self):
        try:
            print("Iniciando formatação de arquivos aninhados...")
            target_dir = self.custom_formatar_path.get() or None
            formatar_aninhados.format_excel_files(target_directory=target_dir)
            # messagebox.showinfo("Sucesso", "Formatação concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro na Formatação", f"Ocorreu um erro: {e}")
        finally:
            self.master.after(0, lambda: self.formatar_button.config(state="normal"))


    def open_link(self, event):
        webbrowser.open_new_tab("https://www.caixa.gov.br/site/paginas/downloads.aspx")

    def open_link_sicro(self, event):
        webbrowser.open_new_tab("https://www.gov.br/dnit/pt-br/assuntos/planejamento-e-pesquisa/custos-referenciais/sistemas-de-custos/sicro/relatorios/relatorios-sicro")
        
    def open_link_orse(self, event):
        webbrowser.open_new_tab("https://orse.cehop.se.gov.br/default.asp")
        
    def open_link_orse_drive(self, event):
        webbrowser.open_new_tab("https://drive.google.com/drive/folders/1ZqlnNuiCGrnKmj2jtEm1UltncGWgOplc?hl=pt-br")
        
        
        
        
        
    def select_project_file(self):
        filepath = filedialog.askopenfilename(
            title="Selecione o arquivo do projeto",
            filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*"))
        )
        if not filepath:
            return

        self.project_full_path = filepath
        
        # valida presença da aba 'Curva ABC' via controller
        try:
            if validate_project_has_curva(filepath):
                messagebox.showinfo("Sucesso", "Arquivo do projeto selecionado com sucesso!")
            else:
                messagebox.showwarning("Aviso", 'A planilha "Curva ABC" não foi encontrada no arquivo do projeto.')
                self.project_full_path = None
        except Exception as e:
            messagebox.showerror("Erro ao ler arquivo", f"Não foi possível ler o arquivo do projeto: {e}")
            self.project_full_path = None

    def select_database_file(self):
        filepath = filedialog.askopenfilename(
            title="Selecione o arquivo das bases de dados",
            filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*"))
        )
        if filepath:
            self.database_full_path = filepath
            messagebox.showinfo("Sucesso", "Arquivo das bases de dados selecionado com sucesso!")
        else:
            print("Nenhum arquivo das bases de dados selecionado.")

    def start_comparison(self):
        print("INICIANDO COMPARAÇÃO!!!")
        if not self.project_full_path or not self.database_full_path:
            messagebox.showwarning("Aviso", "Por favor, selecione o arquivo do projeto e o arquivo das bases de dados.")
            return

        threading.Thread(target=self._run_comparison_thread, daemon=True).start()

    def _run_comparison_thread(self):
        try:
            output_dir = self.custom_comparison_path.get() or None
            
            project_code_col = self.project_code_col.get().upper()
            project_value_col = self.project_value_col.get().upper()
            database_code_col = self.database_code_col.get().upper()
            database_value_col = self.database_value_col.get().upper()

            output = compare_workbooks(
                self.project_full_path, 
                self.database_full_path, 
                output_dir=output_dir,
                project_code_col=project_code_col,
                project_value_col=project_value_col,
                db_code_col=database_code_col,
                db_value_col=database_value_col
            )
            messagebox.showinfo("Sucesso", f"Comparação concluída! Resultados salvos em:\n{output}")
        except Exception as e:
            messagebox.showerror("Erro na Comparação", f"Ocorreu um erro durante a comparação: {e}")




    def execute_aninhar(self):
        tipo_arquivo = self.selected_file_type.get()
        sicro_composicoes = self.sicro_composicoes.get()
        sicro_equipamentos_desonerado = self.sicro_equipamentos_desonerado.get()
        sicro_equipamentos = self.sicro_equipamentos.get()
        sicro_materiais = self.sicro_materiais.get()
        
        base_dir = self.custom_aninhar_path.get() or None
        
        # Executa diretamente em thread para suportar o argumento base_dir
        threading.Thread(target=aninhar.aninhar_arquivos, 
                         args=(base_dir, tipo_arquivo, sicro_composicoes, sicro_equipamentos_desonerado, sicro_equipamentos, sicro_materiais), daemon=True).start()

    def execute_sicro(self):
        year = self.selected_year.get()
        month = self.selected_month.get()
        
        target_states = [s for s, v in self.selected_sicro_states.items() if v.get() == 1]

        if not target_states:
            messagebox.showwarning("Aviso", "Escolha pelo menos um estado.")
            return

        if not all([year, month]):
            messagebox.showwarning("Aviso", "Por favor, selecione ano e mês.")
            return

        links_to_open = []
        for state in target_states:
            link = get_sicro_link(self.sicro_links_data, state, year, month)
            if link:
                links_to_open.append(link)
            else:
                print(f"AVISO: Link de download não encontrado para SICRO {state}/{year}/{month}.")
        
        if links_to_open:
            threading.Thread(target=abrir_links_sync, args=(links_to_open,), daemon=True).start()
        else:
            messagebox.showerror("Erro", "Nenhum link de download encontrado para os parâmetros selecionados.")

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
            funcao_orse_sync(ano, mes, tipo)
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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

        elif ano == 2018 and mes_str == "7 e 8":
            selected_months_2018 = []
            if self.jul_2018.get(): selected_months_2018.append(7)
            if self.aug_2018.get(): selected_months_2018.append(8)

            if not selected_months_2018:
                messagebox.showwarning("Aviso", "Selecione pelo menos um mês de 7 e 8.")
                return

            first_selected_month = selected_months_2018[0]
            for st in target_states:
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

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
                links = gerar_links_sync(ano, first_selected_month, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

        else:
            try:
                mes = int(mes_str)
            except ValueError:
                messagebox.showerror("Erro", "Mês inválido.")
                return
            
            for st in target_states:
                links = gerar_links_sync(ano, mes, tipo, estados_list=[st])
                if links:
                    threading.Thread(target=abrir_links_sync, args=(links,), daemon=True).start()

    def add_state(self):
        # abre outra instância da mesma janela em Toplevel
        new_win = Toplevel(self.master)
        SinapiApp(new_win)

    def toggle_all_states(self):
        target_value = 1 if self.select_all_states_var.get() else 0
        for var in self.selected_states.values():
            var.set(target_value)

    def toggle_all_sicro_states(self):
        target_value = 1 if self.select_all_sicro_states_var.get() else 0
        for var in self.selected_sicro_states.values():
            var.set(target_value)

def run_app():
    root = Tk()
    app = SinapiApp(root)
    root.mainloop()

if __name__ == "__main__":
    run_app()