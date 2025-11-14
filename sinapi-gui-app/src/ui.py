from tkinter import Tk, StringVar, IntVar, BooleanVar, Label, OptionMenu, Radiobutton, Checkbutton, Button, Toplevel, messagebox, Frame
from datetime import datetime
import threading
import importlib
import sys
import os

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
        definir_valor_agosto_2018
    )
    from aninhar import aninhar_arquivos, apagar_dados_sinapi
except Exception:
    src_dir = os.path.dirname(__file__)  # pasta src
    if src_dir not in sys.path:
        sys.path.insert(0, src_dir)
    sinapi_mod = importlib.import_module("sinapi")
    gerar_links_sinapi = getattr(sinapi_mod, "gerar_links_sinapi")
    abrir_links_no_navegador = getattr(sinapi_mod, "abrir_links_no_navegador")
    definir_valor_janeiro_2021 = getattr(sinapi_mod, "definir_valor_janeiro_2021")
    definir_valor_fevereiro_2021 = getattr(sinapi_mod, "definir_valor_fevereiro_2021")
    definir_valor_marco_2021 = getattr(sinapi_mod, "definir_valor_marco_2021")
    definir_valor_abril_2021 = getattr(sinapi_mod, "definir_valor_abril_2021")
    aninhar_mod = importlib.import_module("aninhar")
    aninhar_arquivos = getattr(aninhar_mod, "aninhar_arquivos")
    apagar_dados_sinapi = getattr(aninhar_mod, "apagar_dados_sinapi")

    definir_valor_setembro_2020 = getattr(sinapi_mod, "definir_valor_setembro_2020")
    definir_valor_outubro_2020 = getattr(sinapi_mod, "definir_valor_outubro_2020")
    definir_valor_novembro_2020 = getattr(sinapi_mod, "definir_valor_novembro_2020")
    definir_valor_dezembro_2020 = getattr(sinapi_mod, "definir_valor_dezembro_2020")
    definir_valor_maio_2020 = getattr(sinapi_mod, "definir_valor_maio_2020")
    definir_valor_junho_2020 = getattr(sinapi_mod, "definir_valor_junho_2020")
    definir_valor_julho_2020 = getattr(sinapi_mod, "definir_valor_julho_2020")
    definir_valor_agosto_2020 = getattr(sinapi_mod, "definir_valor_agosto_2020")
    definir_valor_janeiro_2020 = getattr(sinapi_mod, "definir_valor_janeiro_2020")
    definir_valor_fevereiro_2020 = getattr(sinapi_mod, "definir_valor_fevereiro_2020")
    definir_valor_marco_2020 = getattr(sinapi_mod, "definir_valor_marco_2020")
    definir_valor_abril_2020 = getattr(sinapi_mod, "definir_valor_abril_2020")
    definir_valor_julho_2018 = getattr(sinapi_mod, "definir_valor_julho_2018")
    definir_valor_agosto_2018 = getattr(sinapi_mod, "definir_valor_agosto_2018")


from resources.states import estados

class SinapiApp:
    def __init__(self, master):
        self.master = master
        master.title("SINAPI")

        self.selected_service = StringVar(value="SINAPI")
        current_year = datetime.now().year
        default_year = str(current_year if 2017 <= current_year <= 2024 else 2024)
        self.selected_year = StringVar(value=default_year)
        self.selected_month = StringVar(value=datetime.now().strftime("%m"))
        self.selected_type = IntVar(value=0)
        self.selected_file_type = StringVar(value="Ambos")
        self.selected_states = {estado: IntVar(value=0) for estado in estados}

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

        self.months_1_to_4_frame = None
        self.month_menu = None

        self.create_widgets()

        self.selected_year.trace_add("write", self._on_year_change)
        self.selected_month.trace_add("write", self._on_month_change)
        
        self._on_year_change()

    def _on_year_change(self, *args):
        year = self.selected_year.get()
        
        months_full = [f"{m:02d}" for m in range(1, 13)]
        months_2021 = ["1 a 4"] + [f"{m:02d}" for m in range(5, 13)]
        months_2020 = ["1 a 4", "5 a 8", "9 a 12"] + [f"{m:02d}" for m in range(1, 13) if m not in [1,2,3,4,5,6,7,8,9,10,11,12]]
        months_2018 = ["7 e 8"] + [f"{m:02d}" for m in range(1, 13) if m not in [7,8]]

        if year == "2021":
            new_months = months_2021
        elif year == "2020":
            new_months = months_2020
        elif year == "2018":
            new_months = months_2018
        else:
            new_months = months_full
        
        current_month = self.selected_month.get()

        menu = self.month_menu["menu"]
        menu.delete(0, "end")
        for month in new_months:
            menu.add_command(label=month, command=lambda value=month: self.selected_month.set(value))
            
        if current_month not in new_months:
            self.selected_month.set(new_months[0])
        else:
            # This is to trigger the _on_month_change callback if the month is still valid
            self.selected_month.set(current_month)

    def _on_month_change(self, *args):
        year = self.selected_year.get()
        month = self.selected_month.get()

        # Clear previous checkboxes
        for widget in self.months_1_to_4_frame.winfo_children():
            widget.destroy()

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

    def create_widgets(self):
        Label(self.master, text="Selecione a base de dados:").pack()
        OptionMenu(self.master, self.selected_service, "SINAPI", "SICRO", "ORSE").pack()

        Label(self.master, text="Selecione a data:").pack()
        date_frame = Frame(self.master)
        date_frame.pack(pady=2)

        years = [str(y) for y in range(2017, 2025)]

        year_menu = OptionMenu(date_frame, self.selected_year, *years)
        year_menu.pack(side='left', padx=(0, 6))

        self.month_menu = OptionMenu(date_frame, self.selected_month, "")
        self.month_menu.pack(side='left')

        self.months_1_to_4_frame = Frame(self.master)
        self.months_1_to_4_frame.pack()

        radio_frame = Frame(self.master)
        radio_frame.pack(pady=2)

        type_frame = Frame(radio_frame)
        type_frame.pack(side='left', padx=10)
        
        Label(type_frame, text="Tipo:").pack(anchor='w')
        Radiobutton(type_frame, text="Ambos", variable=self.selected_type, value=0).pack(anchor='w')
        Radiobutton(type_frame, text="Desonerado", variable=self.selected_type, value=1).pack(anchor='w')
        Radiobutton(type_frame, text="Não Desonerado", variable=self.selected_type, value=2).pack(anchor='w')

        file_type_frame = Frame(radio_frame)
        file_type_frame.pack(side='left', padx=10)

        Label(file_type_frame, text="Tipo de Arquivo:").pack(anchor='w')
        Radiobutton(file_type_frame, text="Ambos", variable=self.selected_file_type, value="Ambos").pack(anchor='w')
        Radiobutton(file_type_frame, text="Insumos", variable=self.selected_file_type, value="Insumos").pack(anchor='w')
        Radiobutton(file_type_frame, text="Sintéticos", variable=self.selected_file_type, value="Sintetico").pack(anchor='w')

        Label(self.master, text="Selecione os estados:").pack()

        num_rows = 3
        total = len(estados)
        chunk = (total + num_rows - 1) // num_rows
        rows = [Frame(self.master) for _ in range(num_rows)]
        for r in rows:
            r.pack(fill='x', padx=8, pady=2)

        for i, estado in enumerate(estados):
            row_idx = min(i // chunk, num_rows - 1)
            cb = Checkbutton(rows[row_idx], text=estado, variable=self.selected_states[estado])
            cb.pack(side='left', anchor='w', padx=4, pady=2)

        bottom_frame = Frame(self.master)
        bottom_frame.pack(pady=10, fill='x')

        Button(bottom_frame, text="Baixar", command=self.execute_sinapi).pack(side='left', padx=5)
        Button(bottom_frame, text="Aninhar", command=self.execute_aninhar).pack(side='left', padx=5)
        Button(bottom_frame, text="+1", command=self.add_state).pack(side='left', padx=5)
        Button(bottom_frame, text="apagar dados", command=apagar_dados_sinapi).pack(side='right', padx=5)

    def execute_aninhar(self):
        tipo_arquivo = self.selected_file_type.get()
        # Executa aninhar_arquivos em uma thread para não bloquear a UI
        threading.Thread(target=aninhar_arquivos, args=(None, tipo_arquivo), daemon=True).start()

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