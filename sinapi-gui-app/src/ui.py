from tkinter import Tk, StringVar, IntVar, Label, OptionMenu, Radiobutton, Checkbutton, Button, Toplevel, messagebox, Entry, Frame
from datetime import datetime
import threading
import importlib
import sys
import os

# tenta import direto; se houver conflito entre dois módulos 'sinapi' no sys.path,
# força carregar o módulo a partir da pasta do src (garante consistência ao rodar main.py)
try:
    from sinapi import gerar_links_sinapi, abrir_links_no_navegador
    from aninhar import aninhar_arquivos
except Exception:
    src_dir = os.path.dirname(__file__)  # pasta src
    if src_dir not in sys.path:
        sys.path.insert(0, src_dir)
    sinapi_mod = importlib.import_module("sinapi")
    gerar_links_sinapi = getattr(sinapi_mod, "gerar_links_sinapi")
    abrir_links_no_navegador = getattr(sinapi_mod, "abrir_links_no_navegador")

from resources.states import estados

class SinapiApp:
    def __init__(self, master):
        self.master = master
        master.title("SINAPI")

        self.selected_service = StringVar(value="SINAPI")
        # replace single date entry with separate year/month dropdowns
        current_year = datetime.now().year
        default_year = str(current_year if 2017 <= current_year <= 2024 else 2024)
        self.selected_year = StringVar(value=default_year)
        self.selected_month = StringVar(value=datetime.now().strftime("%m"))
        self.selected_type = IntVar(value=0)  # 0: Ambos, 1: Desonerado, 2: Não Desonerado
        self.selected_states = {estado: IntVar(value=0) for estado in estados}

        self.create_widgets()

    def create_widgets(self):
        Label(self.master, text="Selecione a base de dados:").pack()
        OptionMenu(self.master, self.selected_service, "SINAPI", "SICRO", "ORSE").pack()

        Label(self.master, text="Selecione a data:").pack()
        # new frame to hold two dropdowns side-by-side
        date_frame = Frame(self.master)
        date_frame.pack(pady=2)

        years = [str(y) for y in range(2017, 2025)]  # 2017..2024
        months = [f"{m:02d}" for m in range(1, 13)]  # 01..12

        year_menu = OptionMenu(date_frame, self.selected_year, *years)
        year_menu.pack(side='left', padx=(0, 6))

        month_menu = OptionMenu(date_frame, self.selected_month, *months) if False else OptionMenu(date_frame, self.selected_month, *months)
        month_menu.pack(side='left')

        Label(self.master, text="").pack()
        Radiobutton(self.master, text="Ambos", variable=self.selected_type, value=0).pack()
        Radiobutton(self.master, text="Desonerado", variable=self.selected_type, value=1).pack()
        Radiobutton(self.master, text="Não Desonerado", variable=self.selected_type, value=2).pack()

        Label(self.master, text="Selecione os estados:").pack()

        # arrange state checkbuttons into 3 rows
        num_rows = 3
        total = len(estados)
        chunk = (total + num_rows - 1) // num_rows  # items per row (roughly equal)
        rows = [Frame(self.master) for _ in range(num_rows)]
        for r in rows:
            r.pack(fill='x', padx=8, pady=2)

        for i, estado in enumerate(estados):
            row_idx = min(i // chunk, num_rows - 1)
            cb = Checkbutton(rows[row_idx], text=estado, variable=self.selected_states[estado])
            cb.pack(side='left', anchor='w', padx=4, pady=2)

        # bottom button row: Concluir, Aninhar (to right of Concluir), +1
        bottom_frame = Frame(self.master)
        bottom_frame.pack(pady=10, fill='x')

        Button(bottom_frame, text="Concluir", command=self.execute_sinapi).pack(side='left', padx=5)
        Button(bottom_frame, text="Aninhar", command=aninhar_arquivos).pack(side='left', padx=5)
        Button(bottom_frame, text="+1", command=self.add_state).pack(side='left', padx=5)

    def execute_sinapi(self, estado: str = None):
        """
        Se estado for fornecido, executa apenas para esse estado.
        Se estado for None, executa para todos os estados selecionados na GUI.
        """
        try:
            ano = int(self.selected_year.get())
            mes = int(self.selected_month.get())
        except ValueError:
            messagebox.showerror("Erro", "Seleção de data inválida. Selecione mês e ano.")
            return

        tipo = ["Ambos", "Desonerado", "NaoDesonerado"][self.selected_type.get()]

        if estado:
            target_states = [estado]
        else:
            target_states = [s for s, v in self.selected_states.items() if v.get() == 1]

        if not target_states:
            messagebox.showwarning("Aviso", "Escolha pelo menos um estado.")
            return

        # Para cada estado selecionado, chama gerar_links_sinapi apenas para esse estado
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