from tkinter import messagebox
from datetime import datetime
from src.sinapi import gerar_links_sinapi, abrir_links_no_navegador
from src.resources.states import estados

class SinapiController:
    def __init__(self, ui):
        self.ui = ui

    def on_concluir(self):
        try:
            selected_date = self.ui.date_input.get()
            selected_type = self.ui.type_var.get()
            selected_states = [state for state, var in zip(estados, self.ui.state_vars) if var.get()]

            if not selected_date or not selected_states:
                messagebox.showwarning("Input Error", "Please select a date and at least one state.")
                return
            
            year, month = map(int, selected_date.split('-'))
            links = gerar_links_sinapi(year, month, selected_type)

            if links:
                abrir_links_no_navegador(links)
            else:
                messagebox.showinfo("No Links", "No links generated for the selected criteria.")
        
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid date in YYYY-MM format.")

    def on_add_one(self):
        # Logic for "+1" button can be implemented here
        pass