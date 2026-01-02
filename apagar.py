import tkinter as tk
from tkinter import ttk

def acao_principal():
    print("Ação Principal: Baixar Arquivo Padrão")

def acao_secundaria_1():
    print("Ação Secundária: Baixar PDF")

def acao_secundaria_2():
    print("Ação Secundária: Baixar Excel")

root = tk.Tk()
root.geometry("300x200")

# --- Criação do Split Button ---
# 1. Um Frame para agrupar os botões e parecer um só componente
frame_split = tk.Frame(root)
frame_split.pack(pady=20)

# 2. O Botão Principal (Lado Esquerdo)
btn_main = ttk.Button(frame_split, text="Baixar", command=acao_principal)
btn_main.pack(side="left", padx=(0, 1)) # padx pequeno para separar visualmente

# 3. O Botão da Seta (Lado Direito)
btn_arrow = ttk.Menubutton(frame_split, text="▼", width=2)
btn_arrow.pack(side="left")

# 4. O Menu do Dropdown
menu_split = tk.Menu(btn_arrow, tearoff=0)
menu_split.add_command(label="Baixar como PDF", command=acao_secundaria_1)
menu_split.add_command(label="Baixar como Excel", command=acao_secundaria_2)

# Associa o menu à seta
btn_arrow.config(menu=menu_split)

root.mainloop()
