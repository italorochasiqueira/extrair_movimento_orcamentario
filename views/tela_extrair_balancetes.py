import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from controls.caminho_relativo import caminho_relativo
from views.tela_login import abrir_tela_login
from controls.extrair_balancete import extrair_balancete_planos

def abrir_tela_extrair_balancete():
    janela = tk.Toplevel()
    janela.title("Postalis - GCO/CPF")
    janela.geometry("400x300")
    janela.iconbitmap(caminho_relativo("images/icon_postalis.ico"))
    janela.resizable(False, False)

    # Frame principal
    frame_principal = tk.Frame(janela, padx=20, pady=20)
    frame_principal.grid(row=0, column=0, sticky="nsew")
    janela.grid_rowconfigure(0, weight=1)
    janela.grid_columnconfigure(0, weight=1)

    # ======== Frame Título ========
    frame_titulo = tk.Frame(frame_principal, pady=10)
    frame_titulo.grid(row=0, column=0, sticky="ew")
    lbl_titulo = tk.Label(frame_titulo, text="Extrair Balancetes - Atena", font=("Arial", 14, "bold"))
    lbl_titulo.grid(row=0, column=0, sticky="n", pady=5)

    # ======== Frame Período (Mês/Ano) ========
    frame_data = tk.LabelFrame(frame_principal, text="Período (Mês/Ano)", padx=10, pady=10)
    frame_data.grid(row=1, column=0, pady=10, sticky="ew")

    # Data Inicial
    tk.Label(frame_data, text="Data Inicial:").grid(row=0, column=0, padx=5, pady=5)
    mes_inicio = tk.Entry(frame_data, width=7)
    mes_inicio.grid(row=0, column=1, padx=5)
    ano_inicio = tk.Entry(frame_data, width=7)
    ano_inicio.grid(row=0, column=2, padx=5)

    # Data Final
    tk.Label(frame_data, text="Data Final:").grid(row=1, column=0, padx=5, pady=5)
    mes_fim = tk.Entry(frame_data, width=7)
    mes_fim.grid(row=1, column=1, padx=5)
    ano_fim = tk.Entry(frame_data, width=7)
    ano_fim.grid(row=1, column=2, padx=5)


    # ======== Frame Botões ========
    frame_botoes = tk.Frame(frame_principal, pady=20)
    frame_botoes.grid(row=4, column=0)

    def login_callback(usuario, senha):
        data_ini = f"{mes_inicio.get()}/{ano_inicio.get()}"
        data_fim = f"{mes_fim.get()}/{ano_fim.get()}"

        print(f"[INFO] Data inicial: {data_ini}")
        print(f"[INFO] Data final: {data_fim}")
        
        messagebox.showinfo("INFO", "Esse processo dará início à extração dos balancetes do PGA, Plano BD e PostalPrev.")

        extrair_balancete_planos(usuario, senha, data_ini, data_fim)
    
    btn_extrair = tk.Button(frame_botoes, text="Extrair", width=15, command=lambda: abrir_tela_login(login_callback))
    btn_extrair.grid(row=0, column=0, padx=10)

    btn_sair = tk.Button(frame_botoes, text="Sair", width=15, command=janela.destroy)
    btn_sair.grid(row=0, column=1, padx=10)

    janela.mainloop()



