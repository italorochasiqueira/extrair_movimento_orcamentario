import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from controls.caminho_relativo import caminho_relativo
from views.tela_login import abrir_tela_login
from controls.extrair_analise_comparativa_atena import extrair_orcamento_atena
from controls.extrair_analise_comparativa_bd import extrair_orcamento_atena_plano_bd
from controls.extrair_analise_comparativa_pp import extrair_orcamento_atena_plano_pp

def abrir_tela_orcamento():
    janela = tk.Toplevel()
    janela.title("Postalis - GCO/CPF")
    janela.geometry("500x500")
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
    lbl_titulo = tk.Label(frame_titulo, text="Extrair Análise Comparativas - Atena", font=("Arial", 14, "bold"))
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

    # ======== Frame Tipo de Plano ========
    frame_plano = tk.LabelFrame(frame_principal, text="Selecionar Plano", padx=10, pady=10)
    frame_plano.grid(row=2, column=0, pady=10, sticky="ew")
    plano_var = tk.StringVar(value="PGA")

    rdb_pga = tk.Radiobutton(frame_plano, text="Plano PGA", variable=plano_var, value="PGA")
    rdb_pga.grid(row=0, column=0, padx=10, pady=5, sticky="w")
    rdb_bd = tk.Radiobutton(frame_plano, text="Plano BD", variable=plano_var, value="BD")
    rdb_bd.grid(row=0, column=1, padx=10, pady=5, sticky="w")
    rdb_postalprev = tk.Radiobutton(frame_plano, text="PostalPrev", variable=plano_var, value="PostalPrev")
    rdb_postalprev.grid(row=0, column=2, padx=10, pady=5, sticky="w")

    # ======== Frame Movimento Contábil/Financeiro ========
    frame_movimento = tk.LabelFrame(frame_principal, text="Selecionar Tipo de Realização", padx=10, pady=10)
    frame_movimento.grid(row=3, column=0, pady=15, sticky="ew")
    movimento_var = tk.StringVar(value="0")

    rdb_contabil = tk.Radiobutton(frame_movimento, text="Contábil", variable=movimento_var, value="0")
    rdb_contabil.grid(row=0, column=0, padx=10, pady=5, sticky="w")
    rdb_financeiro = tk.Radiobutton(frame_movimento, text="Financeiro", variable=movimento_var, value="1")
    rdb_financeiro.grid(row=0, column=1, padx=10, pady=5, sticky="w")


    # ======== Frame Botões ========
    frame_botoes = tk.Frame(frame_principal, pady=20)
    frame_botoes.grid(row=4, column=0)

    def importar_arquivos():
        movimento = movimento_var.get()
        plano = plano_var.get()
        data_ini = f"{mes_inicio.get()}/{ano_inicio.get()}"
        data_fim = f"{mes_fim.get()}/{ano_fim.get()}"

        print(f"[INFO] Data inicial: {data_ini}")
        print(f"[INFO] Data final: {data_fim}")

        # Aqui você coloca a lógica de importação
        try:
            df = None
            if plano == "PGA":
                def login_callback(usuario, senha):
                    extrair_orcamento_atena(usuario, senha, data_ini, data_fim, movimento)
                abrir_tela_login(login_callback)

            elif plano == "BD":
                def login_callback(usuario, senha):
                    extrair_orcamento_atena_plano_bd(usuario, senha, data_ini, data_fim, movimento)
                abrir_tela_login(login_callback)
            elif plano == "PostalPrev":
                def login_callback(usuario, senha):
                    extrair_orcamento_atena_plano_pp(usuario, senha, data_ini, data_fim, movimento)
                abrir_tela_login(login_callback)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro na importação: {e}")

    btn_extrair = tk.Button(frame_botoes, text="Extrair", width=15, command=importar_arquivos)
    btn_extrair.grid(row=0, column=0, padx=10)

    btn_sair = tk.Button(frame_botoes, text="Sair", width=15, command=janela.destroy)
    btn_sair.grid(row=0, column=1, padx=10)

    janela.mainloop()
