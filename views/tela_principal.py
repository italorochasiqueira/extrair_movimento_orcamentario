import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from controls.caminho_relativo import caminho_relativo
from views.tela_extrair_orcamento import abrir_tela_orcamento
from views.tela_extrair_balancetes import abrir_tela_extrair_balancete



def abrir_tela_principal():
    janela = tk.Tk()
    janela.title("Postalis - GCO/CPF")
    janela.geometry("600x500")
    janela.iconbitmap(caminho_relativo("images/icon_postalis.ico"))
    janela.resizable(False, False)

    menubar = tk.Menu(janela)
    menu_principal = tk.Menu(menubar, tearoff=0)
    menu_principal.add_command(label="Parâmetros", command=...)
    menu_principal.add_separator()
    menu_principal.add_command(label="Sair", command=janela.destroy)

    menubar.add_cascade(label="Menu", menu=menu_principal)

    menu_previdenciario = tk.Menu(menubar, tearoff=0)
    menu_previdenciario.add_command(label="Extrair Análise Comparativa", command=abrir_tela_orcamento)
    menu_previdenciario.add_command(label="Importar Integração Contábil", command=...)
    
    menu_previdenciario.add_separator()
    menu_relatorio_prev = tk.Menu(menu_previdenciario, tearoff=0)
    menu_previdenciario.add_cascade(label="Relatórios", menu=menu_relatorio_prev)
    menubar.add_cascade(label="Orçamento", menu=menu_previdenciario)

     
    menu_contabil = tk.Menu(menubar, tearoff=0)
    menu_contabil.add_command(label="Extrair Movimento", command=abrir_tela_extrair_balancete)
    menu_contabil.add_separator()

    #sub-menu relatórios
    menu_relatorios = tk.Menu(menu_contabil, tearoff=0)
    menu_relatorios.add_command(label="Rel. Balancete", command=...)
    menu_relatorios.add_command(label="Rel. Razão", command=...)

    menu_contabil.add_cascade(label="Relatórios", menu=menu_relatorios)
    menubar.add_cascade(label="Movimento Contábil", menu=menu_contabil)

    janela.config(menu=menubar)

    # Frame principal
    frame = tk.Frame(janela, padx=20, pady=20)
    frame.place(relx=0.5, rely=0.5, anchor="center")

    # === Carregar a imagem da logo ===
    caminho_logo = caminho_relativo("images/logo_postalis.jpg")
    imagem_original = Image.open(caminho_logo)
    imagem_redimensionada = imagem_original.resize((400, 80))  # ajuste o tamanho conforme necessário
    logo_img = ImageTk.PhotoImage(imagem_redimensionada)

    # === Label da logo ===
    lbl_logo = tk.Label(frame, image=logo_img)
    lbl_logo.image = logo_img  # mantém uma referência para não ser coletado pelo garbage collector
    lbl_logo.grid(row=0, column=0, pady=(0, 10))

    # Título
    lbl_titulo = tk.Label(frame, text="Sistema Coordenação Planejamento Econômico-Financeira", font=("Arial", 16, "bold"))
    lbl_titulo.grid(row=1, column=0, pady=(0, 15))

    janela.mainloop()