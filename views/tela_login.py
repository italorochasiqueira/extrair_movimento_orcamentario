import tkinter as tk
from tkinter import messagebox
from controls.caminho_relativo import caminho_relativo


def abrir_tela_login(callback):
    janela = tk.Toplevel()
    janela.title("Tela Acesso")
    janela.geometry("400x250")
    janela.iconbitmap(caminho_relativo("images/icon_postalis.ico"))
    janela.resizable(False, False)

    #Frame Central
    frame = tk.Frame(janela, padx=20, pady=20)
    frame.place(relx=0.5, rely=0.5, anchor="center")

    #Título
    lbl_titulo = tk.Label(
        frame,
        text='Acesso',
        font=("Arial", 16, "bold"), #características da fonte
        bg="#3399FF",       # azul claro
        fg="white",         # fonte branca
        width=20,           # largura semelhante aos campos de entrada
        anchor="center",    # centraliza o texto
        pady=5              # padding vertical interno
    )
    lbl_titulo.grid(row=0, column=0, columnspan=4, pady=5)

    #Usuário
    lbl_usuario = tk.Label(frame, text="Usuário:")
    lbl_usuario.grid(row=1, column=0, sticky="e", pady=5)
    entrada_usuario = tk.Entry(frame, width=30)
    entrada_usuario.grid(row=1, column=1, columnspan=2, pady=5)

    # Senha
    lbl_senha = tk.Label(frame, text="Senha:")
    lbl_senha.grid(row=2, column=0, sticky="e", pady=5)
    entrada_senha = tk.Entry(frame, width=30, show="*")
    entrada_senha.grid(row=2, column=1, columnspan=2, pady=5)

    # Botão de login
    def realizar_login():
        usuario = entrada_usuario.get()
        senha = entrada_senha.get()
        if usuario and senha:
            janela.destroy()
            callback(usuario, senha)
        else:
            messagebox.showerror("Erro", "Preencha usuário e senha.")

    btn_login = tk.Button(frame, text="Entrar", width=12, command=realizar_login)
    btn_login.grid(row=3, column=1, pady=20)

    btn_sair = tk.Button(frame, text="Sair", width=12, command=janela.destroy)
    btn_sair.grid(row=3, column=2, pady=20)

    janela.grab_set()


