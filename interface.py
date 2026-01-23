import tkinter as tk
from tkinter import ttk, messagebox

"""
Interface gráfica, pega o nome do produto e escolhe o site
"""

class AppInterface:
    """
    Classe principal da Interface Gráfica.
    """

    def __init__(self, master, comando_buscar, comando_abrir, comando_interagir):
        self.master = master
        self.master.title("Buscador de Preços Modular")
        self.master.geometry("600x500")
        self.master.configure(bg="#f0f2f5")
        self.master.resizable(False, False)

        # Configuração visual
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TLabel", background="#ffffff", font=("Segoe UI", 10), foreground="#333")
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), foreground="#2c3e50")


        self.main_frame = tk.Frame(master, bg="#ffffff", bd=1, relief="solid")
        self.main_frame.place(relx=0.5, rely=0.5, anchor="center", width=500, height=380)

        # Título
        lbl_titulo = ttk.Label(self.main_frame, text="Buscador de Ofertas", style="Title.TLabel")
        lbl_titulo.pack(pady=(30, 20))

        # Frame de entrada
        input_frame = tk.Frame(self.main_frame, bg="#ffffff")
        input_frame.pack(fill="x", padx=40)

        ttk.Label(input_frame, text="Produto:").grid(row=0, column=0, sticky="w", pady=5)
        self.entrada_produto = ttk.Entry(input_frame, width=40)
        self.entrada_produto.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        self.entrada_produto.bind("<KeyRelease>", self.verificar_campos)

        ttk.Label(input_frame, text="Loja:").grid(row=2, column=0, sticky="w", pady=5)
        self.combo_loja = ttk.Combobox(
            input_frame,
            values=["Mercado Livre", "Amazon"],
            state="readonly"
        )
        self.combo_loja.grid(row=3, column=0, sticky="ew", pady=(0, 20))
        self.combo_loja.bind("<<ComboboxSelected>>", self.verificar_campos)

        input_frame.columnconfigure(0, weight=1)

        # Frame dos botões
        btn_frame = tk.Frame(self.main_frame, bg="#ffffff")
        btn_frame.pack(fill="x", padx=40, pady=10)

        #botão buscar produtos
        self.btn_buscar = ttk.Button(
            btn_frame,
            text="BUSCAR PRODUTOS",
            command=lambda: comando_buscar(
                self.entrada_produto.get(),
                self.combo_loja.get()
            ),
            state="disabled"
        )
        self.btn_buscar.pack(fill="x", pady=5)

        #botão abrir excel final
        self.btn_abrir = ttk.Button(
            btn_frame,
            text="📂 Abrir Excel Final",
            command=comando_abrir,
            state="disabled"
        )
        self.btn_abrir.pack(fill="x", pady=5)

        #botão interagir com excel
        self.btn_excel = ttk.Button(
            btn_frame,
            text="📊 Interagir com Excel",
            command=comando_interagir,
            state="enabled"
        )
        self.btn_excel.pack(fill="x", pady=5)

        #botão selecionar Loja
        self.lbl_status = tk.Label(
            self.main_frame,
            text="Selecione a loja e o produto...",
            bg="#ffffff",
            fg="#7f8c8d",
            font=("Segoe UI", 8)
        )
        self.lbl_status.pack(side="bottom", pady=10)

    #verifica se os campos estão preenchidos
    def verificar_campos(self, event=None):
        produto = self.entrada_produto.get()
        loja = self.combo_loja.get()

        if produto and loja:
            self.btn_buscar.config(state="normal")
            self.lbl_status.config(text="Pronto para buscar!", fg="#2c3e50")
        else:
            self.btn_buscar.config(state="disabled")
            self.lbl_status.config(
                text="Preencha o produto e selecione a loja...",
                fg="#e74c3c"
            )

    def mostrar_mensagem(self, titulo, mensagem):
        messagebox.showinfo(titulo, mensagem)

    def alternar_estado_botao(self, estado):
        self.btn_buscar.config(state=estado)

        if estado == "disabled":
            self.lbl_status.config(text="Buscando dados, aguarde...", fg="#e67e22")
            self.btn_abrir.config(state="disabled")
            self.btn_excel.config(state="disabled")
        else:
            self.lbl_status.config(text="Busca finalizada com sucesso!", fg="#27ae60")
            self.btn_abrir.config(state="normal")
            self.btn_excel.config(state="normal")
            self.verificar_campos()
