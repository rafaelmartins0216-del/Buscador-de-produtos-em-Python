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
        #Função master (função principal da tela)
        self.master = master
        self.master.title("Buscador de Preços Modular")
        self.master.geometry("600x450")
        self.master.configure(bg="#f0f2f5") 
        self.master.resizable(False, False) 

        """Configuração visual dos componentes"""
        style = ttk.Style()
        # Tema Limpo
        style.theme_use('clam') 
        style.configure("TLabel", background="#ffffff", font=("Segoe UI", 10), foreground="#333")
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), foreground="#2c3e50")
        
        #Muda de cor (efeito ao clicar)
        style.map("Accent.TButton",
            background=[('active', '#2980b9'), ('!disabled', '#3498db')],
            foreground=[('!disabled', 'white')]
        )

        self.main_frame = tk.Frame(master, bg="#ffffff", bd=1, relief="solid")
        self.main_frame.place(relx=0.5, rely=0.5, anchor="center", width=500, height=350)

        # Título do Card programa principal
        lbl_titulo = ttk.Label(self.main_frame, text="Buscador de Ofertas", style="Title.TLabel")
        lbl_titulo.pack(pady=(30, 20))

        # Campo de entrada de dados css
        input_frame = tk.Frame(self.main_frame, bg="#ffffff")
        input_frame.pack(fill="x", padx=40)

        # Campo: Nome do Produto
        ttk.Label(input_frame, text="Produto:").grid(row=0, column=0, sticky="w", pady=5)
        self.entrada_produto = ttk.Entry(input_frame, width=40, font=("Segoe UI", 10))
        self.entrada_produto.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        
        #Bind para verificar quando o usuário digita ---
        self.entrada_produto.bind("<KeyRelease>", self.verificar_campos)

        # Campo: Seleção de Loja
        ttk.Label(input_frame, text="Loja:").grid(row=2, column=0, sticky="w", pady=5)
        self.combo_loja = ttk.Combobox(input_frame, values=["Mercado Livre", "Amazon"], state="readonly", font=("Segoe UI", 10))
        self.combo_loja.grid(row=3, column=0, sticky="ew", pady=(0, 20))

        # Bind para verificar quando o usuário seleciona a loja ---
        self.combo_loja.bind("<<ComboboxSelected>>", self.verificar_campos)

        # Garante que a coluna se expanda se necessário
        input_frame.columnconfigure(0, weight=1)

        #Botão de ação
        btn_frame = tk.Frame(self.main_frame, bg="#ffffff")
        btn_frame.pack(fill="x", padx=40, pady=10)

        # Botão Buscar
        self.btn_buscar = ttk.Button(
            btn_frame, 
            text="BUSCAR PRODUTOS", 
            style="Accent.TButton",
            command=lambda: comando_buscar(self.entrada_produto.get(), self.combo_loja.get()),
            state="disabled" # --- ALTERADO: Inicia desabilitado ---
        )
        self.btn_buscar.pack(fill="x", pady=5)

        # Botão Abrir Excel (Inicia desabilitado)
        self.btn_abrir = ttk.Button(
            btn_frame, 
            text="📂 Abrir Excel Final", 
            command=comando_abrir,
            state="disabled"  # Só será ativado após uma busca com sucesso
        )
        self.btn_abrir.pack(fill="x", pady=5)

        # botão para interagir com exel
        self.bnt_exel=ttk.Button(
            btn_frame,
            text="interagir com exel",
            command=comando_interagir,
            state="disabled" #so será ativado se tiver um exel criado
        )
        
        # Barra de status
        self.lbl_status = tk.Label(self.main_frame, text="Selecione a loja e o produto...", bg="#ffffff", fg="#7f8c8d", font=("Segoe UI", 8))
        self.lbl_status.pack(side="bottom", pady=10)

    # função habilitar botões
    def verificar_campos(self, event=None):
        """
        Verifica se o produto foi digitado E se a loja foi selecionada.
        Se ambos estiverem ok, habilita o botão de busca.
        """
        produto = self.entrada_produto.get()
        loja = self.combo_loja.get()

        # Se tiver texto no produto E texto na loja
        if produto and loja:
            self.btn_buscar.config(state="normal")
            self.lbl_status.config(text="Pronto para buscar!", fg="#2c3e50")
        else:
            self.btn_buscar.config(state="disabled")
            self.lbl_status.config(text="Preencha o produto e selecione a loja...", fg="#e74c3c")

    """Verifica a mensagem de erro"""
    def mostrar_mensagem(self, titulo, mensagem):
        messagebox.showinfo(titulo, mensagem)

    """Habilita o estado do botão depois de pesquisar"""
    def alternar_estado_botao(self, estado):
        # Nota: Se o estado for 'normal', precisamos garantir que os campos ainda 
        # estejam validos, mas geralmente aqui é controlado pelo processo de busca
        self.btn_buscar.config(state=estado)
        
        if estado == "disabled":
            # Estado durante o processamento (Busca em andamento)
            self.lbl_status.config(text="Buscando dados, aguarde...", fg="#e67e22") 
            self.btn_abrir.config(state="disabled") 
        else:
            # Estado após o término (Pronto)
            self.lbl_status.config(text="Busca finalizada com sucesso!", fg="#27ae60") 
            self.btn_abrir.config(state="normal") 
            # Re-executa verificação para garantir UI consistente
            self.verificar_campos()