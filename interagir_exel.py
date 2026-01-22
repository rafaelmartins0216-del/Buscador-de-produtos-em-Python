from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, messagebox

def mostrar_janela_exel(master=None):
    """Mostra uma janela simples informando que a função está em desenvolvimento."""
    
    janela = tk.Toplevel(master)
    janela.title("Interagir com Excel")
    janela.geometry("600x500")
    janela.configure(bg="#f0f2f5")
    janela.resizable(False, False)

    # Commando pra fixar a janela na frente
    janela.attributes('-topmost', True)


    fonte_geral = ("Segoe UI", 10)

    lbl_info = tk.Label(
        janela,
        text="Função de interagir com Excel\nestá em desenvolvimento.",
        font=fonte_geral,
        pady=20
    )
    lbl_info.pack()

    bnt_verificar = tk.Button(
        janela,
        text="Verificar Existência do Arquivo Excel",
        font=fonte_geral,
        command=verificar_arquivo_excel
    )
    bnt_verificar.pack(pady=30)

    btn_fechar = tk.Button(
        janela,
        text="Fechar",
        command=janela.destroy,
        font=fonte_geral
    )
    btn_fechar.pack(pady=10)


def verificar_arquivo_excel():
    """Verifica se o arquivo Excel existe e pode ser aberto."""
    arquivo = "comparativo_precos.xlsx"
    try:
        # Tenta abrir apenas para ver se existe e é válido
        workbook = load_workbook(filename=arquivo)
        workbook.close()
        return True
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo não encontrado:\n{arquivo}")
        return False
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir o arquivo:\n{e}")
        return False