from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

def mostrar_janela_exel(master=None):
    """Mostra a janela principal com interface mais agradável."""

    janela = tk.Toplevel(master)
    janela.title("Interagir com Excel")
    janela.geometry("500x450")
    janela.configure(bg="#f5f7fa")
    janela.resizable(False, False)

    # Mantém a janela na frente
    janela.attributes('-topmost', True)

    # Fonte padrão
    fonte_titulo = ("Segoe UI", 14, "bold")
    fonte_botao = ("Segoe UI", 10)

    # Container principal
    frame = tk.Frame(janela, bg="#f5f7fa")
    frame.pack(expand=True, fill="both", padx=20, pady=20)

    # Título
    lbl_titulo = tk.Label(
        frame,
        text="📊 Ferramentas para Excel",
        font=fonte_titulo,
        bg="#f5f7fa",
        fg="#333"
    )
    lbl_titulo.pack(pady=(0, 20))

    # Função para criar botões padronizados
    def criar_botao(texto, comando):
        return tk.Button(
            frame,
            text=texto,
            command=comando,
            font=fonte_botao,
            bg="#4f46e5",
            fg="white",
            activebackground="#4338ca",
            activeforeground="white",
            relief="flat",
            height=2,
            width=35,
            cursor="hand2"
        )

    # Botões
    criar_botao(
        "Somar Todos os Valores do Excel",
        somar_valores_excel
    ).pack(pady=6)

    criar_botao(
        "Retornar o Menor Valor do Excel",
        menor_valor_excel
    ).pack(pady=6)

    criar_botao(
        "Retornar o Maior Valor do Excel",
        maior_valor_exel
    ).pack(pady=6)

    criar_botao(
        "Abrir Arquivo Excel",
        abrir_arquivo_excel
    ).pack(pady=6)

    # Separador
    ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=20)

    # Botão fechar
    btn_fechar = tk.Button(
        frame,
        text="Fechar",
        command=janela.destroy,
        font=fonte_botao,
        bg="#e11d48",
        fg="white",
        activebackground="#be123c",
        relief="flat",
        height=2,
        width=20,
        cursor="hand2"
    )
    btn_fechar.pack()
    """Mostra uma janela simples informando que a função está em desenvolvimento."""
    
    janela = tk.Toplevel(master)
    janela.title("Interagir com Excel")
    janela.geometry("600x500")
    janela.configure(bg="#f0f2f5")
    janela.resizable(False, False)

    # Commando pra fixar a janela na frente
    janela.attributes('-topmost', True)


    fonte_geral = ("Segoe UI", 10)


    # botao para somar valores do excel
    bnt_verificar = tk.Button(
        janela,
        text="Somar Todos os Valores do Excel",
        font=fonte_geral,
        command=somar_valores_excel
    )
    bnt_verificar.pack(pady=10)


    # botao para retornar o menor valor do excel
    bnt_menor = tk.Button(
        janela,
        text="Retornar o Menor Valor do Excel",
        font=fonte_geral,
        command=menor_valor_excel
    )
    bnt_menor.pack(pady=10)

    #botão para retornar maior valor do exel
    bnt_maior=tk.Button(
        janela,
        text="Retornar o Maior Valor",
        font=fonte_geral,
        command=maior_valor_exel
    )

    bnt_maior.pack(pady=10)


    #botão paraabrir o arquivo excel
    btn_abrir_excel = tk.Button(
        janela,
        text="Abrir Arquivo Excel",
        font=fonte_geral,
        command=abrir_arquivo_excel
    )
    btn_abrir_excel.pack(pady=10)


    # Botão para fechar a janela
    btn_fechar = tk.Button(
        janela,
        text="Fechar",
        command=janela.destroy,
        font=fonte_geral
    )

    btn_fechar.pack(side="bottom", pady=20)


#verifica se o arquivo excel existe
def verificar_arquivo_excel():
    """Verifica se o arquivo Excel existe """

    # Nome do arquivo Excel a ser verificado
    arquivo = "comparativo_precos.xlsx"

    try:
        # Tenta abrir apenas para ver se existe e é válido
        workbook = load_workbook(arquivo)
        workbook.close()
        print("Arquivo Encontrado e válido.")
        return True
    
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo não encontrado:\n{arquivo}")
        return False
    
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir o arquivo:\n{e}")
        return False
    


def somar_valores_excel():
    """Soma apenas os valores da coluna de Preço (coluna B).
        returnando o menor valor
    """

    if not verificar_arquivo_excel():
        return

    arquivo = "comparativo_precos.xlsx"

    try:
        workbook = load_workbook(arquivo)
        sheet = workbook.active

        soma_total = 0.0

        # Começa da linha 3 (pulando cabeçalho)
        for row in sheet.iter_rows(min_row=3, min_col=2, max_col=2, values_only=True):
            cell = row[0]

            if isinstance(cell, (int, float)):
                soma_total += cell
            elif isinstance(cell, str):
                soma_total += tratar_preco(cell)

        # Estilo
        alinhar_centro = Alignment(horizontal="center", vertical="center")

        # Ajuste de coluna
        sheet.column_dimensions['E'].width = 25

        # Escreve resultado
        sheet["E3"] = "Soma Total dos Valores:"
        sheet["E4"] = soma_total

        sheet["E3"].alignment = alinhar_centro
        sheet["E4"].alignment = alinhar_centro

        # Formatação monetária no Excel
        sheet["E4"].number_format = 'R$ #,##0.00'

        # Salva antes de fechar
        workbook.save(arquivo)
        workbook.close()

        messagebox.showinfo("Resultado", f"A soma total dos preços é: R$ {soma_total:,.2f}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")



def maior_valor_exel():
    """Escreve o maior valor da coluna de Preço ao lado da Soma Total."""

    if not verificar_arquivo_excel():
        return

    arquivo = "comparativo_precos.xlsx"

    try:
        workbook = load_workbook(arquivo)
        sheet = workbook.active

        cell=[]

        # Começa da linha 3 (pulando cabeçalho)
        for row in sheet.iter_rows(min_row=3, min_col=2, max_col=2, values_only=True):
            cell.append(row[0])
        

        maior_valor=cell[0]
        #percorrendo a lista pra pegar o menor valor
        for i in cell:
            if i>maior_valor:
                maior_valor=i
        
        # Estilo
        alinhar_centro = Alignment(horizontal="center", vertical="center")

        # Ajuste de coluna
        sheet.column_dimensions['I'].width = 25

        # Escreve resultado
        sheet["I3"] = "Maior Valor:"
        sheet["I4"] = maior_valor

        sheet["I3"].alignment = alinhar_centro
        sheet["I4"].alignment = alinhar_centro

        # Formatação monetária no Excel
        sheet["I4"].number_format = 'R$ #,##0.00'

        # Salva antes de fechar
        workbook.save(arquivo)
        workbook.close()

        messagebox.showinfo("Resultado", f"O menor valor da: R$ {maior_valor:,.2f}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")



def menor_valor_excel():
    """Escreve o menor valor da coluna de Preço ao lado da Soma Total."""

    if not verificar_arquivo_excel():
        return

    arquivo = "comparativo_precos.xlsx"

    try:
        workbook = load_workbook(arquivo)
        sheet = workbook.active

        cell=[]

        # Começa da linha 3 (pulando cabeçalho)
        for row in sheet.iter_rows(min_row=3, min_col=2, max_col=2, values_only=True):
            cell.append(row[0])
        

        menor_valor=cell[0]
        #percorrendo a lista pra pegar o menor valor
        for i in cell:
            if i<menor_valor:
                menor_valor=i
        
        # Estilo
        alinhar_centro = Alignment(horizontal="center", vertical="center")

        # Ajuste de coluna
        sheet.column_dimensions['G'].width = 25

        # Escreve resultado
        sheet["G3"] = "Menor Valor:"
        sheet["G4"] = menor_valor

        sheet["G3"].alignment = alinhar_centro
        sheet["G4"].alignment = alinhar_centro

        # Formatação monetária no Excel
        sheet["G4"].number_format = 'R$ #,##0.00'

        # Salva antes de fechar
        workbook.save(arquivo)
        workbook.close()

        messagebox.showinfo("Resultado", f"O menor valor da: R$ {menor_valor:,.2f}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")



def tratar_preco(valor):
    """Converte valor monetário BR para float corretamente."""
    
    if not valor:
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    valor_str = str(valor).strip().lower()
    
    # Remove símbolos
    valor_str = valor_str.replace('r$', '').replace('us$', '').strip()

    if ',' in valor_str:
        valor_str = valor_str.replace('.', '')  # remove milhar
        valor_str = valor_str.replace(',', '.')  # vírgula vira ponto

    elif '.' in valor_str:
        partes = valor_str.split('.')
        
        if len(partes[-1]) == 3:
            valor_str = valor_str.replace('.', '')
    
    try:
        return float(valor_str)
    except ValueError:
        return 0.0


#função abrir arquivo excel
def abrir_arquivo_excel():
    """Abre o arquivo Excel usando o aplicativo padrão do sistema."""
    try:
        os.startfile("comparativo_precos.xlsx")
    except:
        messagebox.showerror("Erro", "Não foi possível abrir o arquivo Excel verifique se ele existe.")
