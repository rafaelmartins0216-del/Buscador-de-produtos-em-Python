from openpyxl import load_workbook
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

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
        text="Função de interagir com Excel\nestá em desenvolvimento....",
        font=fonte_geral,
        pady=20
    )
    lbl_info.pack()

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
    

#soma os valores do excel
def somar_valores_excel():
    """Soma todos os valores numéricos na primeira planilha do arquivo Excel."""
    
    # Verifica se o arquivo Excel existe antes de prosseguir
    if not verificar_arquivo_excel():
        return

    arquivo = "comparativo_precos.xlsx"

    try:
        # Carrega o arquivo Excel
        workbook = load_workbook(arquivo)
        # Pega a primeira planilha (fazer depois opção de escolher planilha opção valida somente para primeira planilha)

        sheet = workbook.active  

        
        soma_total = 0

        # Itera sobre todas as células na planilha e somente valores numéricos são somados
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                # Verifica se a célula é um número (int ou float)
                if isinstance(cell, (int, float)):
                    soma_total += cell
                    

        #Variavel css
        alinhar_centro = Alignment(horizontal="center", vertical="center")


        soma_total = tratar_preco(soma_total)

        #Redimensionando colunas para melhor visualização
        sheet.column_dimensions['E'].width = 25

        #Adicionando Dados na planilha
        sheet["E3"] = "Soma Total do Valores:"
        sheet["E4"] = soma_total


        #adicionando efeito na celula do excel
        sheet["E3"].alignment = alinhar_centro
        sheet["E4"].alignment = alinhar_centro


        #Feito Cedula
        workbook.close()
        messagebox.showinfo("Resultado", f"A soma total dos preços é: {soma_total}")

        workbook.save("comparativo_precos.xlsx")

    #exceção genérica
    except:
        messagebox.showerror("Erro", "Ocorreu um erro ao processar o arquivo Excel.")




#Retorna o menor valor do excel
def menor_valor_excel():
    """Retorna o menor valor numérico na primeira planilha do arquivo Excel."""
    
    # Verifica se o arquivo Excel existe antes de prosseguir
    if not verificar_arquivo_excel():
        return

    arquivo = "comparativo_precos.xlsx"

    try:
        workbook = load_workbook(arquivo)
        # Pega a primeira planilha (fazer depois opção de escolher planilha)
        sheet = workbook.active  

        menor_valor = None

        # Itera sobre todas as células na planilha para encontrar o menor valor numérico
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if isinstance(cell, (int, float)):
                    if menor_valor is None or cell < menor_valor:
                        menor_valor = cell

        #converte o valor para float formatado para o exel
        menor_valor = tratar_preco(menor_valor)

        #Variavel css
        alinhar_centro = Alignment(horizontal="center", vertical="center")

        #Redimensionando colunas para melhor visualização
        sheet.column_dimensions['F'].width = 25

        #Fecha a cedula
        workbook.close()

        if menor_valor is not None:
            messagebox.showinfo("Resultado", f"O menor Preço é: {menor_valor}")
        else:
            messagebox.showinfo("Resultado", "Nenhum valor numérico encontrado na planilha.")

    #exceção genérica
    except:
        messagebox.showerror("Erro", "Ocorreu um erro ao processar o arquivo Excel.")


def tratar_preco(valor):
    """Converte valor monetário BR/US para float corretamente."""
    
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
