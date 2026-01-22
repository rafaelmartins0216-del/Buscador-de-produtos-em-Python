import tkinter as tk
from interface import AppInterface
from busca_web import buscar_produtos
from exel import salvar_dados, abrir_arquivo_excel
from interagir_exel import interagir_exel

"""
    Função controladora acionada pelo botão de busca dos produtos.
    """
def executar_busca(produto, loja):
    
    """Verifica se há produto"""
    if not produto:
        app.mostrar_mensagem("Erro", "Por favor, digite um produto.")
        return

    try:
        # 1. Bloqueia a interface visualmente botão abrir exel
        app.alternar_estado_botao("disabled") 
        print(f"--- Iniciando busca por '{produto}' na '{loja}' ---")
        app.master.update() 


        """chamamos o arquivo de busca , passando o site e o produto"""
        dados = buscar_produtos(loja, produto)
        
        """verifica os dados e caso de tudo certo salva ele no exel"""
        if dados:
            salvar_dados(dados, loja)
            app.mostrar_mensagem("Sucesso", f"{len(dados)} produtos encontrados e salvos no Excel!")
        else:
            app.mostrar_mensagem("Aviso", "Nenhum produto encontrado ou erro na busca.")
    
    except Exception as e:
        # Captura erros para o programa não fechar
        print(f"Erro crítico: {e}")
        app.mostrar_mensagem("Erro Crítico", f"Ocorreu um erro inesperado: {e}")
    
    finally:
        # 4. Sempre executa isso no final: Libera o botão novamente
        app.alternar_estado_botao("normal")


def abrir_planilha():
    """Função para abrir o excel e verificar possível erro"""
    try:
        print("Tentando abrir o Excel...")
        abrir_arquivo_excel()
    except Exception as e:
        app.mostrar_mensagem("Erro", f"Não foi possível abrir o arquivo: {e}")

# --- Execução Principal ---
if __name__ == "__main__":
    """
    Configuração inicial da janela
    """
    janela = tk.Tk()

    # Passamos a janela e as funções para a interface
    app = AppInterface(janela, executar_busca, abrir_planilha, interagir_exel)
    
    janela.mainloop()