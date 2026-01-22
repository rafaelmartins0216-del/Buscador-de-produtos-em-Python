from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep

def iniciar_driver():
    """Inicia o navegador maximizado"""
    driver = webdriver.Chrome()
    driver.maximize_window()
    return driver

def scroll_ate_o_fim(driver):
    """
    Função scrool para ver todos os produtos
    """
    print("Rolando a página para carregar produtos...")
    last_height = driver.execute_script("return document.body.scrollHeight")
    
    while True:
        # Rola para baixo
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        
        # Espera carregar
        sleep(2)
        
        # Calcula nova altura e compara com a antiga
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break # Chegou no fim
        last_height = new_height

def buscar_produtos(loja, produto):
    driver = iniciar_driver()
    resultados = []
    
    try:
        print(f"--- Iniciando busca na {loja} por: {produto} ---")

        #
        # MERCADO LIVRE
        # 
        if loja == "Mercado Livre":
            driver.get("https://www.mercadolivre.com.br")
            sleep(2)
            
            # Busca
            try:
                #pega a barra de busca e digita o produto
                barra = driver.find_element(By.CLASS_NAME, "nav-search-input")
                barra.clear()
                barra.send_keys(produto)
                #Tecla Enter
                barra.send_keys(Keys.RETURN)
            except Exception:
                print("Erro ao buscar no Mercado Livre.")
                return []

            #Tempo para carregar resultados
            sleep(3)

            #função de scroll
            scroll_ate_o_fim(driver)

            # Captura Itens usando o item pesquisado
            items = driver.find_elements(By.CSS_SELECTOR, "li.ui-search-layout__item")
            
            print(f"Encontrei {len(items)} items. Extraindo...")
            
            # percorre os itens encontrados 
            for item in items:
                # delimita cada item inicialmente caso houver algum erro no processo imprimi ele com erro e pula para o próximo
                try:
                    titulo = "Error"
                    preco = "Error"
                    link = "#"

                    # Extraindo titulo do produto
                    try:
                        titulo = item.find_element(By.CLASS_NAME, "poly-component__title").text
                    except:
                        try:
                            titulo = item.find_element(By.CLASS_NAME, "ui-search-item__title").text
                        except:
                            pass

                    # Extraindo o preço do produto
                    try:
                        preco_obj = item.find_elements(By.CLASS_NAME, "poly-price__current")
                        if preco_obj:
                            preco = preco_obj[0].find_element(By.CLASS_NAME, "andes-money-amount__fraction").text
                        else:
                                preco = item.find_element(By.CLASS_NAME, "andes-money-amount__fraction").text
                    except:
                        pass

                    # 3.extraindo o link do produto
                    try:
                        link = item.find_element(By.TAG_NAME, "a").get_attribute("href")
                    except:
                        pass
                    
                    
                    # Adiciona o resultado se o título for válido
                    if titulo != "Error":
                        resultados.append([titulo, preco, link])

                # Validação final para pular itens com erro e ver no exel
                except Exception as e:
                    continue

        #
        # AMAZON
        #
        elif loja == "Amazon":
            driver.get("https://www.amazon.com.br")
            sleep(2)

            #busca todos os produtos
            try:
                #pega a barra de busca e digita o produto
                barra = driver.find_element(By.ID, "twotabsearchtextbox")
                # Limpa a barra antes de digitar
                barra.clear()

                barra.send_keys(produto)
                #Tecla Enter
                barra.send_keys(Keys.RETURN)

            except Exception:
                print("Erro ao buscar na Amazon, Barra de busca")
                return []

            sleep(3)

            # Sua função de scroll passando a tela
            scroll_ate_o_fim(driver) 

            #tempo para carregar os produtos
            sleep(3)

            # Carrega todos os produtos de uma vez
            items = driver.find_elements(By.CSS_SELECTOR, 'div[data-component-type="s-search-result"]')
            print(f"Encontrei {len(items)} items na Amazon. Extraindo dados...")

            #percorre os itens encontrados 
            for item in items:
                try:
                    titulo = "Error"
                    preco = "Error"
                    link = "#"

                    #pega o atributo alt da imagem
                    try:
                        # Encontra a imagem principal do card do produto
                        imagem = item.find_element(By.CSS_SELECTOR, "img.s-image")
                        
                        # Extrai o texto do atributo 'alt' usando get_atribute
                        titulo = imagem.get_attribute("alt")

                        #removendo a frases "Anuncio Patrocinado e -"
                        titulo = titulo.replace("Anúncio", "").replace("patrocinado", "").replace("–", "").strip()

                    except Exception as e:
                        # print(f"Erro ao pegar título da imagem: {e}")
                        titulo = "Erro após pegar título da imagem"
                        

                    #Pega o preço
                    try:
                        preco_completo = item.find_element(By.CSS_SELECTOR, "span.a-offscreen").get_attribute("textContent")
                        # Limpa o R$ e espaços strip()
                        preco = preco_completo.replace("R$", "").replace("U$", "").strip()
                    except:
                        preco = "Erro ao pegar informação de preço"

                    #Pega o Link
                    try:
                        link_elem = item.find_element(By.CSS_SELECTOR, "a.s-no-outline")
                        link = link_elem.get_attribute("href")
                    except:
                        #Coloca o preço como # para o codigo rodar caso houver erro
                        link = "#"

                    #adicionando a resultado
                    resultados.append([titulo, preco, link])

                #Validação final para pular itens com erro e ver no exel
                except Exception:
                    continue

# Tratamento de erros gerais
    except Exception :
        print(f"Erro Crítico na busca de produtos na {loja}.")
    finally:
        #fecha o navegador
        driver.quit()
    # Retorna os resultados encontrados
    return resultados

