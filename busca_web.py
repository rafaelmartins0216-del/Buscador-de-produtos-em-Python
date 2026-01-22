from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep

def iniciar_driver():
    """Inicia o navegador maximizado"""
    # Descomente a linha abaixo para rodar em modo 'headless' (sem janela)
    # options.add_argument("--headless") 
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
                barra = driver.find_element(By.CLASS_NAME, "nav-search-input")
                barra.clear()
                barra.send_keys(produto)
                barra.send_keys(Keys.RETURN)
            except Exception as e:
                print(f"Erro na barra de busca ML: {e}")
                return []

            sleep(3)
            scroll_ate_o_fim(driver)
            
            # Captura Itens (Tenta pegar tanto layout antigo lista quanto layout novo grade)
            items = driver.find_elements(By.CSS_SELECTOR, "li.ui-search-layout__item")
            if not items:
                items = driver.find_elements(By.CSS_SELECTOR, "div.poly-card")
            
            print(f"Encontrei {len(items)} items. Extraindo...")

            for item in items:
                try:
                    titulo = "Sem Título"
                    preco = "0"
                    link = "#"

                    # 1. TÍTULO
                    try:
                        titulo = item.find_element(By.CLASS_NAME, "poly-component__title").text
                    except:
                        try:
                            titulo = item.find_element(By.CLASS_NAME, "ui-search-item__title").text
                        except:
                            pass

                    # 2. PREÇO
                    try:
                        
                        preco_obj = item.find_elements(By.CLASS_NAME, "poly-price__current")
                        if preco_obj:
                            preco = preco_obj[0].find_element(By.CLASS_NAME, "andes-money-amount__fraction").text
                        else:
                            
                            preco_obj = item.find_elements(By.CLASS_NAME, "ui-search-price__second-line")
                            if preco_obj:
                                preco = preco_obj[0].find_element(By.CLASS_NAME, "andes-money-amount__fraction").text
                            else:
                                # Fallback genérico
                                preco = item.find_element(By.CLASS_NAME, "andes-money-amount__fraction").text
                    except:
                        pass

                    # 3. LINK
                    try:
                        link = item.find_element(By.TAG_NAME, "a").get_attribute("href")
                    except:
                        pass
                    
                    if titulo != "Sem Título":
                        resultados.append([titulo, preco, link])

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
                barra = driver.find_element(By.ID, "twotabsearchtextbox")
                barra.clear()
                barra.send_keys(produto)
                barra.send_keys(Keys.RETURN)
            except Exception as e:
                print(f"Erro na barra de busca Amazon: {e}")
                return []

            sleep(3)
            scroll_ate_o_fim(driver) # Sua função de scroll

            # --- CAPTURA ---
            # O seletor do print está correto:
            items = driver.find_elements(By.CSS_SELECTOR, 'div[data-component-type="s-search-result"]')
            print(f"Encontrei {len(items)} items na Amazon. Extraindo dados...")

            for item in items:
                try:
                    titulo = "Sem Título"
                    preco = "0"
                    link = "#"

                    #pega o atributo alt da imagem
                    try:
                        # Encontra a imagem principal do card do produto
                        imagem = item.find_element(By.CSS_SELECTOR, "img.s-image")
                        
                        # Extrai o texto do atributo 'alt' usando get_atribute
                        titulo = imagem.get_attribute("alt")

                        #removendo a frase Anuncio Patrocinado e -
                        titulo = titulo.replace("Anúncio", "").replace("patrocinado", "").replace("–", "").strip()

                    except Exception as e:
                        # print(f"Erro ao pegar título da imagem: {e}")
                        titulo = "Sem Título"
                        

                    #Pega o preço
                    try:
                        preco_oculto = item.find_element(By.CSS_SELECTOR, "span.a-offscreen")
                        preco_completo = preco_oculto.get_attribute("textContent")
                        
                        # Limpa o R$ e espaços
                        preco = preco_completo.replace("R$", "").replace("U$", "").strip()
                    except:
                        # Se falhar, tenta o método visual (inteiro)
                        try:
                            preco = item.find_element(By.CLASS_NAME, "a-price-whole").text
                        except:
                            preco = "0"

                    #Pega o Link
                    try:
                        link_elem = item.find_element(By.CSS_SELECTOR, "a.s-no-outline")
                        link = link_elem.get_attribute("href")
                    except:
                        #Coloca o preço como # para o codigo rodar
                        link = "#"

                    #adicionando a resultado
                    resultados.append([titulo, preco, link])

                #Validação para Debug
                except Exception as e:
                    # print(f"Pulei um item: {e}") #descomente para debugar
                    continue

    except Exception as e:
        print(f"Erro Crítico: {e}")
    finally:
        driver.quit()
    
    return resultados

