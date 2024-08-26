from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import re
import time

# Configurando o Selenium WebDriver
driver = webdriver.Chrome()

try:
    # Acessando a página do Magazine Luiza e realiza a busca por "notebooks"
    driver.get("https://www.magazineluiza.com.br/")
    search_box = driver.find_element(By.ID, "input-search")
    search_box.send_keys("notebooks")
    search_box.submit()

    # Aguarda o carregamento da página
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "[data-testid='product-card-container']")))

    # Rola a página para baixo para garantir que todos os produtos sejam carregados
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # Coletando os produtos
    data = []
    seen_urls = set()  # Para rastrear URLs dos produtos já processados
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:
        try:
            products = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[data-testid='product-card-container']"))
            )
            if not products:
                break
            
            new_products_found = False
            
            for product in products:
                try:
                    driver.execute_script("arguments[0].scrollIntoView();", product)  # Rolagem para garantir visibilidade total da página
                    
                    nome = product.find_element(By.CSS_SELECTOR, "[data-testid='product-title']").text
                    
                    # Tenta encontrar o link dentro do produto
                    try:
                        link_element = product.find_element(By.XPATH, ".//ancestor::a")
                        url = link_element.get_attribute("href")
                    except Exception as e:
                        print(f"Erro ao encontrar link: {e}")
                        url = None

                    # Se não conseguir encontrar o URL, pula para o próximo
                    if not url:
                        continue
                    
                    # Verifica se a URL do produto já foi processada
                    if url in seen_urls:
                        continue
                    seen_urls.add(url)
                    
                    # Filtra para garantir que o produto é um notebook
                    if 'notebook' not in nome.lower():
                        continue

                    # Extrai a quantidade total de avaliações
                    try:
                        review_element = product.find_element(By.CSS_SELECTOR, "[data-testid='review']")
                        qtd_aval_text = review_element.find_element(By.CSS_SELECTOR, "span[format='score-count']").text
                        
                        # Extraindo a quantidade total de avaliações usando regex
                        qtd_aval_match = re.search(r'\d+', qtd_aval_text)
                        qtd_aval = int(qtd_aval_match.group()) if qtd_aval_match else 0
                    except Exception as e:
                        print(f"Erro ao extrair avaliações: {e}")
                        qtd_aval = 0
                    
                    # Verifique se os dados estão sendo extraídos corretamente
                    print(f"Produto: {nome}, Avaliações: {qtd_aval_text}, URL: {url}")

                    # Adiciona os dados na lista
                    data.append([nome, qtd_aval, url])
                except Exception as e:
                    print(f"Erro ao processar produto: {e}")
                    continue

            # Rola a página para baixo para carregar mais produtos
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)  # Pausa para o carregamento do conteúdo

            # Verifica se a altura da página mudou para evitar loop infinito
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
            
        except Exception as e:
            print(f"Erro ao coletar produtos: {e}")
            break

finally:
    # Fecha o driver
    driver.quit()

# Filtra os dados
df = pd.DataFrame(data, columns=['PRODUTO', 'QTD_AVAL', 'URL'])
df = df[df['QTD_AVAL'] > 0]  # Remove os produtos sem avaliações

# Verifique o DataFrame antes de salvar
print(df.head())
print(df.info())

# Cria as abas "Piores" e "Melhores"
pior_df = df[df['QTD_AVAL'] < 100]
melhor_df = df[df['QTD_AVAL'] >= 100]

# Cria a pasta Output se não existir
os.makedirs('Output', exist_ok=True)

# Salva o arquivo Excel com as abas
output_path = os.path.join('Output', 'Notebooks.xlsx')
with pd.ExcelWriter(output_path) as writer:
    pior_df.to_excel(writer, sheet_name='Piores', index=False)
    melhor_df.to_excel(writer, sheet_name='Melhores', index=False)

print(f'Arquivo Excel salvo em: {output_path}')

# Configura o envio de e-mail
sender_email = 'gabriel_felype@hotmail.com'  
receiver_email = 'gabriel_felype@hotmail.com'  
app_password = 'jhlafuoentswcnuc'  

subject = 'Relatório Notebooks'
body = """Olá,

Aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.

Atenciosamente,
Robô"""

# Configura o e-mail
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = subject
msg.attach(MIMEText(body, 'plain'))

# Anexa o arquivo Excel
with open(output_path, 'rb') as attachment:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename=Notebooks.xlsx')
    msg.attach(part)

# Envia o e-mail
try:
    with smtplib.SMTP('smtp.office365.com', 587) as server:  # Substitua pelo servidor SMTP e porta
        server.starttls()
        server.login(sender_email, app_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
    print('E-mail enviado com sucesso!')
except Exception as e:
    print(f'Falha ao enviar e-mail: {e}')
