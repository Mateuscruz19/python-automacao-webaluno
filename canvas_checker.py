from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.firefox import GeckoDriverManager
from dotenv import load_dotenv
from excel_writer import criar_excel_frequencia
import os
import time

# Carrega as variáveis de ambiente
load_dotenv()

class WebAlunoChecker:
    def __init__(self):
        self.url = "https://432f5d.mannesoftprime.com.br/webaluno/"
        self.setup_driver()

    def setup_driver(self):
        """Configura o driver do Firefox com as opções necessárias"""
        try:
            firefox_options = Options()
            # firefox_options.add_argument("--headless")  # Descomente para executar sem interface gráfica
            firefox_options.add_argument("--start-maximized")
            
            # Configuração do GeckoDriver
            driver_path = GeckoDriverManager().install()
            service = Service(executable_path=driver_path)
            
            self.driver = webdriver.Firefox(
                service=service,
                options=firefox_options
            )
            print("Driver configurado com sucesso!")
            
        except Exception as e:
            print(f"Erro ao configurar o driver: {str(e)}")
            raise

    def login(self, username, password):
        """Realiza o login no WebAluno através da Microsoft"""
        try:
            print("Acessando a página do WebAluno...")
            self.driver.get(self.url)
            time.sleep(2)  # Reduzido de 5 para 2
            
            print("Aceitando cookies...")
            # Aguarda e clica no botão de aceitar cookies
            cookie_button = WebDriverWait(self.driver, 10).until(  # Reduzido de 20 para 10
                EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))
            )
            cookie_button.click()
            time.sleep(1)  # Reduzido de 2 para 1
            
            print("Clicando no botão de login da Microsoft...")
            # Aguarda e clica no botão da Microsoft
            microsoft_button = WebDriverWait(self.driver, 10).until(  # Reduzido de 20 para 10
                EC.element_to_be_clickable((By.XPATH, '//*[@id="MICROSOFT_ASC"]/div/button'))
            )
            microsoft_button.click()
            time.sleep(2)  # Reduzido de 3 para 2
            
            print("Preenchendo email...")
            # Aguarda o campo de email da Microsoft
            email_field = WebDriverWait(self.driver, 10).until(  # Reduzido de 20 para 10
                EC.presence_of_element_located((By.ID, "i0116"))
            )
            email_field.clear()  # Limpa o campo antes de digitar
            email_field.send_keys(username)
            
            print("Clicando em próximo...")
            # Clica no botão próximo
            next_button = WebDriverWait(self.driver, 5).until(  # Reduzido de 10 para 5
                EC.element_to_be_clickable((By.ID, "idSIButton9"))
            )
            next_button.click()
            time.sleep(2)  # Reduzido de 3 para 2
            
            print("Preenchendo senha...")
            # Aguarda o campo de senha
            password_field = WebDriverWait(self.driver, 10).until(  # Reduzido de 20 para 10
                EC.presence_of_element_located((By.ID, "i0118"))
            )
            password_field.clear()  # Limpa o campo antes de digitar
            password_field.send_keys(password)
            
            print("Clicando em entrar...")
            # Clica no botão de login
            sign_in_button = WebDriverWait(self.driver, 5).until(  # Reduzido de 10 para 5
                EC.element_to_be_clickable((By.ID, "idSIButton9"))
            )
            sign_in_button.click()
            
            print("Verificando botão 'Não' para manter conectado...")
            # Aguarda e clica no botão "Não" para manter conectado
            try:
                stay_signed_in = WebDriverWait(self.driver, 3).until(  # Reduzido de 5 para 3
                    EC.element_to_be_clickable((By.ID, "idBtn_Back"))
                )
                stay_signed_in.click()
                print("Clicou em 'Não' para manter conectado")
            except:
                print("Botão 'Não' não encontrado, continuando...")
            
            # Aguarda o carregamento da página do WebAluno
            time.sleep(3)  # Reduzido de 5 para 3
            
            return True
            
        except Exception as e:
            print(f"Erro durante o login: {str(e)}")
            return False

    def navegar_para_frequencia(self):
        """Navega até a página de frequência"""
        try:
            print("Abrindo menu de navegação...")
            # Aguarda a página carregar completamente
            time.sleep(2)  # Reduzido de 5 para 2
            
            # Tenta encontrar o botão de navegação usando JavaScript
            print("Procurando botão de navegação...")
            nav_button = self.driver.execute_script("""
                return document.querySelector('#myNavbar span');
            """)
            
            if nav_button:
                print("Botão encontrado, tentando clicar...")
                # Usa JavaScript para clicar no botão
                self.driver.execute_script("arguments[0].click();", nav_button)
                time.sleep(1)  # Reduzido de 2 para 1
                
                print("Procurando link de frequência...")
                # Tenta diferentes métodos para encontrar o link de frequência
                frequencia_link = None
                
                # Método 1: Usando XPath
                try:
                    frequencia_link = WebDriverWait(self.driver, 3).until(  # Reduzido de 5 para 3
                        EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/nav/div[2]/div[2]/ul/li[4]/a'))
                    )
                except:
                    print("XPath não encontrou o link")
                
                if frequencia_link:
                    print("Link de frequência encontrado, tentando clicar...")
                    try:
                        # Tenta clicar usando JavaScript
                        self.driver.execute_script("""
                            arguments[0].scrollIntoView(true);
                            arguments[0].click();
                        """, frequencia_link)
                        time.sleep(2)  # Reduzido de 3 para 2
                        print("Navegação para frequência concluída!")
                        return True
                            
                    except Exception as e:
                        print(f"Erro ao tentar clicar no link: {str(e)}")
                        return False
                else:
                    print("Link de frequência não encontrado")
                    return False
            else:
                print("Botão de navegação não encontrado")
                return False
            
        except Exception as e:
            print(f"Erro ao navegar para frequência: {str(e)}")
            return False

    def coletar_dados_frequencia(self):
        """Coleta os dados de frequência de todas as matérias"""
        try:
            print("Coletando dados de frequência...")
            time.sleep(2)  # Reduzido de 5 para 2
            
            # Lista para armazenar os dados das matérias
            materias = []
            
            # Encontra a tabela principal
            tabela = self.driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/table")
            print("Tabela principal encontrada")
            
            # Pega todas as linhas da tabela (exceto o cabeçalho)
            linhas = tabela.find_elements(By.XPATH, ".//tbody/tr")
            print(f"Encontradas {len(linhas)} linhas na tabela")
            
            # Processa cada linha
            for linha in linhas:
                try:
                    # Pega o nome da matéria (primeira coluna)
                    nome_materia = linha.find_element(By.XPATH, ".//td[1]").text.strip()
                    
                    # Pega a carga horária (segunda coluna)
                    carga_horaria = linha.find_element(By.XPATH, ".//td[2]").text.strip()
                    
                    # Pega as faltas (terceira coluna)
                    faltas = linha.find_element(By.XPATH, ".//td[3]").text.strip()
                    
                    # Pega a frequência (quarta coluna)
                    frequencia = linha.find_element(By.XPATH, ".//td[4]").text.strip()
                    
                    # Só adiciona se tiver todos os dados
                    if nome_materia and carga_horaria and faltas and frequencia:
                        materias.append({
                            "nome": nome_materia,
                            "carga_horaria": carga_horaria,
                            "faltas": faltas,
                            "frequencia": frequencia
                        })
                        
                        print(f"\nMatéria: {nome_materia}")
                        print(f"Carga Horária: {carga_horaria}")
                        print(f"Faltas: {faltas}")
                        print(f"Frequência: {frequencia}")
                        print("-" * 50)
                    
                except Exception as e:
                    continue
            
            if materias:
                print(f"\nTotal de matérias encontradas: {len(materias)}")
            else:
                print("Nenhuma matéria encontrada")
            
            return materias
            
        except Exception as e:
            print(f"Erro ao coletar dados de frequência: {str(e)}")
            return []

    def calcular_status_frequencia(self, materias):
        """Calcula o status de aprovação por frequência e faltas restantes"""
        print("\n=== ANÁLISE DE FREQUÊNCIA ===")
        print("=" * 50)
        
        for materia in materias:
            try:
                # Limpa e converte os valores para números
                carga_horaria = float(materia['carga_horaria'].replace(',', '.'))
                faltas = float(materia['faltas'].replace(',', '.'))
                frequencia = float(materia['frequencia'].replace('%', '').replace(',', '.'))
                
                # Calcula o máximo de faltas permitidas (25% da carga horária)
                max_faltas = carga_horaria * 0.25
                
                # Calcula quantas faltas ainda são permitidas
                faltas_restantes = max_faltas - faltas
                
                # Determina o status
                if faltas > max_faltas:
                    status = "REPROVADO POR FALTA"
                else:
                    status = "APROVADO POR FREQUÊNCIA"
                
                print(f"\nMatéria: {materia['nome']}")
                print(f"Carga Horária Total: {carga_horaria:.1f} horas")
                print(f"Faltas Atuais: {faltas:.1f} horas")
                print(f"Frequência Atual: {frequencia:.1f}%")
                print(f"Máximo de Faltas Permitidas: {max_faltas:.1f} horas")
                print(f"Faltas Restantes Permitidas: {faltas_restantes:.1f} horas")
                print(f"Status: {status}")
                print("-" * 50)
                
            except Exception as e:
                print(f"Erro ao calcular status para {materia['nome']}: {str(e)}")
                print(f"Dados originais: carga_horaria={materia['carga_horaria']}, faltas={materia['faltas']}, frequencia={materia['frequencia']}")
                continue

    def close(self):
        """Fecha o navegador"""
        self.driver.quit()

def main():
    # Obtém as credenciais das variáveis de ambiente
    username = os.getenv('WEALUNO_USERNAME')
    password = os.getenv('WEALUNO_PASSWORD')
    
    if not username or not password:
        print("Erro: Credenciais não encontradas nas variáveis de ambiente.")
        print("Por favor, configure WEALUNO_USERNAME e WEALUNO_PASSWORD no arquivo .env")
        return
    
    checker = WebAlunoChecker()
    
    try:
        if checker.login(username, password):
            print("Login realizado com sucesso!")
            if checker.navegar_para_frequencia():
                print("Navegação para frequência realizada com sucesso!")
                materias = checker.coletar_dados_frequencia()
                if materias:
                    print(f"\nTotal de matérias encontradas: {len(materias)}")
                    checker.calcular_status_frequencia(materias)
                    # Cria o arquivo Excel
                    criar_excel_frequencia(materias)
                else:
                    print("Nenhuma matéria encontrada")
            else:
                print("Falha ao navegar para frequência")
        else:
            print("Falha no login")
    finally:
        checker.close()

if __name__ == "__main__":
    main() 