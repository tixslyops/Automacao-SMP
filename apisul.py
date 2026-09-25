from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas as pd
import time
import sys


print("INICIANDO AUTOMAÇÃO!")
smp_excel = r"\\enterprise.ad\dados\ELD_TRANSPORTADORA\09 - Torre de Controle\09 - Rotinas Diárias\07 - Conferencia\01 - Conferencia Alocação X SMP.xlsm"
aba_alocacao = "ALOCAÇÃO"
aba_smp = "BASE SMP APISUL"
coluna_fazenda = "Fazenda"
coluna_placa_aloc = "Placa"
coluna_placa_smp = "Placa"
coluna_frente = "Frente"
coluna_operacao = "Área Operacional"

print("Buscando placas...")

df_excluir = pd.read_excel(smp_excel, sheet_name=aba_alocacao)
df_destino = pd.read_excel(smp_excel, sheet_name=aba_smp)

lista_alocacao = df_excluir[coluna_placa_aloc].dropna().unique().tolist()
lista_smp = df_destino[coluna_placa_smp].dropna().unique().tolist()

sem_smp = [i for i in lista_alocacao if i not in lista_smp]
df_criar_smp = df_excluir[df_excluir[coluna_placa_aloc].isin(sem_smp)]

if df_criar_smp.empty:
    print("Todas as frotas já possuem SMP cadastrada. A automação encerrou.")
    sys.exit()

placas_pendentes = df_criar_smp[coluna_placa_aloc].dropna().unique().tolist()
print(f"SMPs para criar: {len(df_criar_smp)}")
print(f"Placas que serão criadas: {placas_pendentes}")
print("================================================")

def fazer_login(driver_bot):
    print("Executando rotina de LOGIN...")
    driver_bot.get("https://novoapisullog.apisul.com.br/Login")
    try:
        usuario_input = WebDriverWait(driver_bot, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='txtUsuario']")))
        usuario_input.clear()
        usuario_input.send_keys("leticia.reis")

        senha_input = WebDriverWait(driver_bot, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='txtSenha']")))
        senha_input.clear()
        senha_input.send_keys("L3ticia@01" + Keys.ENTER)
        time.sleep(5)
        print("Login efetuado com sucesso!")
    except Exception as e:
        print(f"Erro ao tentar realizar o login: {e}")

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico, options=options)

servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico, options=options)

fazer_login(driver)

for index, linha in df_criar_smp.iterrows():
    placa = str(linha[coluna_placa_aloc]).strip().upper()
    fazenda = linha[coluna_fazenda]
    frente = str(linha[coluna_frente])
    area_operacional = str(linha[coluna_operacao])
    data_atual = datetime.now().strftime('%d/%m/%Y 00:00')

    print(f"[PROCESSANDO] Placa: {placa} ({index + 1}/{len(df_criar_smp)})")

    if "Login" in driver.current_url or len(driver.find_elements(By.XPATH, "//*[@id='menu1']/span")) == 0:
        print("[ALERTA] Sessão expirada ou página de Login detectada. Refazendo login...")
        fazer_login(driver)

    try:
  
        print("Acessando o menu e selecionando a opção de SMP...")
        
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='menu1']/span"))).click()
        time.sleep(3)

        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_MenuPrincipal_i1_i0_ctl00']/ul/li[4]/div/ul/li[5]/a"))).click()
        time.sleep(4)

        print("Clicando em Novo Formulário...")
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='btnNovo']"))).click()
        time.sleep(4)

       
        campo_data = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_MainContent_txtDataInicioViagem_dateInput']")))
        actions = ActionChains(driver)
        actions.move_to_element(campo_data).click().perform()
        campo_data.send_keys(Keys.CONTROL + "a")
        campo_data.send_keys(Keys.BACKSPACE)
        time.sleep(0.5)
        campo_data.send_keys(data_atual)
        campo_data.send_keys(Keys.TAB)
        time.sleep(3)

        
        campo_operacao = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_MainContent_cmbTipoOperacao_Input']")))  
        campo_operacao.click()
        campo_operacao.send_keys(area_operacional)
        
        try:
            opcao_op = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '{area_operacional}')]")))
            opcao_op.click()
        except Exception:
            print(f"Não foi possível selecionar a opção {area_operacional}")
        time.sleep(3)

        
        campo_trp = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_MainContent_txtEmitenteTransportadora_Input']")))
        campo_trp.click()
        campo_trp.send_keys(frente)
        
        try:
            opcao_trp = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '{frente}')]")))
            opcao_trp.click()
        except Exception:
             print(f"Não foi possível selecionar a opção: {frente}")
        time.sleep(3)

        
        campo_veic = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='txtVeiculo_Input']")))
        campo_veic.click()
        time.sleep(1)
        campo_veic.send_keys(placa)
        time.sleep(3)
        
        try:
            opcao_veic = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(text(), '{placa}')]")))
            opcao_veic.click()
        except Exception:
            print(f"Não foi possível selecionar a opção {placa}")
        time.sleep(2)

        btn_incluir = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_MainContent_btnVinculoVeiculo']")))
        btn_incluir.click()
        time.sleep(5)

        
        vinc_pontos = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_MainContent_gridPontosVinculados_ctl00_ctl02_ctl00_lnkPontoExistente']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", vinc_pontos)
        time.sleep(2)
        driver.execute_script("arguments[0].click();", vinc_pontos)

        identificador = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='rcbIdentificadorPonto_Input']")))
        identificador.click()
        time.sleep(1)
        identificador.send_keys("ELDORADO FABRICA")
        time.sleep(5)

        try:
            opcao_eldorado = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@class, 'rcbItem') or contains(@class, 'rcbHovered')][contains(text(), 'ELDORADO FABRICA')]")))
            opcao_eldorado.click()
            time.sleep(3)
        except Exception:
            print("Não foi possível clicar na sugestão para 'ELDORADO FABRICA'.")

        btn_salvar = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_gridPontosVinculados_ctl00_ctl02_ctl02_btnSalvarPontoSMP")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
        time.sleep(2)
        driver.execute_script("arguments[0].click();", btn_salvar)
        time.sleep(8)

        
        vinc_pontos = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_MainContent_gridPontosVinculados_ctl00_ctl02_ctl00_lnkPontoExistente']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", vinc_pontos)
        time.sleep(2)
        driver.execute_script("arguments[0].click();", vinc_pontos)

        identificador = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='rcbIdentificadorPonto_Input']")))
        identificador.click()
        time.sleep(1)
        identificador.send_keys(fazenda)
        time.sleep(5)

        try:
            opcao_fazenda = WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.XPATH, f"//*[contains(@class, 'rcbItem') or contains(@class, 'rcbHovered')][contains(text(), '{fazenda}')]")))
            opcao_fazenda.click()
            time.sleep(1.5)
        except Exception:
            identificador.send_keys(Keys.ENTER)
            time.sleep(2)

        
        data_chegada = datetime.now().strftime('%d/%m/%Y 23:00')
        campo_previsao = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_MainContent_gridPontosVinculados_ctl00_ctl02_ctl02_txtPrevisaoChegada_dateInput']")))
        actions = ActionChains(driver)
        actions.move_to_element(campo_previsao).click().perform()
        campo_previsao.send_keys(Keys.CONTROL + "a")
        campo_previsao.send_keys(Keys.BACKSPACE)
        time.sleep(0.5)
        campo_previsao.send_keys(data_chegada)
        campo_previsao.send_keys(Keys.TAB)
        time.sleep(2)

        btn_salvar = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_gridPontosVinculados_ctl00_ctl02_ctl02_btnSalvarPontoSMP")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
        time.sleep(2)
        driver.execute_script("arguments[0].click();", btn_salvar)
        time.sleep(8)    

        
        print("GERANDO ROTA DINÂMICA...")
        dinamica = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='MainContent_rblTipoRota_0']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dinamica)
        time.sleep(2)
        driver.execute_script("arguments[0].click();", dinamica)
        time.sleep(3)

        gerar_rota = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='MainContent_btnGerarRota']")))
        driver.execute_script("arguments[0].  click();", gerar_rota)
        time.sleep(4)

        
        print("CONFIGURANDO SMP AGENDADA...")
        smp_agendada = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='MainContent_smpAgendada_txtNome']")))
        smp_agendada.click()
        smp_agendada.send_keys(placa)
        time.sleep(2)

        scroll_dias = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='MainContent_smpAgendada_rblTipoSMPAgendada_2']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", scroll_dias)
        time.sleep(2)
        driver.execute_script("arguments[0].click();", scroll_dias)
        time.sleep(3)

        print("Selecionando dias da semana...")
        for i in range(1, 8):
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='ctl00_MainContent_smpAgendada_rtbDiasSemana']/div/div/div/ul/li[{i}]/a"))).click()
            time.sleep(1)

        print("Configurando Prioridade Alta...")
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_MainContent_smpAgendada_cmbTipoPrioridade_Input']"))).click()
        time.sleep(2)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_MainContent_smpAgendada_cmbTipoPrioridade_DropDown']/div/ul/li[4]"))).click()
        time.sleep(4)

        print("Salvando SMP Agendada...")
        salvar_smp = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_MainContent_btnNovo']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", salvar_smp)
        time.sleep(2)
        driver.execute_script("arguments[0].click();", salvar_smp)

        print(f"[SUCESSO] SMP Agendada criada para a placa '{placa}'!")
        time.sleep(5)

    except Exception as e:
        print(f"[ERRO] Ocorreu uma falha ao processar a placa {placa}: {e}")
        print("Aguardando 5 segundos e tentando passar para a próxima placa...")
        time.sleep(5)
        driver.get("https://novoapisullog.apisul.com.br/Home") 
        time.sleep(3)

print("\nAutomação concluída para todas as frotas pendentes!")
driver.quit()
