from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
 
# options = Options()
# options.add_argument('--headless=chrome')
# options.add_argument('--no-sandbox')
# options.add_argument('--disable-dev-shm-usage')
# driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
 


def servico():
    servico = Service(ChromeDriverManager().install())

    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    
    navegador = webdriver.Chrome(service=servico, options=options)
   
    return navegador

def acharRecibo(entrada,sem,pedido):
    navegador = servico()

    navegador.get("https://conveniar.finatec.org.br/Fundacao/Login.aspx?ReturnUrl=%2fFundacao%2f")
    navegador.find_element('xpath','//*[@id="ctl00_ContentPlaceHolder1_ObjWucLoginCaptcha_lgUsuario_UserName"]').send_keys(f"{entrada}")
    navegador.find_element('xpath','//*[@id="ctl00_ContentPlaceHolder1_ObjWucLoginCaptcha_lgUsuario_Password"]').send_keys(f"{sem}")
    navegador.find_element('xpath','//*[@id="ctl00_ContentPlaceHolder1_ObjWucLoginCaptcha_lgUsuario_btnLogin"]').click()
    navegador.implicitly_wait(30)
    select = Select(navegador.find_element('xpath','//*[@id="ctl00_ddlModulos"]'))
    select.select_by_value('4')
    navegador.implicitly_wait(5)
    navegador.get("https://conveniar.finatec.org.br/Fundacao/Forms/Convenio/ControleFinanceiro.aspx")
    #pagina controle financeiro
    navegador.find_element('xpath','//*[@id="ctl00_ContentPlaceHolder1_FiltroBaixaLancamentoUserControl1_txtValor"]').send_keys(f"{pedido}")
    select = Select(navegador.find_element('xpath','//*[@id="ctl00_ContentPlaceHolder1_FiltroBaixaLancamentoUserControl1_ddlStatus"]'))
    select.select_by_visible_text('Todos')
    navegador.implicitly_wait(5)
    #aplicarfiltro
    navegador.find_element('xpath','//*[@id="ctl00_ContentPlaceHolder1_FiltroBaixaLancamentoUserControl1_btnFiltrar"]').click()
    navegador.implicitly_wait(5)
    navegador.find_element('xpath','//*[@id="ctl00_ContentPlaceHolder1_gvPrincipal_ctl02_lbtEditar"]').click()
    navegador.implicitly_wait(5)
    navegador.find_element('xpath','//*[@id="ctl00_ContentPlaceHolder1_ControleFinanceiroUserControl2_btnAnexos"]/span[2]').click()
    navegador.implicitly_wait(5)
    #recibo
    navegador.find_element('xpath','//*[@id="btnReciboPedido"]/span').click()

    #pdf :
    # navegador.find_element('xpath','//*[@id="icon"]/iron-icon').click()
    # navegador.get('xpath','//*[@id="main-content"]/a')
    #navegador.find_element('xpath','//*[@id="main-content"]/a').click()



    #pdf_element = navegador.find_element_by_xpath('//*[@id="main-content"]/a')
    pdf_element = WebDriverWait(navegador, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="iModalResponsive"]'))
    )
    # Extract the URL of the PDF file
    pdf_url = pdf_element.get_attribute("src")
    print(pdf_url)

    navegador.close()

    return pdf_url
