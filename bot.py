import customtkinter as ctk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import threading
import time

# Configurações fixas
CREDENTIALS_PATH = 'pragmatic-braid-429419-n6-887ce9594d96.json'
SPREADSHEET_ID = '1tvmZ88h3KQz2HHFt4IKrkrPqlTI2HNkOAIGEGmV6fFI'

# Função de automação
def run_automation(username, password, log_widget):
    log_widget.insert(ctk.END, "Inicializando automação...\n")

    try:
        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico)

        # API do Google Sheets
        scope = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_PATH, scope)
        client = gspread.authorize(creds)

        planilha = client.open_by_key(SPREADSHEET_ID)
        sheet = planilha.get_worksheet(0)

        # Leitura dos dados
        dados = sheet.col_values(1)
        log_widget.insert(ctk.END, "Dados da planilha lidos com sucesso.\n")

        # Acesso
        navegador.get("https://sip.pi.gov.br/sip/login.php?sigla_orgao_sistema=GOV-PI&sigla_sistema=SEI&infra_url=L3NlaS8=")

        # Login
        navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div[2]/form/div/div[2]/input[1]').send_keys(username)
        navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div[2]/form/div/div[2]/input[2]').send_keys(password)

        # Selecionar a instituição no dropdown
        navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div[2]/form/div/div[2]/select[1]').click()
        dropdown_element = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[3]/div/div[2]/form/div/div[2]/select[1]'))
        )
        select = Select(dropdown_element)
        select.select_by_visible_text("SESAPI-PI")
        navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div[2]/form/div/div[4]/div[2]/button').click()

        # Iterar pelos dados e simular Enter
        for dado in dados:
            input_element = navegador.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[4]/span/form/input')
            input_element.clear()
            input_element.send_keys(dado)
            input_element.send_keys(Keys.ENTER)
            
          
            time.sleep(1)  
            log_widget.insert(ctk.END, f"Acessado concluido: {dado}\n")

        log_widget.insert(ctk.END, "Automação concluída.\n")
        navegador.quit()

    except Exception as e:
        log_widget.insert(ctk.END, f"Erro: {str(e)}\n")

# Função de iniciar automação em thread separada
def start_automation():
    username = entry_username.get()
    password = entry_password.get()
    log_widget.delete("1.0", ctk.END)
    threading.Thread(target=run_automation, args=(username, password, log_widget)).start()

# Configuração da interface gráfica
ctk.set_appearance_mode("light")
app = ctk.CTk()
app.title("Automação - SEI")
app.geometry("800x400")

# Criar frames para a organização do layout
frame_left = ctk.CTkFrame(app)
frame_left.pack(side="left", padx=20, pady=20, fill="y")

frame_right = ctk.CTkFrame(app)
frame_right.pack(side="right", padx=20, pady=20, fill="both", expand=True)

# Widgets do frame esquerdo
label_username = ctk.CTkLabel(frame_left, text="Usuário SEI:")
label_username.pack(pady=4)
entry_username = ctk.CTkEntry(frame_left, width=400)
entry_username.pack(pady=4)

label_password = ctk.CTkLabel(frame_left, text="Senha:")
label_password.pack(pady=4)
entry_password = ctk.CTkEntry(frame_left, width=400, show="*")
entry_password.pack(pady=4)

button_start = ctk.CTkButton(frame_left, text="Iniciar Automação", command=start_automation)
button_start.pack(pady=20)

# Widgets do frame direito
log_widget = ctk.CTkTextbox(frame_right, width=500, height=350)
log_widget.pack(pady=5, fill="both", expand=True)

app.mainloop()
