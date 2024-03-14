from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import PySimpleGUI as sg
import time

def changePassword(niss,oldPassword):

    newPassword = "Passwordtemporary12!"
    
    updatepassword(niss,oldPassword,newPassword)
    
    updatepassword(niss,newPassword,oldPassword)
    
def read_excel_data(filepath):
    try:
        # Read Excel file
        df = pd.read_excel(filepath)
        # Extract 'NISS' and 'Password' columns
        niss_values = df['NISS'].tolist()
        password_values = df['Password'].tolist()
        return niss_values, password_values
    except Exception as e:
        print("Error reading Excel file:", e)
        return None, None

def updatepassword(niss,oldPassword,newPassword):
     # Start a new instance of Chrome WebDriver
    driver = webdriver.Chrome()

    # Open the login page
    driver.get("https://app.seg-social.pt/sso/login?service=https%3A%2F%2Fapp.seg-social.pt%2Fptss%2Fcaslogin")

    # Find username/email input field and enter NISS
    username_field = driver.find_element(By.ID, "username")
    username_field.send_keys(niss)

    # Find password input field and enter password
    password_field = driver.find_element(By.ID, "password")
    password_field.send_keys(oldPassword)

    # Submit the login form
    submit_button = driver.find_element(By.XPATH, "//input[@name='submitBtn']")
    submit_button.click()
 
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "frawSearchForm:operationSearch_input")))

    # Find the input field
    input_field = driver.find_element(By.ID, "frawSearchForm:operationSearch_input")

    # Clear any existing text in the input field (if needed)
    input_field.clear()

    # Enter the text you want to search for
    input_field.send_keys("Alterar Palavra-passe")

    # Find the "Pesquisar" button and click on it
    pesquisar_button = driver.find_element(By.ID, "frawSearchForm:frawPesquisa")
    pesquisar_button.click()
    
   

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@id, ':0:title')]")))

    
    link = driver.find_element(By.XPATH, "//*[contains(@id, ':0:title')]")

        
    link.click()

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "oldPassword")))
    driver.refresh()
    driver.refresh()
    driver.refresh()
    # Localize cada campo de entrada pelo ID
    old_password_input = driver.find_element(By.ID,"oldPassword")
    new_password_input = driver.find_element(By.ID,"newPassword")
    confirm_new_password_input = driver.find_element(By.ID,"confirmNewPassword")

    # Agora você pode interagir com esses campos de entrada, por exemplo, preenchendo-os com dados:
    old_password_input.send_keys(oldPassword)
    new_password_input.send_keys(newPassword)
    confirm_new_password_input.send_keys(newPassword)

    if(newPassword != "Passwordtemporary12!"):
        #command to end this 
        button_element = driver.find_element(By.XPATH, "//div[@class='row btn-row']//input[@type='submit']")
        button_element.click()
        print("Senha alterada com sucesso!")
    else:
        button_element = driver.find_element(By.XPATH, "//div[@class='row btn-row']//input[@type='submit']")
        button_element.click()
        print("Senha alterada com sucesso!")


def main():
    layout = [
        [sg.Text("Selecione o arquivo Excel:")],
        [sg.InputText(key="file_path", disabled=True), sg.FileBrowse()],
        [sg.Button("OK"), sg.Button("Cancelar")]
    ]
    
    layout2 = [
        [sg.Text("")],
    ]

    window = sg.Window("Escolher Arquivo Excel", layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "Cancelar":
            break
        elif event == "OK":
            filepathexcel = values["file_path"]
            if filepathexcel:
                niss_values, password_values = read_excel_data(filepathexcel)
                if niss_values is None or password_values is None:
                    sg.popup_error("Erro ao ler o arquivo Excel.")
                    break
                else:
                    for niss, password in zip(niss_values, password_values):
                        print("NISS value:", niss)
                        print("Password value:", password)
                        changePassword(niss, password)
                    break
            else:
                sg.popup_error("Por favor, selecione um arquivo Excel.")

    window.close()

if __name__ == "__main__":
    main()

