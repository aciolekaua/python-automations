import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui

navegador = webdriver.Chrome()
url = "https://efisco.sefaz.pe.gov.br/sfi_com_sca/PRMontarMenuAcesso"
navegador.get(url)
navegador.maximize_window()
time.sleep(5)
pyautogui.alert("Selecione o certificado desejado!!")
try:
    wait = WebDriverWait(navegador,10)
    certificado =  wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/form/div/div[1]/div/div[2]/div/div[3]/div/div/div/div/div/div/div/div[2]/div[3]/button/span")))
    certificado.click()
    buscar_notas_form = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/form/div[2]/div/div[1]/div/div/input")))
    buscar_notas_form.click()
    buscar_notas_input = wait.until(EC.visibility_of_element_located((By.XPATH,"/html/body/form/div[2]/div/div[1]/div/div/div/div/div/div[1]/div[1]/div/input")))
    buscar_notas_input.send_keys("nfce")
    buscar_notas_input.send_keys(Keys.ENTER)
    nfce_link = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/form/div[2]/div/div[1]/div/div/div/div/div/div[2]/div/div[1]/div/ul/li/a")))
    nfce_link.click()
    time.sleep(5)
    pyautogui.click(x=272, y=180,clicks=1)
    pyautogui.write("01062024",interval=0.25)
    pyautogui.press("tab")
    pyautogui.write("30062024",interval=0.25)
    pyautogui.press("tab")
    pyautogui.write("022560378",interval=0.25)
    pyautogui.alert("Resolva o captcha!")
    time.sleep(50)
    pyautogui.click(x=640, y=423)
    time.sleep(15)
finally:
    navegador.quit()
    #print(pyautogui.position())
    #time.sleep(1000)