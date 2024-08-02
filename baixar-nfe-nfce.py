import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui

# Listas de empresas para cada contador
empresas_veridiana = [
    {"nomeEmpresa": "Empresa1", "InscricaoEstadual": "022560378"},
    {"nomeEmpresa": "Empresa2", "InscricaoEstadual": "080511430"},
    # Adicione mais empresas conforme necessário
]

empresas_roberto = [
    {"nomeEmpresa": "EmpresaA", "InscricaoEstadual": "044780591"},
    {"nomeEmpresa": "EmpresaB", "InscricaoEstadual": "055890612"},
    # Adicione mais empresas conforme necessário
]

def on_submit():
    colaborador = colaborador_var.get()
    contador = contador_var.get()
    data_inicial = data_inicial_entry.get()
    data_final = data_final_entry.get()
    
    if colaborador == "Digite outro":
        colaborador = outro_colaborador_var.get()

    if not colaborador or not contador or not data_inicial or not data_final:
        messagebox.showerror("Erro", "Por favor, preencha todos os campos")
    else:
        if contador == "Veridiana":
            empresas = empresas_veridiana
        else:
            empresas = empresas_roberto
        executar_automacao(colaborador, contador, data_inicial, data_final, empresas)

def executar_automacao(colaborador, contador, data_inicial, data_final, empresas):
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
        pyautogui.click(x=272, y=180, clicks=1)
        pyautogui.write(data_inicial.replace("/", ""), interval=0.25)  # Data inicial
        pyautogui.press("tab")
        pyautogui.write(data_final.replace("/", ""), interval=0.25)  # Data final
        pyautogui.press("tab")

        for i, empresa in enumerate(empresas):
            if i == 0:
                pyautogui.write(empresa["InscricaoEstadual"], interval=0.25)
                pyautogui.alert("Resolva o captcha!")
                time.sleep(50)
                pyautogui.click(x=640, y=423)
            else:
                pyautogui.click(x=299, y=150)
                pyautogui.hotkey("ctrl", "a")
                pyautogui.write(empresa["InscricaoEstadual"], interval=0.25)
                pyautogui.click(x=663, y=292)
            time.sleep(15)
    finally:
        navegador.quit()
        # Exibir mensagem de sucesso após a automação
        messagebox.showinfo("Sucesso", f"Automação concluída!\nColaborador: {colaborador}\nContador: {contador}\nData Inicial: {data_inicial}\nData Final: {data_final}")

app = tk.Tk()
app.title("Seleção de Colaborador e Período")

# Campo de seleção de colaborador
tk.Label(app, text="Selecione o Colaborador:").grid(row=0, column=0, padx=10, pady=10)
colaborador_var = tk.StringVar(value="Rebeka")
colaboradores = ["Rebeka", "Digite outro"]
colaborador_menu = ttk.Combobox(app, textvariable=colaborador_var, values=colaboradores, state="readonly")
colaborador_menu.grid(row=0, column=1, padx=10, pady=10)

# Campo para digitar outro colaborador
outro_colaborador_var = tk.StringVar()
outro_colaborador_entry = tk.Entry(app, textvariable=outro_colaborador_var)
outro_colaborador_entry.grid(row=1, column=1, padx=10, pady=10)
outro_colaborador_entry.grid_remove()

def on_colaborador_change(event):
    if colaborador_var.get() == "Digite outro":
        outro_colaborador_entry.grid()
    else:
        outro_colaborador_entry.grid_remove()

colaborador_menu.bind("<<ComboboxSelected>>", on_colaborador_change)

# Campo de seleção de contador
tk.Label(app, text="Selecione o Contador:").grid(row=2, column=0, padx=10, pady=10)
contador_var = tk.StringVar(value="Veridiana")
contadores = ["Veridiana", "Roberto"]
contador_menu = ttk.Combobox(app, textvariable=contador_var, values=contadores, state="readonly")
contador_menu.grid(row=2, column=1, padx=10, pady=10)

# Campo de data inicial
tk.Label(app, text="Data Inicial:").grid(row=3, column=0, padx=10, pady=10)
data_inicial_entry = DateEntry(app, date_pattern='dd/mm/yyyy')
data_inicial_entry.grid(row=3, column=1, padx=10, pady=10)

# Campo de data final
tk.Label(app, text="Data Final:").grid(row=4, column=0, padx=10, pady=10)
data_final_entry = DateEntry(app, date_pattern='dd/mm/yyyy')
data_final_entry.grid(row=4, column=1, padx=10, pady=10)

# Botão de submissão
submit_button = tk.Button(app, text="Enviar", command=on_submit)
submit_button.grid(row=5, columnspan=2, pady=20)

app.mainloop()