from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementNotInteractableException
from datetime import datetime
import openpyxl
import time
import tkinter as tk
from tkinter import messagebox
import subprocess
import threading
import queue



def show_success_popup():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    messagebox.showinfo("Sucesso", "A operação foi executada com sucesso!")
    root.destroy()



def run_automation(file_path, update_label_func=None):
    workbook = openpyxl.load_workbook(file_path)
    worksheet = workbook.active

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    Finaliza = True

    try:
        linha = 2
        coluna = 3

        driver.get("https://geridinss.dataprev.gov.br/gpa")
        driver.implicitly_wait(10)
        driver.maximize_window()

        print(f"Aguardando procedimento de Login...")
        WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/ul/li")))
        print(f"Login realizado com sucesso.")

        while Finaliza:
            servidor = worksheet.cell(row=linha, column=1).value
            Sistema = worksheet.cell(row=linha, column=coluna).value
            Subsistema = worksheet.cell(row=linha, column=coluna + 1).value
            Papel = worksheet.cell(row=linha, column=coluna + 2).value
            UO = worksheet.cell(row=linha, column=2).value
            validade = worksheet.cell(row=linha, column=coluna + 3).value
            Situacao = worksheet.cell(row=linha, column=coluna + 4).value



            if not servidor:
                driver.quit()
                print("Final!")
                show_success_popup()
                break

            if not Sistema:
                print("Avança próxima linha...")
                coluna = 3
                linha += 1                
                
                if update_label_func:
                    update_label_func(linha)
                
                continue

            if Situacao is not None:
                coluna += 5
                continue


            print("=======================================================")
            print("servidor:", servidor)
            print("Sistema:", Sistema)
            print("Subsistema:", Subsistema)
            print("Papel:", Papel)
            print("UO:", UO)
            print("validade:", validade)
            print("Situacao:", Situacao)

            driver.execute_script("document.getElementById('formMenu:btAtribAcesso').click();")
            
            
            try:
                Select(driver.find_element(By.ID, "form:sistema")).select_by_visible_text(Sistema)
            except (NoSuchElementException, ElementNotInteractableException):
                print("Sistema não localizado dentre as opções disponíveis.")
                worksheet.cell(row=linha, column=coluna + 4).value = "Sistema não localizado dentre as opções disponíveis."
                coluna += 5
                continue           

            Select(driver.find_element(By.ID, "form:subsistema")).select_by_visible_text(Subsistema)
            Select(driver.find_element(By.ID, "form:papel")).select_by_visible_text(Papel)
            Select(driver.find_element(By.ID, "form:tipoDominio")).select_by_visible_text("UO")
            driver.find_element(By.ID, "form:dominio").clear()
            driver.find_element(By.ID, "form:dominio").send_keys(UO)
            driver.find_element(By.ID, "form:usuario").clear()
            driver.find_element(By.ID, "form:usuario").send_keys(servidor)
            driver.find_element(By.ID, "form:filtrar").click()

            # APTO PARA REVALIDAÇÃO
            try:
                element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/form[2]/table/tbody/tr/td[7]")))
                data_string = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/form[2]/table/tbody/tr/td[7]").text
                data = datetime.strptime(data_string, "%d/%m/%Y")

                print("data excel:", validade)
                print("data extraída:", data)

                data_string_dt = datetime.strptime(data_string, "%d/%m/%Y")

                # Salvando a planilha
                try:
                    workbook.save(file_path)
                except Exception as e:
                    print(f"Erro ao salvar o arquivo: {e}")

                if data_string_dt < validade:
                    print("A data extraída do gerid é menor que a validade. Avançando para revalidação...")
                    driver.find_element(By.ID, "dataTableCredencial:selected").click()
                    driver.find_element(By.ID, "form2:btAlterar").click()
                    driver.find_element(By.ID, "form2:dataValidade").send_keys(Keys.CONTROL, 'a')
                    validade_str = validade.strftime("%d/%m/%Y")
                    driver.find_element(By.ID, "form2:dataValidade").send_keys(validade_str)
                    driver.find_element(By.ID, "form2:confirmar").click()

                    # VERIFICA SE DEU CERTO!================================================
                    try:
                        elements = driver.find_elements(By.XPATH, "/html/body/div[1]/div[2]/form[1]/div[2]/ul/li")
                        if len(elements) == 1:
                            teste = elements[0].text
                            if teste == "A operação foi executada com sucesso.":
                                print("A operação foi executada com sucesso.")
                                worksheet.cell(row=linha, column=coluna + 4).value = "A operação foi executada com sucesso."
                                coluna += 5
                                continue
                    except:
                        pass
                    # =====================================================================

                    # VERIFICA SE DEU CERTO!================================================
                    try:
                        elements = driver.find_elements(By.XPATH, "/html/body/div[1]/div[2]/ul/li")
                        if len(elements) == 1:
                            teste = elements[0].text
                            if teste == "Domínio não existe.":
                                print("Domínio não existe.")
                                worksheet.cell(row=linha, column=coluna + 4).value = "Domínio não existe."
                                coluna += 5
                                continue
                    except:
                        pass
                    # =====================================================================

                    # VERIFICA SE DEU CERTO!================================================
                    try:
                        elements = driver.find_elements(By.XPATH, "/html/body/div[1]/div[2]/ul/li")
                        if len(elements) == 1:
                            teste = elements[0].text
                            if teste == "A Data de Validade não deve ser superior a Data de Validade da credencial do usuário emissor.":
                                print("A Data de Validade não deve ser superior a Data de Validade da credencial do usuário emissor.")
                                worksheet.cell(row=linha, column=coluna + 4).value = "A Data de Validade não deve ser superior a Data de Validade da credencial do usuário emissor."
                                coluna += 5
                                continue
                    except:
                        pass
                    # =====================================================================
                    
                    
                    
                    # VERIFICA SE DEU CERTO!================================================
                    try:
                        elements = driver.find_elements(By.XPATH, "/html/body/div[1]/div[2]/ul/li")
                        if len(elements) == 1:
                            teste = elements[0].text
                            if teste == "Gestor de Acesso só pode atribuir acesso no seu próprio domínio ou domínio abaixo de sua abrangência.":
                                print("Gestor de Acesso só pode atribuir acesso no seu próprio domínio ou domínio abaixo de sua abrangência.")
                                worksheet.cell(row=linha, column=coluna + 4).value = "Gestor de Acesso só pode atribuir acesso no seu próprio domínio ou domínio abaixo de sua abrangência."
                                coluna += 5
                                continue
                    except:
                        pass
                    # =====================================================================                    
                    
                    

                    # VERIFICA SE DEU CERTO!================================================
                    try:
                        elements = driver.find_elements(By.XPATH, "/html/body/div[1]/div[2]/ul/li")
                        if len(elements) == 1:
                            teste = elements[0].text
                            if teste == "Não é permitido dar uma autorização a si mesmo.":
                                print("Não é permitido dar uma autorização a si mesmo.")
                                worksheet.cell(row=linha, column=coluna + 4).value = "Não é permitido dar uma autorização a si mesmo."
                                coluna += 5
                                continue
                    except:
                        pass
                    # =====================================================================

                    # salvando==============================
                    try:
                        workbook.save(file_path)
                    except Exception as e:
                        print(f"Erro ao salvar o arquivo: {e}")
                    # salvando==============================

                else:
                    worksheet.cell(row=linha, column=coluna + 4).value = f"Sistema já revalidado ({validade.strftime('%d/%m/%Y')})"
                    print(f"Sistema já revalidado ({validade.strftime('%d/%m/%Y')})")
                    # salvando==============================
                    try:
                        workbook.save(file_path)
                    except Exception as e:
                        print(f"Erro ao salvar o arquivo: {e}")
                    # salvando==============================

            except (NoSuchElementException, TimeoutException):
                print("Atribuindo novo acesso....")
                try:
                    driver.find_element(By.ID, "form2:novo").click()
                    Select(driver.find_element(By.ID, "form:sistema")).select_by_visible_text(Sistema)
                    Select(driver.find_element(By.ID, "form:subsistema")).select_by_visible_text(Subsistema)
                    Select(driver.find_element(By.ID, "form2:papel")).select_by_visible_text(Papel)
                    Select(driver.find_element(By.ID, "form2:tipoDominio")).select_by_visible_text("UO")
                    driver.find_element(By.ID, "form2:dominio").clear()
                    driver.find_element(By.ID, "form2:dominio").send_keys(UO)
                    driver.find_element(By.ID, "form2:usuario").click()
                    driver.find_element(By.ID, "form2:usuario").clear()
                    driver.find_element(By.ID, "form2:usuario").send_keys(servidor)
                    driver.find_element(By.ID, "form2:dominio").click()
                    driver.find_element(By.ID, "form2:dataValidade").click()
                    driver.find_element(By.ID, "form2:dataValidade").clear()
                    driver.find_element(By.ID, "form2:dataValidade").send_keys(Keys.CONTROL, 'a')

                    # Convertendo o objeto datetime para string no formato correto
                    validade_str = validade.strftime("%d/%m/%Y")
                    driver.find_element(By.ID, "form2:dataValidade").send_keys(validade_str)

                    # Selecionando os dias da semana
                    driver.find_element(By.ID, "form2:periodo:0").click()
                    driver.find_element(By.ID, "form2:periodo:6").click()
                    driver.find_element(By.ID, "form2:periodo:7").click()

                    # Configurando horários
                    driver.find_element(By.ID, "form2:horaAcessoInicio").click()
                    driver.find_element(By.ID, "form2:horaAcessoInicio").send_keys(Keys.CONTROL, 'a')
                    driver.find_element(By.ID, "form2:horaAcessoInicio").send_keys("0000")

                    driver.find_element(By.ID, "form2:horaAcessoFim").click()
                    driver.find_element(By.ID, "form2:horaAcessoFim").send_keys(Keys.CONTROL, 'a')
                    driver.find_element(By.ID, "form2:horaAcessoFim").send_keys("2359")

                    driver.find_element(By.ID, "form2:confirmar").click()

                    # Verificando mensagens de retorno
                    try:
                        element = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/ul/li")
                        mensagem = element.text
                        print(mensagem)
                        worksheet.cell(row=linha, column=coluna + 4).value = mensagem
                        coluna += 5
                        continue
                    except:
                        pass

                    print(f"Acesso: {UO} - {Sistema} - {Subsistema} - {Papel} - {servidor} - {validade_str}")

                    # Aguardando e verificando mensagem de sucesso
                    wait = WebDriverWait(driver, 3)
                    element = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/form[1]/div[2]/ul/li")))
                    mensagem_sucesso = element.text
                    print(mensagem_sucesso)
                    worksheet.cell(row=linha, column=coluna + 4).value = mensagem_sucesso

                    driver.execute_script("document.getElementById('formMenu:btAtribAcesso').click()")
                    print()

                    # Salvando a planilha
                    try:
                        workbook.save(file_path)
                    except Exception as e:
                        print(f"Erro ao salvar o arquivo: {e}")

                    coluna += 5
                    continue

                except Exception as e:
                    print(f"Erro ao atribuir novo acesso: {e}")
                    worksheet.cell(row=linha, column=coluna + 4).value = f"Erro ao atribuir acesso: {str(e)}"
                    workbook.save(file_path)
                    coluna += 5
                    continue

    except Exception as e:
        print(f"Erro geral: {e}")
    finally:
        driver.quit()
