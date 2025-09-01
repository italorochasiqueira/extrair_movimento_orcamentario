import time
import sys
import calendar
import os
import glob
import win32com.client as win32
import win32com.client.gencache
import shutil
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from tkinter import messagebox

# Limpa cache corrompido do win32com
cache_folder = Path(win32com.client.gencache.GetGeneratePath()).parent
shutil.rmtree(cache_folder, ignore_errors=True)

def extrair_orcamento_atena_plano_pp(usuario: str, senha: str, data_ini: str, data_fim: str):

    # Extrair mês e ano das strings recebidas
    mes_inicio, ano_inicio = map(int, data_ini.split("/"))
    mes_fim, ano_fim = map(int, data_fim.split("/"))

    # Criar datas para o loop
    data_atual = datetime(ano_inicio, mes_inicio, 1)
    data_limite = datetime(ano_fim, mes_fim, 1)

    print(f"[DEBUG - Extrair Orçamento] Meses para apuração da extração: {data_atual.strftime('%m/%Y')} a {data_limite.strftime('%m/%Y')}")
    # Lista para guardar todos os dataframes
    todos_dfs = []

    # Caminho correto para o EdgeDriver.exe
    caminho_executavel = Path(__file__).resolve().parents[1]
    caminho_driver = caminho_executavel / "drivers" / "msedgedriver.exe"
    print(f'[DEBUG - Extrair Orçamento] Caminho do executável driver: {caminho_executavel}')
    
    # Inicializa o navegador Edge
    service = Service(executable_path=caminho_driver)
    edge_options = Options()
    edge_options.add_argument("--disable-popup-blocking")
    edge_options.add_argument("--start-maximized")
    edge_options.set_capability("ms:loggingPrefs", {"browser": "ALL"})

    driver = webdriver.Edge(service=service, options=edge_options)
    driver.implicitly_wait(2)
    driver.maximize_window()
    wait = WebDriverWait(driver, 50)

    try:
        # === Login e navegação inicial ===
        data_login = f"{mes_inicio:02d}/{ano_inicio}"
        ultimo_dia = calendar.monthrange(int(ano_fim), int(mes_fim))[1]
        data_fim = f"{mes_fim:02d}/{ano_fim}"

        print(f"Filtrando entre {data_login} e {data_fim}...")

        # Abre a URL de login
        url = 'https://financeiro.postalis.org.br/ControleAcesso/login/Login.aspx'
        driver.get(url)

        # Preenche os campos de login
        usuario_login = usuario
        senha_login = senha

        wait.until(EC.presence_of_element_located((By.NAME, "ctl00$MainContent$lgnLogin$UserName"))).send_keys(usuario_login)
        driver.find_element(By.NAME, "ctl00$MainContent$lgnLogin$Password").send_keys(senha_login)
        driver.find_element(By.NAME, "ctl00$MainContent$lgnLogin$LoginButton").click()
        print("[DEBUG] Login enviado.")

        time.sleep(2)
        print("[DEBUG] URL atual após login:", driver.current_url)

        try:
            botao = wait.until(EC.presence_of_element_located((By.NAME, "ctl00$MainContent$imbOrcamento")))
            print("[DEBUG] Botão 'Orçamento' encontrado.")
            driver.execute_script("arguments[0].scrollIntoView(true);", botao)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", botao)
            #2ª tentativa de acesso
            botao = wait.until(EC.presence_of_element_located((By.NAME, "ctl00$MainContent$imbOrcamento")))
            print("[DEBUG] Botão 'Orçamento' encontrado.")
            driver.execute_script("arguments[0].scrollIntoView(true);", botao)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", botao)
            print("[DEBUG] Clique executado com sucesso.")
        except Exception as e:
            print("[ERRO] Não encontrou ou não conseguiu clicar no botão 'Orçamento':", e)
            driver.quit()
            sys.exit()

        try:
            menu_relatorios = wait.until(EC.presence_of_element_located((By.ID, "MenuContent_tvwMenut55")))
            print("[DEBUG] Menu 'Relatórios' encontrado.")
            driver.execute_script("arguments[0].scrollIntoView(true);", menu_relatorios)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", menu_relatorios)
            print("[DEBUG] Menu 'Relatórios' clicado com sucesso.")
        except Exception as e:
            print("[ERRO] Não conseguiu encontrar ou clicar no menu 'Relatórios':", e)
            driver.quit()
            sys.exit()

        try:
            menu_acompanhamento = wait.until(EC.presence_of_element_located((By.ID, "MenuContent_tvwMenut65")))
            driver.execute_script("arguments[0].click();", menu_acompanhamento)
            print("[DEBUG] Submenu 'Movimento' clicado.")
        except Exception as e:
            print("[ERRO] Falha ao clicar em 'Movimento':", e)
            driver.quit()
            sys.exit()

        # === Laço para processar cada mês ===
        while data_atual <= data_limite:
            try:
                data_inicio = f"{data_atual.month:02d}/{data_atual.year}"
                
                # Esperar e clicar em Análise Comparativa
                try:
                    time.sleep(1)
                    menu_analise_comparativa = wait.until(EC.presence_of_element_located((By.ID, "MenuContent_tvwMenut67")))
                    driver.execute_script("arguments[0].click();", menu_analise_comparativa)
                    print("[DEBUG] Submenu 'Realizado' clicado.")
                except Exception as e:
                    print("[ERRO] Falha ao clicar em 'Realizado':", e)
                    continue
                
                time.sleep(2)
                # Preencher o valor no campo e disparar o evento de "blur"
                campo_data = driver.find_element(By.ID, "MainContent_MainContent_dbData")
                driver.execute_script("arguments[0].value = arguments[1];", campo_data, data_inicio)
                driver.execute_script("arguments[0].dispatchEvent(new Event('blur'));", campo_data)
                print(f"[DEBUG] Data inserida com sucesso: {data_inicio}")

                time.sleep(5)

                # Aguarda o combo estar disponível para escolher o orçamento
                select_element = wait.until(EC.presence_of_element_located((By.ID, "MainContent_MainContent_ddlVersao")))
                select = Select(select_element)
                opcoes = [(int(opt.get_attribute("value")), opt.text) for opt in select.options if opt.get_attribute("value").isdigit()]
                maior_valor = max(opcoes, key=lambda x: x[0])[0]
                select.select_by_value(str(maior_valor))
                print(f"[DEBUG] Tipo de orçamento escolhido: {maior_valor}")

                time.sleep(5)
                driver.execute_script("""
                    const ddl = document.getElementById('MainContent_MainContent_ddlBalancete');
                    ddl.value = '3';
                    ddl.dispatchEvent(new Event('change', { bubbles: true }));
                """)
                print("[DEBUG] Plano 03 - Plano PostalPrev escolhido.")

                time.sleep(2)

                # Selecionar opção "Contábil"
                radio_contabil = wait.until(EC.presence_of_element_located((By.ID, "MainContent_MainContent_rblTipoRealizacao_0")))
                radio_contabil.click()
                print("[DEBUG] Opção 'Contábil' selecionada.")

                botao_imprimir = wait.until(EC.element_to_be_clickable((By.ID, "MainContent_MainContent_btnImprimir")))
                botao_imprimir.click()
                print("[DEBUG] Botão 'Imprimir' clicado com sucesso.")

                # Captura janelas existentes antes do clique
                handles_antes = driver.window_handles
                print(f"[DEBUG] Janelas antes do clique: {handles_antes}")

                # Espera nova janela pop-up
                try:
                    wait.until(lambda d: len(d.window_handles) > len(handles_antes))
                    nova_janela = [h for h in driver.window_handles if h not in handles_antes][0]
                    driver.switch_to.window(nova_janela)
                    print(f"[DEBUG] Foco alterado para nova janela: {driver.current_url}")
                except Exception as e:
                    print(f"[ERRO] Não foi possível detectar ou trocar para nova janela: {e}")
                    continue

                time.sleep(3)

                # Clica na aba "Análise Comparativa Por Centro"
                try:
                    aba = wait.until(EC.element_to_be_clickable((By.ID, "__tab_tabReportSelector_ctl00")))
                    aba.click()
                    print("[DEBUG] Aba 'Análise Comparativa por Centro' clicada com sucesso.")
                except TimeoutException:
                    print("[ERRO] Aba 'Análise Comparativa por Centro' não encontrada.")
                    continue

                # Clica no campo de formato
                try:
                    campo_formato = wait.until(EC.element_to_be_clickable((By.ID, "rtbControles_Menu_ITCNT11_SaveFormat_I")))
                    campo_formato.click()
                    print("[DEBUG] Campo de formato clicado.")
                except TimeoutException:
                    print("[ERRO] Campo de formato não encontrado.")
                    continue

                # Seleciona "Excel sem formatação"
                try:
                    input_format = wait.until(EC.element_to_be_clickable((By.ID, "rtbControles_Menu_ITCNT11_SaveFormat_I")))
                    input_format.click()
                    input_format.send_keys(Keys.ARROW_DOWN)
                    input_format.send_keys(Keys.ENTER)
                    print("[DEBUG] Opção 'Excel sem formatação (*.xlsx)' selecionada via teclado.")
                except TimeoutException:
                    print("[ERRO] Botão de exportação não encontrado.")
                    continue

                try:
                    time.sleep(1)
                    botao_salvar_td = wait.until(EC.element_to_be_clickable((By.ID, "rtbControles_Menu_DXI12_I")))
                    driver.execute_script("arguments[0].click();", botao_salvar_td)
                    print("[DEBUG] Clique no <td> do botão salvar realizado.")
                except Exception as e:
                    print(f"[ERRO] Falha ao clicar no botão salvar (td): {e}")
                    continue

                time.sleep(5)

                # Processamento do arquivo Excel
                usuario_caminho = os.getlogin()
                pasta_downloads = os.path.join("C:\\Users", usuario_caminho, "Downloads")
                padrao_arquivo = os.path.join(pasta_downloads, "Análise Comparativa*.xlsx")
                arquivos_encontrados = glob.glob(padrao_arquivo)
                arquivos_encontrados.sort(key=os.path.getmtime, reverse=True)
                caminho_entrada = arquivos_encontrados[0] if arquivos_encontrados else None

                if not caminho_entrada:
                    print("[ERRO] Nenhum arquivo encontrado.")
                    continue

                caminho_saida = os.path.join(pasta_downloads, "Análise Comparativa PostalPrev - Limpo.xlsx")

                excel = win32.gencache.EnsureDispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(caminho_entrada)
                ws = wb.Sheets(1)
                ws.Cells.Select()
                excel.Selection.UnMerge()
                wb.SaveAs(caminho_saida)
                wb.Close()
                excel.Quit()

                print(f"[INFO] Mesclagens desfeitas e planilha salva em: {caminho_saida}")

                df = pd.read_excel(caminho_saida, header=5)
                colunas_para_excluir = [9, 11]
                df.drop(df.columns[colunas_para_excluir], axis=1, inplace=True)

                # Renomear colunas de forma padronizada
                df.columns = [
                    "Conta",
                    "Orcado_Anual",
                    "Composicao_Anual_Percentual",
                    "Orcado_Mes",
                    "Realizado_Mes",
                    "Desvio_Mes_Percentual",
                    "Orcado_Acumulado",
                    "Realizado_Acumulado",
                    "Composicao_Executado_Percentual",
                    "Desvio_Acumulado_Percentual",
                    "Saldo_Disponivel",
                    "Saldo_Disponivel_Percentual"
                ]

                # Adiciona a data referente ao mês do relatório
                df["Data_Referencia"] = data_atual

                # Encontra a última linha válida com valor na coluna "Orcado_Mes"
                ultima_linha_valida = df[~df["Orcado_Mes"].isna()].last_valid_index()
                df = df.loc[:ultima_linha_valida]

                # Adiciona à lista geral de DataFrames
                todos_dfs.append(df)

            except Exception as e:
                print(f"[ERRO] Falha ao processar {data_atual.strftime('%b/%Y')}: {e}")
                continue

            try:
            # Fecha a janela pop-up, mas mantém o navegador principal aberto
                if len(driver.window_handles) > 1:
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    print("[DEBUG] Janela pop-up fechada, retornando à janela principal.")
            
            except Exception as e:
                print(f"[DEBUG] Erro ao tentar fechar janela pop-up: {e}")
            
            print(f"[DEBUG] Avançando para o próximo mês após {data_atual.strftime('%m/%Y')}")
            data_atual += relativedelta(months=1)
                           
    finally:
        driver.quit()

    # Consolidar e salvar o resultado final
    df_final = pd.concat(todos_dfs, ignore_index=True)
    caminho_saida_final = os.path.join(pasta_downloads, "Analise Comparativa PostalPrev - Consolidado.xlsx")
    df_final.to_excel(caminho_saida_final, index=False)

    print(f"[SUCESSO] Arquivo consolidado salvo em: {caminho_saida_final}")
    messagebox.showinfo("Sucesso", f"Arquivo consolidado salvo com sucesso em:\n{caminho_saida_final}")

