from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import messagebox
from pathlib import Path
import time
import os
import pandas as pd
import re

def extrair_balancete_planos(usuario: str, senha: str, data_ini: str, data_fim: str):
    #Variáveis
    base_arquivos = []
    meses_num = {
        'JAN': '01',
        'FEV': '02',
        'MAR': '03',
        'ABR': '04',
        'MAI': '05',
        'JUN': '06',
        'JUL': '07',
        'AGO': '08',
        'SET': '09',
        'OUT': '10',
        'NOV': '11',
        'DEZ': '12'
    }


    periodos = pd.date_range(start=data_ini, end=data_fim, freq="MS")

    # Preenche os campos de login
    usuario_login = usuario
    senha_login = senha

    xpath_planos = {
        1: '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[1]',  # BD
        3: '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[3]',  # Postal Prev
        4: '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[4]',  # Balancete Auxiliar
        5: '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[5]'  # PGA

    }

    # Caminho correto para o EdgeDriver.exe
    caminho_executavel = Path(__file__).resolve().parents[1]
    caminho_driver = caminho_executavel / "drivers" / "msedgedriver.exe"
    
    # Inicializa o navegador Edge
    service = Service(executable_path=caminho_driver)

    driver = webdriver.Edge(service=service)
    driver.refresh()
    driver.maximize_window()
    wait = WebDriverWait(driver, 50)

    # Abre a URL de login
    url = 'https://financeiro.postalis.org.br/ControleAcesso/login/Login.aspx'
    driver.get(url)


    wait.until(EC.presence_of_element_located((By.NAME, "ctl00$MainContent$lgnLogin$UserName"))).send_keys(usuario_login)
    print(f"[INFO] Usuário identificado: {usuario_login}")
    driver.find_element(By.NAME, "ctl00$MainContent$lgnLogin$Password").send_keys(senha_login)

    # Clica no botão de login
    driver.find_element(By.NAME, "ctl00$MainContent$lgnLogin$LoginButton").click()

    # Aguarda o botão de "Contabilidade" estar visível e clica
    wait.until(EC.element_to_be_clickable((By.NAME, "ctl00$MainContent$imbContabilidade"))).click()


    wait.until(EC.element_to_be_clickable((By.NAME, "ctl00$MainContent$imbContabilidade"))).click()

    # Aguarda o ícone da árvore estar presente
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="MenuContent_tvwMenut96"]')
    )).click()

    #Abrir Menu "Operacionais"
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="MenuContent_tvwMenut97"]')
    )).click()

    #Abrir menu dos Balancetes
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="MenuContent_tvwMenut99"]')
    )).click()


    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="MenuContent_tvwMenut100"]')
    )).click()

    for periodo_data in periodos:
        data_str = periodo_data.strftime("%m/%Y")
        for codigo in [1, 3, 4, 5]:
            xpath_plano = xpath_planos.get(codigo)
            if not xpath_plano:
                continue

            # Aguarde e preencha o campo de data de início
            driver.execute_script(
                "document.getElementById('MainContent_MainContent_perPeriodo_dbxInicio').value = arguments[0];",
                data_str
            )

            driver.execute_script(
                "document.getElementById('MainContent_MainContent_perPeriodo_dbxFim').value = arguments[0];",
                data_str
            )

            # Abrir o botão de seleção novamente
            wait.until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="MainContent_MainContent_fpuBalancete_lnkSelecionado"]')
            ))
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="MainContent_MainContent_fpuBalancete_lnkSelecionado"]')
            )).click()

            # Seleciona o plano
            wait.until(EC.presence_of_element_located((By.XPATH, xpath_plano)))
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath_plano))).click()

            # Clica em Adicionar
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_btnAdd"]')
            )).click()

            # Clica em Cancelar para sair da tela de seleção
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="MainContent_MainContent_fpuBalancete_btnCancelar"]')
            )).click()

            # Exporta
            exportar_btn = wait.until(EC.element_to_be_clickable(
                (By.ID, 'MainContent_MainContent_btnExportar')
            ))
            driver.execute_script("arguments[0].click();", exportar_btn)

            time.sleep(5)

            print("[DEBUG] Download finalizado e arquivo movido com sucesso!")

            wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="MenuContent_tvwMenut100"]')
            )).click()

            print("Código chegou no final")

            pasta_download = os.path.join(os.path.expanduser("~"), "Downloads")

            arquivos_xlsx = [os.path.join(pasta_download, f) for f in os.listdir(pasta_download) if f.endswith('.xlsx')]
            if not arquivos_xlsx:
                print("[ERRO] Nenhum arquivo .xlsx encontrado.")
                continue

            arquivo_mais_recente = max(arquivos_xlsx, key=os.path.getctime)

            try:
                df_raw = pd.read_excel(arquivo_mais_recente, header=None)
                linha_periodo = df_raw.iloc[1, 0]
                match_periodo = re.search(r"([A-Za-z]{3})/(\d{4})", linha_periodo)
                if match_periodo:
                    mes, ano = match_periodo.groups()
                    periodo = f"{mes.upper()}/{ano}"
                else:
                    periodo = "NÃO IDENTIFICADO"

                linha_plano = df_raw.iloc[2, 0]
                match_plano = re.search(r"Balancete\(s\): (.+)", linha_plano)
                plano = match_plano.group(1) if match_plano else "PLANO NÃO IDENTIFICADO"

                df = pd.read_excel(arquivo_mais_recente, header=4)
                
                # Função que transforma data "ABR/2025" em "01/04/2025"
                def converter_para_data_completa(data_str):
                    mes_abrev, ano = data_str.split('/')
                    mes_num = meses_num.get(mes_abrev.upper())
                    if mes_num:
                        return f'01/{mes_num}/{ano}'  # Dia fixo 01
                    else:
                        return None

                df['Período'] = periodo
                df['Período'] = df['Período'].apply(converter_para_data_completa)
                df['Período'] = pd.to_datetime(df['Período'], format='%d/%m/%Y')
                df['Plano'] = plano

                base_arquivos.append(df)

                #Limpar a coluna de descrição
                if 'Descrição' in df.columns:
                    df['Descrição'] = (
                        df['Descrição']
                        .astype(str)
                        .str.replace("¦", "", regex=False)
                        .str.strip()
                        .str.replace(r"\s+", " ", regex=True)
                    )

                def classificar_plano_contas(codigo):
                    try:
                        primeiro = str(codigo).split('.')[0]
                        if primeiro == '1':
                            return "Ativo"
                        elif primeiro == '2':
                            return "Passivo"
                        elif primeiro == '3':
                            return "Gestão Previdencial"
                        elif primeiro == '4':
                            return "Plano Gestão Administrativa"
                        elif primeiro == '5':
                            return "Investimentos"
                        else:
                            return "Não Classificado"
                    except:
                        return "Não Identificado"
                    
                df['Plano de Contas'] = df['Conta Contábil'].apply(classificar_plano_contas)

                def calcular_nivel(codigo):
                    try:
                        partes = str(codigo).split('.')
                        nivel = 0
                        for parte in partes:
                            if parte != "00":
                                nivel += 1
                            else:
                                break
                        return nivel
                    except:
                        return None
                    
                df['Nível'] = df['Conta Contábil'].apply(calcular_nivel)

                def criar_hierarquia_vetorizada(df):
                    df = df.copy()
                    df_niveis = pd.DataFrame(index=df.index)
                    max_niveis = df['Nível'].max()
                    for i in range(1, max_niveis + 1):
                        df_niveis[f'Nível {i}'] = None
                    ultimos_valores = [None] * max_niveis
                    for idx, row in df.iterrows():
                        nivel_atual = row['Nível']
                        descricao = row['Descrição']
                        if nivel_atual > 0:
                            ultimos_valores[nivel_atual - 1] = descricao
                        for i in range(nivel_atual):
                            df_niveis.at[idx, f'Nível {i+1}'] = ultimos_valores[i]
                    df = pd.concat([df, df_niveis], axis=1)
                    return df

                #base_arquivos.append(df)

                print(f"[DEBUG] Arquivo processado com sucesso: {os.path.basename(arquivo_mais_recente)}")

            except Exception as e:
                print(f"[ERRO] Falha ao processar o arquivo: {e}")
            
            print(f"[INFO] Perído {data_str} extraído!")

    driver.quit()

    if base_arquivos:
        df_final = pd.concat(base_arquivos, ignore_index=True)
        df_final = criar_hierarquia_vetorizada(df_final)
        print(df_final.head())  # Exibe os primeiros registros como conferência

        # Caminho para Downloads do usuário atual
        pasta_download = os.path.join(os.path.expanduser("~"), "Downloads")
        caminho_saida = os.path.join(pasta_download, "balancetes_consolidados.xlsx")

        # Salva no Downloads
        df_final.to_excel(caminho_saida, index=False)
        print(f"[INFO] Arquivo salvo em: {caminho_saida}")

    messagebox.showinfo("Aviso","Extração concluída com sucesso!")


