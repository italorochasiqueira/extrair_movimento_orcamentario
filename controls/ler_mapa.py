#Rotina de leitura de arquivos de Mapa

from pathlib import Path
import pdfplumber
import re
import pandas as pd

def ler_mapa_comparativo():

    caminho_usuario = Path.cwd().resolve().parent.parent
    print(f"[DEBUG] Base do diretório acessada: {caminho_usuario}")

    caminho_base = caminho_usuario / "FS_GCO_COR - Documentos" / "GESTÃO ORÇAMENTÁRIA" / "ORÇAMENTO 2025" / "ACOMPANHAMENTO ORÇAMENTÁRIO" / "Contratações e Renovações Contratuais"

    caminho_teste = r'C:\Users\italo.siqueira\Postalis\FS_GCO_COR - Documentos\GESTÃO ORÇAMENTÁRIA\ORÇAMENTO 2025\ACOMPANHAMENTO ORÇAMENTÁRIO\Contratações e Renovações Contratuais\AIN\6 - Junho\11_06_2025 - Mapa Comparativo - Auditoria Contabil_GCO - (chamado 107294).pdf'

    with pdfplumber.open(caminho_teste) as pdf:
        texto = ""
        for pagina in pdf.pages:
            texto += pagina.extract_text() + "\n"

    
    # 1️⃣ Extrair campos principais com regex mais precisa
    orcamento = re.search(r"Orçamento estimado:\s*R\$ ?([\d\.,]+)", texto)
    conta = re.search(r"Conta orçamentária:\s*([^\n]*?)(?=\s*Natureza Financeira)", texto)
    natureza = re.search(r"Natureza Financeira:\s*([^\n]*?)(?=\s*Plano:)", texto)

    # 2️⃣ Identificar qual plano está marcado (☒)
    plano_match = re.search(r"Plano:\s*(?:☒\s*(PGA|BD|PostalPrev)|PGA|BD|PostalPrev)", texto)
    plano = plano_match.group(1) if plano_match else None

    # 🔹 Normalizar e dividir em linhas
    linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]
    print("\n[DEBUG] Linhas detectadas:")
    for i, linha in enumerate(linhas):
        print(f"{i:02d}: {repr(linha)}")

    # 🔹 Inicializar variáveis
    valor_periodo = valor_corrente = valor_proximo = None

    for i, linha in enumerate(linhas):
        if "Para o período contratual" in linha:
            # A linha seguinte contém os três valores
            if i + 1 < len(linhas):
                valores = re.findall(r"R\$ ?([\d\.,]+)", linhas[i + 1])
                if len(valores) >= 3:
                    valor_periodo, valor_corrente, valor_proximo = valores[:3]
                elif len(valores) == 2:
                    valor_periodo, valor_corrente = valores
                elif len(valores) == 1:
                    valor_periodo = valores[0]
            break  # Não precisa continuar o loop

    print("DEBUG:", valor_periodo, valor_corrente, valor_proximo)

    data_aprovacao = re.search(r"Data da análise da COR\s*(\d{2}/\d{2}/\d{4})", texto)
    if data_aprovacao:
        data_aprovacao = data_aprovacao.group(1)

    # Montar DataFrame
    dados = {
        "Orçamento Estimado": [orcamento.group(1) if orcamento else None],
        "Conta Orçamentária": [conta.group(1).strip() if conta else None],
        "Natureza Financeira": [natureza.group(1).strip() if natureza else None],
        "Plano": [plano],
        "Valor Período Contratual": [valor_periodo],
        "Valor Exercício Corrente": [valor_corrente],
        "Valor Próximo Exercício": [valor_proximo],
        "Data aprovação": [data_aprovacao],
    }

    df = pd.DataFrame(dados)

    # Exportar
    df.to_excel("dados_extraidos.xlsx", index=False)

    print(df)
    print("\n✅ Dados exportados com sucesso para 'dados_extraidos.xlsx'")

mapa = ler_mapa_comparativo()
