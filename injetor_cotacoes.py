import pdfplumber
import re
import os
import logging
import win32com.client as win32

# ==========================================
# 1. CONFIGURAÇÕES (DADOS ANONIMIZADOS)
# ==========================================
ARQUIVO_PDF = "cotacao_fornecedor.pdf"
ARQUIVO_EXCEL = "planilha_base.xls"
ARQUIVO_SAIDA_XLS = "planilha_preenchida.xls"
ARQUIVO_RELATORIO = "relatorio_divergencias.txt"

LINHA_INICIAL_EXCEL = 10
COLUNAS_EXCEL = {
    "material": 1, # Coluna A
    "qtd": 3,      # Coluna E
    "preco": 5,    # Coluna G
    "ipi": 6,      # Coluna H
    "icms": 9      # Coluna K
}

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def safe_float(valor_texto):
    if not valor_texto or str(valor_texto).strip() == "":
        return 0.0
    try:
        return float(str(valor_texto).strip().replace('.', '').replace(',', '.'))
    except ValueError:
        return 0.0

# ==========================================
# 2. LEITURA DO PDF
# ==========================================
dados_cotacao = {}
logging.info("Iniciando varredura do PDF...")

with pdfplumber.open(ARQUIVO_PDF) as pdf:
    for page in pdf.pages:
        tabela = page.extract_table()
        if not tabela: continue
        
        for linha in tabela:
            if not linha or len(linha) < 10 or linha[3] == "Referência/Descrição" or linha[3] is None:
                continue
            
            match = re.search(r'^(\d\.\d{9})', str(linha[3]))
            if match:
                codigo_material = match.group(1)
                dados_cotacao[codigo_material] = {
                    "qtd": safe_float(linha[4]),
                    "valor_unit": safe_float(linha[5]),
                    "ipi": safe_float(linha[8]),
                    "icms": safe_float(linha[9])
                }

# ==========================================
# 3. MOTOR NATIVO DO EXCEL
# ==========================================
logging.info("Abrindo o motor nativo do Excel para injetar dados...")

caminho_entrada = os.path.abspath(ARQUIVO_EXCEL)
caminho_saida = os.path.abspath(ARQUIVO_SAIDA_XLS)

if os.path.exists(caminho_saida):
    os.remove(caminho_saida)

try:
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = excel.Workbooks.Open(caminho_entrada)
    ws = wb.Sheets(1)
    
    ult_linha = ws.Cells(ws.Rows.Count, COLUNAS_EXCEL["material"]).End(-4162).Row
    itens_atualizados = 0

    for row in range(LINHA_INICIAL_EXCEL, ult_linha + 1):
        celula_material = ws.Cells(row, COLUNAS_EXCEL["material"]).Value
        
        if celula_material:
            codigo = str(celula_material).strip()
            if codigo in dados_cotacao:
                dados = dados_cotacao[codigo]
                
                # Injeção dinâmica usando o dicionário
                ws.Cells(row, COLUNAS_EXCEL["qtd"]).Value = dados["qtd"]
                ws.Cells(row, COLUNAS_EXCEL["preco"]).Value = dados["valor_unit"]
                ws.Cells(row, COLUNAS_EXCEL["ipi"]).Value = dados["ipi"]
                ws.Cells(row, COLUNAS_EXCEL["icms"]).Value = dados["icms"]
                
                itens_atualizados += 1

    wb.SaveAs(caminho_saida, FileFormat=56)
    wb.Close()
    excel.Application.Quit()
    logging.info(f"Sucesso! {itens_atualizados} linhas injetadas.")

except Exception as e:
    logging.error(f"Erro crítico no motor Excel: {e}")
    if 'excel' in locals(): excel.Application.Quit()
