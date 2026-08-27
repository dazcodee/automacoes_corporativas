import os
import shutil
import pandas as pd

def organizar_arquivos_engenharia(pasta_origem, caminho_excel, pasta_destino_base):
    print("Iniciando a leitura da planilha de referência...")
    
    try:
        
        df = pd.read_excel(caminho_excel, skiprows=8)
        lista_materiais = df['Material'].astype(str).str.strip().tolist()
        print(f"{len(lista_materiais)} itens encontrados na coluna de referência.")
    except Exception as e:
        print(f"Erro ao ler o Excel: {e}")
        return

    
    pasta_pdf = os.path.join(pasta_destino_base, "ARQUIVOS_PDF")
    pasta_dwg = os.path.join(pasta_destino_base, "ARQUIVOS_DWG")
    
    os.makedirs(pasta_pdf, exist_ok=True)
    os.makedirs(pasta_dwg, exist_ok=True)

    arquivos_pdf_copiados = 0
    arquivos_dwg_copiados = 0

    print("Varrendo a rede em busca dos desenhos técnicos...")
    
    
    for raiz, diretorios, arquivos in os.walk(pasta_origem):
       
        if "ARQUIVOS_PDF" in raiz or "ARQUIVOS_DWG" in raiz:
            continue

        for arquivo in arquivos:
            nome_sem_extensao, extensao = os.path.splitext(arquivo)
            extensao = extensao.lower()

            
            if nome_sem_extensao in lista_materiais:
                caminho_origem = os.path.join(raiz, arquivo)
                
                
                if extensao == '.pdf':
                    caminho_final = os.path.join(pasta_pdf, arquivo)
                    if not os.path.exists(caminho_final):
                        shutil.copy(caminho_origem, caminho_final)
                        arquivos_pdf_copiados += 1
                        
                elif extensao == '.dwg':
                    caminho_final = os.path.join(pasta_dwg, arquivo)
                    if not os.path.exists(caminho_final):
                        shutil.copy(caminho_origem, caminho_final)
                        arquivos_dwg_copiados += 1

    print("\n--- RESUMO DA OPERAÇÃO ---")
    print(f"Arquivos PDF roteados: {arquivos_pdf_copiados}")
    print(f"Arquivos DWG roteados: {arquivos_dwg_copiados}")

# ==========================================
# CONFIGURAÇÃO DOS CAMINHOS (ANONIMIZADOS)
# ==========================================
CAMINHO_ORIGEM = r"C:\Temp\Servidor_Engenharia\Projeto_X"
CAMINHO_PLANILHA = r"C:\Temp\Processamento\Lista_Materiais.xls"
CAMINHO_DESTINO = r"C:\Temp\Producao_Liberada"

organizar_arquivos_engenharia(CAMINHO_ORIGEM, CAMINHO_PLANILHA, CAMINHO_DESTINO)
