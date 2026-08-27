import pandas as pd
import os
import matplotlib.pyplot as plt
from fpdf import FPDF

def formatar_moeda(valor):
    """Transforma números no padrão contábil brasileiro (Ex: R$ 1.234,50)"""
    if pd.isna(valor):
        return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def gerar_relatorio_savings(arquivo_entrada, arquivo_saida):
    grafico_global = 'temp_dashboard.png'
    imagens_individuais = []
    
    try:
        print("Lendo e higienizando a base de dados...")
        df = pd.read_excel(arquivo_entrada, skiprows=4)
        df.columns = df.columns.str.strip()

        # Limpeza de dados
        for col in df.columns:
            df = df[~df[col].astype(str).str.strip().eq(str(col))]
        if 'Valor Inicial' in df.columns:
            df = df.dropna(subset=['Valor Inicial', 'Saving (R$)'], how='all')

        colunas_financeiras = ['Valor Inicial', 'Valor Final', 'Saving (R$)']

        # Processamento Analítico
        for col in colunas_financeiras:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        df['Comprador'] = df['Comprador'].fillna('Não Identificado')

        total_inicial = df['Valor Inicial'].sum()
        total_final = df['Valor Final'].sum()
        total_saving = df['Saving (R$)'].sum()
        perc_saving = (total_saving / total_inicial) * 100 if total_inicial > 0 else 0

        # Agregação de Resultados
        resumo = df.groupby('Comprador').agg({
            'Valor Inicial': 'sum',
            'Valor Final': 'sum',
            'Saving (R$)': 'sum'
        }).reset_index().sort_values(by='Saving (R$)', ascending=False)

        # Gráfico Global
        print("Renderizando painéis visuais...")
        plt.figure(figsize=(10, 5))
        bars = plt.bar(resumo['Comprador'], resumo['Saving (R$)'], color='#1F618D')
        plt.title('Ranking Global de Savings', fontsize=12, fontweight='bold')
        
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, yval, formatar_moeda(yval), 
                     ha='center', va='bottom', fontsize=9, fontweight='bold')
                     
        plt.tight_layout()
        plt.savefig(grafico_global)
        plt.close()

        # Construção do PDF Executivo
        print("Montando relatório PDF...")
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(190, 10, 'RELATORIO EXECUTIVO DE SAVINGS', 0, 1, 'C')
        pdf.ln(5)

        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(190, 10, ' INDICADORES GLOBAIS DA EQUIPE', 1, 1, 'L', fill=True)
        pdf.set_font('Arial', '', 10)
        pdf.cell(95, 8, f'Total Inicial: {formatar_moeda(total_inicial)}', 1, 0)
        pdf.cell(95, 8, f'Total Fechado: {formatar_moeda(total_final)}', 1, 1)
        pdf.cell(95, 8, f'ECONOMIA (SAVING): {formatar_moeda(total_saving)}', 1, 0)
        pdf.cell(95, 8, f'EFICIENCIA MEDIA: {perc_saving:.2f}%', 1, 1)
        pdf.ln(8)
        
        
        pdf.image(grafico_global, x=10, y=None, w=190)
        pdf.output(arquivo_saida)
        
        # Limpeza de arquivos temporários
        if os.path.exists(grafico_global): os.remove(grafico_global)

        print(f"[SUCESSO] Relatório '{arquivo_saida}' gerado com sucesso!")

    except Exception as e:
        print(f"[ERRO DO SISTEMA] Falha no processamento: {e}")

if __name__ == "__main__":
    ARQUIVO_ENTRADA = 'base_dados_compras.xlsx'
    ARQUIVO_SAIDA = 'Relatorio_Fechamento.pdf'
    gerar_relatorio_savings(ARQUIVO_ENTRADA, ARQUIVO_SAIDA)
