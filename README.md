# automacoes_corporativas
Automações que fiz com IA e python para automatizar tarefas demoradas e chatas, que desviam tempo e função

# ⚙️ Automações Corporativas com Python

Este repositório contém uma suíte de scripts em Python desenvolvidos para resolver gargalos operacionais reais e eliminar trabalho braçal em setores administrativos e de engenharia. 

O foco destas automações não é criar softwares complexos com interfaces gráficas pesadas, mas sim entregar **ferramentas de terminal rápidas e eficientes**, rodando em background (via PowerShell ou Prompt) para reduzir SLAs de horas para minutos.

> ⚠️ **Nota de Confidencialidade (NDA):** Todos os dados, diretórios de rede, nomes de clientes e variáveis financeiras presentes nestes scripts foram anonimizados (Mock Data) para proteger os segredos industriais e a infraestrutura das empresas onde os problemas originais foram resolvidos.

## 🛠️ As Soluções

### 1. Injetor de Cotações (PDF para Excel Nativo)
* **O Problema:** A equipe de compras gastava horas copiando dados tributários (IPI, ICMS) e valores de PDFs de fornecedores para dentro de planilhas do Excel.
* **A Solução:** Um script que utiliza a biblioteca `pdfplumber` para extrair os dados e o motor nativo do Windows (`win32com`) para abrir o Excel em background e **injetar os dados na planilha original**, preservando todas as fórmulas e formatações corporativas.
* **Impacto:** SLA da tarefa reduzido de 8 horas para 4 minutos. Eliminação de 100% de erros de digitação.

### 2. Roteador de Arquivos de Engenharia
* **O Problema:** Projetistas perdiam tempo valioso vasculhando dezenas de pastas de rede atrás de desenhos técnicos (.DWG) e documentos (.PDF) específicos para enviar para a linha de produção.
* **A Solução:** Um algoritmo de roteamento (`os` e `shutil`) que lê uma lista de materiais aprovados em uma planilha, varre o servidor e segrega (copia) apenas os arquivos necessários para diretórios limpos de produção.
* **Impacto:** Fim do extravio de versões antigas de projetos e padronização da entrega para o chão de fábrica.

### 3. Gerador de Relatórios Executivos (Savings)
* **O Problema:** A gestão precisava de fechamentos visuais de performance financeira da equipe, mas a consolidação dos dados era manual e demorada.
* **A Solução:** Um script analítico que utiliza o `pandas` para higienizar a base de dados (removendo "sujeiras" de digitação), consolida as economias (Savings), plota gráficos com `matplotlib` e exporta um documento profissional formatado via `fpdf`.
* **Impacto:** Geração de relatórios gerenciais padronizados e à prova de falhas humanas em questão de segundos.

---
*Desenvolvido por Guilherme Dazchen junto a gemini como parte da exploração de automações e otimização de processos de negócio.*
