import os
import re
import datetime as dt
import pdfplumber
import pandas as pd
from tqdm import tqdm

#  CONFIGURAÇÕES 
PDFS_FOLDER = r"C:\Users\mazoni\Documents\pdf"  # Pasta com PDFs
OUTPUT_XLSX = r"C:\Users\mazoni\Downloads\dados.xlsx"  # Caminho do arquivo Excel de saída

#  EXPRESSÕES REGULARES 
PATTERNS = {
    "OC": r"OC\s*0*([0-9]{3,})",                                       # Exemplo: OC 000123
    "Número da Nota": r"N[º°]\s*0*([0-9]{3,})",                         # Exemplo: Nº 000001395
    "BIP": r"BIP\s*N[º°]?\s*[:\-]?\s*0*([0-9]{3,})",                    # Exemplo: BIP Nº: 0000024265

    # Captura QUALQUER número que vem depois da palavra QUANTIDADE, com total garantia
    "Quantidade": r"QUANTIDADE[\s\S]*?([0-9]{1,10})",
}

#  FUNÇÕES 
def extrair_por_regex(texto, pattern):
    """Aplica a regex e retorna o valor encontrado (ou vazio)."""
    m = re.search(pattern, texto, re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else ""

def extrair_texto_pdf(caminho_pdf):
    """Extrai todo o texto de um PDF."""
    texto_total = []
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto_total.append(page.extract_text() or "")
    except Exception as e:
        print(f"[ERRO] {caminho_pdf}: {e}")
    return "\n".join(texto_total)

def processar_pdfs(pasta, output_excel):
    """Lê todos os PDFs da pasta e extrai os dados definidos em PATTERNS."""
    arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
    registros = []

    if not arquivos:
        print(f"[{dt.datetime.now().strftime('%H:%M:%S')}]  Nenhum PDF encontrado na pasta.")
        return

    print(f"[{dt.datetime.now().strftime('%H:%M:%S')}] Iniciando leitura de {len(arquivos)} PDFs...")

    for nome in tqdm(arquivos, desc="Extraindo dados"):
        caminho = os.path.join(pasta, nome)
        texto = extrair_texto_pdf(caminho)
        dados = {"Arquivo": nome}
        for titulo, pattern in PATTERNS.items():
            dados[titulo] = extrair_por_regex(texto, pattern)
        registros.append(dados)

    # Cria DataFrame com colunas organizadas
    colunas = ["Arquivo"] + list(PATTERNS.keys())
    df = pd.DataFrame(registros, columns=colunas)

    # Salva em Excel
    df.to_excel(output_excel, index=False)
    print(f"[{dt.datetime.now().strftime('%H:%M:%S')}]  Dados salvos em: {output_excel}")
    print(f"Total de PDFs processados: {len(df)}")

#  EXECUÇÃO
if __name__ == "__main__":
    if not os.path.isdir(PDFS_FOLDER):
        print(f"Pasta '{PDFS_FOLDER}' não encontrada. Crie e coloque os PDFs lá.")
        exit()

    print(" Iniciando extração...")
    processar_pdfs(PDFS_FOLDER, OUTPUT_XLSX)
    print(" Processo concluído.")
