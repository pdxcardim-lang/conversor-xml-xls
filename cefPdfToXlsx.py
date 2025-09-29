import pdfplumber
import pandas as pd
import os
from datetime import datetime
import re

# ============================
# Caminhos de entrada/saída
# ============================
BASE_DIR = r"C:\Users\jac\Documents\cursos python\conversos xml para xls\bases"
PDF_FILE = os.path.join(BASE_DIR, "cef.pdf")
XLSX_FILE = os.path.join(BASE_DIR, "cef.xlsx")

# ============================
# Função de limpeza de valores
# ============================
def parse_valor(valor_str: str) -> float:
    """
    Converte string de valor (brasileiro) e identifica C/D.
    Exemplo:
        "1.234,56 C" -> +1234.56
        "345,67 D"   -> -345.67
    """
    if not valor_str:
        return 0.0

    valor_str = valor_str.strip()
    tipo = valor_str[-1]  # último caractere (C ou D)
    numero = valor_str[:-1].strip()

    numero = numero.replace(".", "").replace(",", ".")
    try:
        valor = float(numero)
    except ValueError:
        valor = 0.0

    if tipo.upper() == "D":
        return -valor
    return valor

# ============================
# Função fallback parse_text
# ============================
def parse_texto(pdf_file):
    dados = []
    padrao = re.compile(r"(\d{2}/\d{2}/\d{4})\s+(\S+)?\s+(.*?)\s+([\d\.,]+ [CD])\s+([\d\.,]+ [CD])")
    # Formato esperado: Data | NrDoc | Historico | Valor | Saldo

    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue

            for linha in texto.split("\n"):
                m = padrao.match(linha)
                if m:
                    data_mov = m.group(1)
                    nr_doc = m.group(2) or ""
                    historico = m.group(3).strip()
                    valor = parse_valor(m.group(4))
                    saldo = parse_valor(m.group(5))
                    dados.append([data_mov, nr_doc, historico, valor, saldo])

    return dados

# ============================
# Tenta extrair tabela
# ============================
dados = []
try:
    with pdfplumber.open(PDF_FILE) as pdf:
        for pagina in pdf.pages:
            tabela = pagina.extract_table()
            if tabela:
                for linha in tabela[1:]:  # ignora cabeçalho
                    if len(linha) >= 5:
                        data_mov, nr_doc, historico, valor_str, saldo_str = linha[:5]
                        valor = parse_valor(valor_str)
                        saldo = parse_valor(saldo_str)
                        dados.append([data_mov, nr_doc, historico, valor, saldo])
except Exception as e:
    print(f"Erro na extração via tabela: {e}")

# Se não encontrou nada, usar fallback de texto
if not dados:
    print("⚠️ Nenhuma tabela encontrada. Tentando fallback via texto...")
    dados = parse_texto(PDF_FILE)

if not dados:
    raise ValueError("Nenhum lançamento foi encontrado no PDF (nem via tabela, nem via texto).")

# ============================
# Montagem DataFrame final
# ============================
linhas_saida = []
saldo_acumulado = 0.0
data_processo = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

for data_mov, nr_doc, historico, valor, saldo in dados:
    tipo = "Entrada" if valor > 0 else "Saída"
    saldo_acumulado += valor

    linhas_saida.append([
        data_mov,          # Data
        tipo,              # Tipo (Entrada/Saída)
        historico,         # Descrição
        valor,             # Valor (R$)
        saldo_acumulado,   # Saldo acumulado (R$)
        "CEF",             # Banco ID fixo
        data_processo,     # Processado em
        "PDF-IMPORT",      # Execução (identificação especial)
        "CREDIT" if valor > 0 else "DEBIT",  # TRNTYPE
        f"{nr_doc}"        # MEMO (usando nr_doc como complemento)
    ])

df_saida = pd.DataFrame(linhas_saida, columns=[
    "Data", "Tipo (Entrada/Saída)", "Descrição",
    "Valor (R$)", "Saldo acumulado (R$)",
    "Banco ID", "Processado em", "Execução",
    "TRNTYPE", "MEMO"
])

df_saida.to_excel(XLSX_FILE, index=False)

print(f"✅ Extrato convertido com sucesso! Arquivo salvo em: {XLSX_FILE}")
