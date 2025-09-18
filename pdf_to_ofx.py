# ==========================================================
# SCRIPT: Ler extrato PDF e gerar arquivo OFX
# Requisitos: pip install pdfplumber
# ==========================================================

import pdfplumber
from datetime import datetime

# Arquivos
entrada_pdf = r"C:\Users\jac\Documents\cursos python\conversos xml para xls\bases\cef_Extrato.pdf.pdf"
saida_ofx = r"C:\Users\jac\Documents\cursos python\conversos xml para xls\bases\cef_Extrato_convertido.ofx"

# Função para criar uma transação OFX
def gerar_transacao(tipo, idx, valor, data, descricao):
    sinal = "" if tipo == "CREDIT" else "-"
    return f"""  <STMTTRN>
    <TRNTYPE>{tipo}</TRNTYPE>
    <DTPOSTED>{data.strftime('%Y%m%d')}</DTPOSTED>
    <TRNAMT>{sinal}{valor:.2f}</TRNAMT>
    <FITID>{tipo}{idx:04d}</FITID>
    <NAME>{descricao}</NAME>
  </STMTTRN>"""

# Lista de transações
transacoes = []

with pdfplumber.open(entrada_pdf) as pdf:
    for pagina in pdf.pages:
        tabela = pagina.extract_table()
        if not tabela:
            continue

        for linha in tabela[1:]:  # pula cabeçalho
            data_txt, descricao, valor_txt = linha

            try:
                data = datetime.strptime(data_txt, "%d/%m/%Y")
                valor = float(valor_txt.replace(".", "").replace(",", "."))
                tipo = "CREDIT" if valor > 0 else "DEBIT"
                idx = len(transacoes) + 1
                transacoes.append(gerar_transacao(tipo, idx, abs(valor), data, descricao.strip()))
            except:
                continue

# Montar OFX
ofx_content = f"""OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
  <SIGNONMSGSRSV1>
    <SONRS>
      <STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>
      <DTSERVER>{datetime.now().strftime('%Y%m%d%H%M%S')}</DTSERVER>
      <LANGUAGE>POR</LANGUAGE>
      <FI><ORG>BANCO</ORG><FID>000</FID></FI>
    </SONRS>
  </SIGNONMSGSRSV1>
  <BANKMSGSRSV1>
    <STMTTRNRS>
      <TRNUID>1</TRNUID>
      <STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>
      <STMTRS>
        <CURDEF>BRL</CURDEF>
        <BANKACCTFROM>
          <BANKID>000</BANKID>
          <ACCTID>123456</ACCTID>
          <ACCTTYPE>CHECKING</ACCTTYPE>
        </BANKACCTFROM>
        <BANKTRANLIST>
{"\n".join(transacoes)}
        </BANKTRANLIST>
      </STMTRS>
    </STMTTRNRS>
  </BANKMSGSRSV1>
</OFX>
"""

# Salvar
with open(saida_ofx, "w", encoding="utf-8") as f:
    f.write(ofx_content)

print(f"Arquivo OFX gerado: {saida_ofx}")
