# pandas -> bases de dados
# os -> arquivos do computador
# pywin32 -> enviar email
# criar arquivos .senha_email e .email_address na mesma pasta do script com a senha e email do remetente respectivamente
import os
from datetime import datetime
import pandas as pd
import win32com.client as win32
import smtplib
from email.message import EmailMessage

# Caminho absoluto para a pasta 'bases'
caminho_bases = os.path.join(os.path.dirname(__file__), "bases")
arquivos = os.listdir(caminho_bases)
print(arquivos)

tabela_consolidada = pd.DataFrame()

for nome_arquivo in arquivos:
    tabela_vendas = pd.read_csv(os.path.join(caminho_bases, nome_arquivo))
    tabela_vendas["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas["Data de Venda"],
                                                                                    unit="d")
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])

tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda")
tabela_consolidada = tabela_consolidada.reset_index(drop=True)
# Salva o arquivo Excel na mesma pasta do script
caminho_saida = os.path.join(os.path.dirname(__file__), "Vendas.xlsx")
tabela_consolidada.to_excel(caminho_saida, index=False)

# Lê a senha do arquivo oculto
with open(os.path.join(os.path.dirname(__file__), ".senha_email"), "r") as f:
    EMAIL_PASSWORD = f.read().strip()
with open(os.path.join(os.path.dirname(__file__), ".email_address"), "r") as f:
    EMAIL_ADDRESS = f.read().strip()

DESTINATARIO = "p.d.x.cardim@gmail.com"
print(EMAIL_ADDRESS,  DESTINATARIO)



msg = EmailMessage()
msg['Subject'] = f"Relatório de Vendas {datetime.today().strftime('%d/%m/%Y')}"
msg['From'] = EMAIL_ADDRESS
msg['To'] = DESTINATARIO
msg.set_content(f"""
Prezados,

Segue em anexo o Relatório de Vendas de {datetime.today().strftime('%d/%m/%Y')} atualizado.
Qualquer coisa estou à disposição.
Abs,
Lira Python
""")

# Anexa o arquivo Excel
with open(caminho_saida, 'rb') as f:
    file_data = f.read()
    file_name = os.path.basename(caminho_saida)
    msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

try:
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
    print("E-mail enviado com sucesso via Gmail!")
except Exception as e:
    print("Não foi possível enviar o e-mail via Gmail.")
    print("Erro:", e)