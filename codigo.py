# Script para consolidar arquivos CSV de vendas, gerar um Excel e enviar por e-mail via Gmail.
# Requer: pandas, openpyxl, smtplib, email.message, arquivos .senha_email e .email_address na mesma pasta do script.

import os  # Manipulação de caminhos e arquivos do sistema operacional
from datetime import datetime  # Para manipulação de datas
import pandas as pd  # Para leitura e manipulação de dados tabulares
import smtplib  # Para envio de e-mails via protocolo SMTP
from email.message import EmailMessage  # Para criar mensagens de e-mail com anexos

# 1. Define o caminho absoluto para a pasta 'bases' onde estão os arquivos CSV
caminho_bases = os.path.join(os.path.dirname(__file__), "bases")

# 2. Lista todos os arquivos presentes na pasta 'bases'
arquivos = os.listdir(caminho_bases)
print(arquivos)  # Exibe os arquivos encontrados para conferência

# 3. Cria um DataFrame vazio para consolidar os dados de todos os arquivos CSV
tabela_consolidada = pd.DataFrame()

# 4. Percorre cada arquivo CSV encontrado na pasta 'bases'
for nome_arquivo in arquivos:
    # 4.1 Lê o arquivo CSV em um DataFrame
    tabela_vendas = pd.read_csv(os.path.join(caminho_bases, nome_arquivo))
    # 4.2 Converte a coluna "Data de Venda" para o formato de data
    tabela_vendas["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(
        tabela_vendas["Data de Venda"], unit="d"
    )
    # 4.3 Adiciona os dados lidos ao DataFrame consolidado
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])

# 5. Ordena os dados consolidados pela coluna "Data de Venda"
tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda")

# 6. Reseta o índice do DataFrame consolidado para manter a sequência correta
tabela_consolidada = tabela_consolidada.reset_index(drop=True)

# 7. Define o caminho para salvar o arquivo Excel consolidado
caminho_saida = os.path.join(os.path.dirname(__file__), "Vendas.xlsx")

# 8. Salva o DataFrame consolidado em um arquivo Excel (necessário ter openpyxl instalado)
tabela_consolidada.to_excel(caminho_saida, index=False)

# 9. Lê a senha do Gmail a partir do arquivo oculto .senha_email (deve conter apenas a senha de app)
with open(os.path.join(os.path.dirname(__file__), ".senha_email"), "r") as f:
    EMAIL_PASSWORD = f.read().strip()

# 10. Lê o endereço de e-mail do remetente a partir do arquivo oculto .email_address
with open(os.path.join(os.path.dirname(__file__), ".email_address"), "r") as f:
    EMAIL_ADDRESS = f.read().strip()

# 11. Define o destinatário do e-mail
DESTINATARIO = "p.d.x.cardim@gmail.com"

# 12. Exibe remetente e destinatário para conferência
print(EMAIL_ADDRESS, DESTINATARIO)

# 13. Cria a mensagem de e-mail
msg = EmailMessage()
msg['Subject'] = f"Relatório de Vendas {datetime.today().strftime('%d/%m/%Y')}"  # Assunto do e-mail
msg['From'] = EMAIL_ADDRESS  # Remetente
msg['To'] = DESTINATARIO  # Destinatário

# 14. Define o corpo do e-mail
msg.set_content(f"""
Prezados,

Segue em anexo o Relatório de Vendas de {datetime.today().strftime('%d/%m/%Y')} atualizado.
Agradeço sua atenção, fico à disposição. Forte Abraço.
Cardim
""")

# 15. Anexa o arquivo Excel gerado ao e-mail
with open(caminho_saida, 'rb') as f:
    file_data = f.read()  # Lê o conteúdo do arquivo Excel
    file_name = os.path.basename(caminho_saida)  # Obtém o nome do arquivo
    # Adiciona o anexo à mensagem de e-mail
    msg.add_attachment(
        file_data,
        maintype='application',
        subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=file_name
    )

# 16. Tenta enviar o e-mail via servidor SMTP do Gmail
try:
    # 16.1 Conecta ao servidor SMTP do Gmail usando SSL
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)  # Faz login no servidor SMTP
        smtp.send_message(msg)  # Envia a mensagem de e-mail
    print("E-mail enviado com sucesso via Gmail! de: " + EMAIL_ADDRESS + " para: " + DESTINATARIO)
except Exception as e:
    # 16.2 Em caso de erro, exibe mensagem de erro
    print("Não foi possível enviar o e-mail via Gmail.")
    print("Erro:", e)