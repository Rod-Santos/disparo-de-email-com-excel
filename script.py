import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

'''
Altere aqui as suas credenciais e nome
'''
seu_nome = "Teste!"
seu_email = "*"
sua_senha = "*"

# Se estiver usando algo diferente do gmail
# altere 'smtp.gmail.com' e 465 na linha abaixo
servidor = smtplib.SMTP('smtp.office365.com', 587)
servidor.ehlo()
servidor.starttls()
servidor.login(seu_email, sua_senha)

# Ler o arquivo
lista_emails = pd.read_excel(r"modelo_emails.xlsx")


# Obter todos os Nomes, Endereços de E-mail, Assuntos e Mensagens
todos_nomes = lista_emails['Nome']
todos_emails = lista_emails['Email']
todos_assuntos = lista_emails['Assunto']
todas_mensagens = lista_emails['Mensagem']
todos_anexos = lista_emails['Anexo']

# Nome do arquivo de log
arquivo_log = os.path.join(os.getcwd(), "log-envio-mail.txt")

# Data atual para verificação de nova data
data_atual = None

# Percorrer os e-mails
for idx in range(len(todos_emails)):
    # Obter as informações de cada registro
    nome = todos_nomes[idx]
    email = todos_emails[idx]
    assunto = todos_assuntos[idx]
    mensagem = todas_mensagens[idx]
    anexo_caminho = todos_anexos[idx]

    # Criar o e-mail
    msg = MIMEMultipart()
    msg['From'] = seu_email
    msg['To'] = email
    msg['Subject'] = assunto

    # Adicionar o corpo da mensagem
    msg.attach(MIMEText(mensagem, 'plain'))

    # Anexar o arquivo
    if anexo_caminho:
        attachment = open(anexo_caminho, 'rb')
        filename = os.path.basename(anexo_caminho)  # Obtém o nome do arquivo sem o diretório
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {filename}")
        msg.attach(part)

    # Enviar o e-mail
    try:
        servidor.send_message(msg)
        print('E-mail para {} enviado com sucesso!\n\n'.format(email))
    except Exception as e:
        print('E-mail para {} não pôde ser enviado :( porque {}\n\n'.format(email, str(e)))

    # Registrar o log
    now = datetime.now()
    data_atual_str = now.strftime("%d/%m/%Y")
    hora_atual_str = now.strftime("%H:%M:%S")
    if data_atual_str != data_atual:
        with open(arquivo_log, "a") as log_file:
            log_file.write("\n\n")
        data_atual = data_atual_str

    with open(arquivo_log, "a") as log_file:
        log_file.write(f"{data_atual_str} {hora_atual_str} - E-mail para {email} enviado com sucesso!\n\n")

    # Fechar o arquivo de anexo
    if anexo_caminho:
        attachment.close()

# Fechar o servidor SMTP
servidor.close()
