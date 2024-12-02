import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import time

# Caminho do arquivo Excel - Recomendo colocar em uma pasta padrao
file_path = '/home/luizsiqueira/Documentos/disparo/bolsistas_2024.xlsx'

# Configurações de e-mail
smtp_server = 'smtp.gmail.com'
smtp_port = 587
sender_email = 'email@gmail.com' # Colocar E-mail 
password = 'btvn kzjd nfwf rmbi'  # Substitua pela sua senha de app gerada 2fa - Senha de Aplicativo do Google

# Arquivo de log
log_file = "emails_enviados_log.txt"

try:
    # Carregar a planilha e filtrar a coluna de e-mails
    df = pd.read_excel(file_path)

    # Verifica se a coluna 'Email' existe
    if 'Email' not in df.columns:
        raise ValueError("A coluna 'Email' com os e-mails não foi encontrada.")

    # Filtrar e-mails válidos
    emails = df['Email'].dropna()  # Remover valores nulos
    emails = emails[emails.str.contains('@')]  # Apenas e-mails que contenham "@"

    # Abrir o arquivo de log para escrita
    with open(log_file, "w") as log:
        log.write("Log de envio de e-mails\n")
        log.write("=======================\n\n")

        # Conectar ao servidor SMTP
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Usar TLS para segurança
            server.login(sender_email, password)  # Login no servidor SMTP com a senha de app

            # Enviar e-mails
            for index, recipient_email in emails.items():
                # Criar a mensagem do e-mail
                message = MIMEMultipart()
                message['From'] = sender_email
                message['To'] = recipient_email
                message['Subject'] = 'Convite para o lançamento do Programa Bolsa Universidade 2025'

                # Corpo do e-mail com HTML (itens em negrito)
                body = """
<html>
  <body>
    <p>Prezados Estudantes,</p>
    <p>É com grande entusiasmo que convidamos todos vocês para o lançamento do <strong>Programa Bolsa Universidade 2025</strong>, que acontecerá no dia <strong>02.12.2024</strong>, às <strong>10:00h</strong>, no auditório da Prefeitura Municipal de Manaus, Localizado na Avenida Brasil, nº 2.971, Bairro Compensa.</p>
    <p>Este evento tem como objetivo dar início ao Programa Bolsa Universidade para 2025, um dos mais importantes programas socioeducacionais de Manaus, que busca oferecer oportunidades para os que mais precisam dar continuidade em seus estudos.</p>
    <p>Será oferecido aos participantes estudantes, <strong>horas extracurriculares (20 horas)</strong>, teremos a participação do Prefeito de Manaus e demais autoridades.</p>
    <p>As horas extracurriculares serão registradas, e todos os estudantes participantes que necessitar, receberão um certificado de participação ao final do evento. Portanto, não percam a oportunidade de prestigiar.</p>
    <p><strong>Obs.</strong> Para confirmar a presença é necessário registrar-se no link abaixo, para o recebimento das horas extracurriculares <strong>deverá chegar 30 minutos antes para assinar a lista de presença</strong>.</p>
    <p><a href="Link do forms">Clique aqui para se registrar.</a></p>
    <p>Contamos com a presença de todos vocês para fazer deste evento um sucesso.</p>
    <br>
    <p style="text-align:center; font-weight:bold;">
      Raimundo Pimentel Filho<br>
      Diretor Geral<br>
      ESPI / SEMAD<br>
      Escola de Serviço Público Municipal e Inclusão Socioeducacional<br>
      Secretaria Municipal de Administração
    </p>
  </body>
</html>
"""
                message.attach(MIMEText(body, 'html'))  # Enviar como conteúdo HTML

                try:
                    # Enviar o e-mail
                    server.send_message(message)
                    log.write(f"Sucesso: E-mail enviado para {recipient_email}\n")
                    print(f"E-mail enviado com sucesso para: {recipient_email}")
                except Exception as e:
                    log.write(f"Falha: Erro ao enviar e-mail para {recipient_email}: {e}\n")
                    print(f"Erro ao enviar para {recipient_email}: {e}")

                # Intervalo opcional entre os envios para evitar bloqueios
                time.sleep(1)  # Ajuste o tempo conforme necessário (em segundos)
                
except Exception as e:
    print(f"Erro ao processar os e-mails: {e}")
    with open(log_file, "a") as log:
        log.write(f"Erro ao processar e-mails: {e}\n")

print(f"Log de e-mails salvo em '{log_file}'.")
