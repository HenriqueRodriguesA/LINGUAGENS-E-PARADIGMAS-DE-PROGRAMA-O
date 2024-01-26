import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Função para extrair dados da planilha
def extrair_dados(planilha_path):
    df = pd.read_excel(planilha_path)
    # Adicione aqui a lógica para extrair os dados necessários do DataFrame (df)
    # Por exemplo, você pode usar df.iloc[:, 0] para obter a primeira coluna.

# Função para enviar e-mail
def enviar_email(destinatario, assunto, corpo):
    remetente_email = "henrique_rodrigues@discente.ufg.br"
    remetente_senha = "sua_senha"

    mensagem = MIMEMultipart()
    mensagem['From'] = remetente_email
    mensagem['To'] = destinatario
    mensagem['Subject'] = assunto

    mensagem.attach(MIMEText(corpo, 'plain'))

    with smtplib.SMTP('smtp.gmail.com', 587) as servidor_smtp:
        servidor_smtp.starttls()
        servidor_smtp.login(remetente_email, remetente_senha)
        servidor_smtp.send_message(mensagem)

# Caminho da planilha
planilha_path = 'caminho/para/sua/planilha.xlsx'

# Extrair dados da planilha
dados_extraidos = extrair_dados(planilha_path)

# Formatar os dados como necessário
# Aqui você pode usar os dados extraídos para construir a mensagem de e-mail

# Enviar e-mail
destinatario = 'henrique_rodrigues@discente.ufg.br'
assunto = 'Assunto do e-mail'
corpo = 'Corpo do e-mail'

enviar_email(destinatario, assunto, corpo)
