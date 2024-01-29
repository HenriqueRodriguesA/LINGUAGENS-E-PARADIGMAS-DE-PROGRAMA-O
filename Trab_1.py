import tkinter as tk
from tkinter import filedialog
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Função para extrair dados da planilha
def extrair_dados(planilha_path):
    df = pd.read_excel(planilha_path)
    primeira_coluna = df.iloc[:, 0]
    dados_nao_vazios = primeira_coluna.dropna().tolist()
    return dados_nao_vazios
# Adicionado a logica para extrair a primeira coluna retornando somente dados das celulas não vazia

# Função para enviar e-mail
def enviar_email(destinatario, assunto, corpo):
    remetente_email = "henrique_rodrigues@discente.ufg.br" #seu email@gmail.com
    remetente_senha = "Senha" #sua senha

    mensagem = MIMEMultipart()
    mensagem['From'] = remetente_email
    mensagem['To'] = destinatario
    mensagem['Subject'] = assunto

    mensagem.attach(MIMEText(corpo, 'plain'))

    with smtplib.SMTP('smtp.gmail.com', 587) as servidor_smtp:
        servidor_smtp.starttls()
        servidor_smtp.login(remetente_email, remetente_senha)
        servidor_smtp.send_message(mensagem)

# Configurar a interface gráfica para selecionar o arquivo
root = tk.Tk()
root.withdraw()  # Esconder a janela principal

# Abrir a janela de diálogo para selecionar a planilha
planilha_path = filedialog.askopenfilename(title="Selecione a Planilha", filetypes=[("Planilhas Excel", "*.xlsx;*.xls")])

# Verificar se há dados para enviar por e-mail
if planilha_path:
    # Extrair dados da primeira coluna não vazia da planilha
    dados_extraidos = extrair_dados(planilha_path)

    # Verificar se há dados para enviar por e-mail
    if dados_extraidos:
        # Formatar os dados como necessário
        dados_formatados = '\n'.join(map(str, dados_extraidos))

        # Enviar e-mail
        destinatario = 'adm.ti.kurujao@gmail.com'
        assunto = 'Nome dos participantes da reunião'
        corpo = f'Segua lista dos participantes da reunião:\n\n{dados_formatados}'

        enviar_email(destinatario, assunto, corpo)
    else:
        print("Não há dados para enviar por e-mail.")
else:
    print("Nenhum arquivo selecionado.")
