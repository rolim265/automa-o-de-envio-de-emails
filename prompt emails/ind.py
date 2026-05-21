# =========================
# IMPORTAÇÕES
# =========================
import pandas as pd
import smtplib
import time

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# =========================
# CONFIGURAÇÃO SMTP (HOSTINGER)
# =========================
EMAIL_REMETENTE = "seu email"
SENHA = "sua senha"

SMTP_SERVER = "smtp.hostinger.com"
SMTP_PORT = 587


# =========================
# LER PLANILHA (COLUNA A SEM NOME)
# =========================
planilha = pd.read_excel("emails.xlsx")  # ou emails.ods se for LibreOffice


# remove linhas vazias (evita erro)
planilha = planilha.dropna()


# =========================
# MENSAGEM DO E-MAIL
# =========================
ASSUNTO = "Apresentação Cablagem - Serviços elétricos"

MENSAGEM = """
Olá, tudo bem?

Encontrei o contato da empresa pelo site e gostaria de apresentar rapidamente a Cablagem.

Meu nome é Bruno Lopes e sou responsável pela empresa.

Atuamos com instalações e manutenção elétrica, montagem de painéis, automação, cabeamento estruturado, CFTV e sistemas de detecção e alarme de incêndio para empresas e comércios.

Estamos à disposição para futuras demandas, adequações ou apoio técnico quando precisarem.

Buscamos manter um atendimento próximo, organizado e com foco na qualidade da execução dos serviços.

Fico à disposição.

--
Bruno Lopes Rodrigues
Cablagem - Energia, Tecnologia e Segurança

(11) 94459-1328
bruno.lopes@cablagem.com
www.cablagem.com
"""


# =========================
# CONEXÃO SMTP
# =========================
server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls()
server.login(EMAIL_REMETENTE, SENHA)


# =========================
# ENVIO DE EMAILS
# =========================
try:
    for email_destino in planilha.iloc[:, 0]:  # coluna A

        if pd.isna(email_destino):
            continue

        email_destino = str(email_destino).strip()

        # monta email
        email = MIMEMultipart()
        email["From"] = EMAIL_REMETENTE
        email["To"] = email_destino
        email["Subject"] = ASSUNTO

        email.attach(MIMEText(MENSAGEM, "plain"))

        try:
            server.sendmail(
                EMAIL_REMETENTE,
                email_destino,
                email.as_string()
            )

            print(f"✔ Enviado para: {email_destino}")

        except Exception as erro:
            print(f"❌ Erro ao enviar para {email_destino}: {erro}")

        time.sleep(5)  # anti-spam

finally:
    server.quit()
    print("🔥 Finalizado - todos os emails processados")