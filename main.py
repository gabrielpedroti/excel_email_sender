import pandas as pd
import smtplib
import ssl
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import os
from datetime import datetime

# Configurações iniciais
load_dotenv()

EMAIL_USER = os.getenv("EMAIL")
EMAIL_PASS = os.getenv("SENHA")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))
USE_SSL = os.getenv("USE_SSL", "false").lower() == "true"

ARQUIVO_EXCEL = os.getenv("ARQUIVO_CONTATOS")
ASSUNTO = os.getenv("ASSUNTO")
INTERVALO_ENVIO = int(os.getenv("INTERVALO_ENVIO", 0))

ARQUIVO_HTML = "email_template.html"
LOG_FILE = "envios.log"

if not all([EMAIL_USER, EMAIL_PASS, SMTP_SERVER, SMTP_PORT, ARQUIVO_EXCEL]):
    raise ValueError("Erro: variáveis obrigatórias ausentes no arquivo .env")

# Contadores
total = 0
enviados = 0
pulados = 0
erros = 0

# Função de log
def registrar_log(texto):
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(texto + "\n")

inicio = datetime.now()
registrar_log(f"\n=== Início do envio: {inicio.strftime('%d/%m/%Y %H:%M:%S')} ===")

# Leitura de dados
df = pd.read_excel(ARQUIVO_EXCEL)

with open(ARQUIVO_HTML, "r", encoding="utf-8") as file:
    html_template = file.read()

# Conexão com servidor
if USE_SSL:
    context = ssl.create_default_context()
    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context)
else:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls(context=ssl.create_default_context())

server.login(EMAIL_USER, EMAIL_PASS)

# Processamento
for index, row in df.iterrows():
    total += 1
    nome = str(row.get("nome", "")) or str(row.get("Nome", "")).strip()
    email_destino = str(row.get("email", "")) or str(row.get("Email", "")).strip()

    if not email_destino or email_destino.lower() == "nan":
        pulados += 1
        registrar_log(f"{nome} | SEM EMAIL | PULADO")
        continue

    try:
        html_personalizado = html_template.replace("{{nome}}", nome)

        msg = MIMEMultipart("alternative")
        msg["From"] = EMAIL_USER
        msg["To"] = email_destino
        msg["Subject"] = ASSUNTO
        msg.attach(MIMEText(html_personalizado, "html"))

        server.sendmail(EMAIL_USER, email_destino, msg.as_string())

        enviados += 1
        registrar_log(f"{nome} | {email_destino} | ENVIADO")

        if INTERVALO_ENVIO > 0:
            time.sleep(INTERVALO_ENVIO)

    except Exception as e:
        erros += 1
        registrar_log(f"{nome} | {email_destino} | ERRO | {str(e)}")

# Encerramento
server.quit()

fim = datetime.now()
duracao = fim - inicio

registrar_log("\n--- RESUMO FINAL ---")
registrar_log(f"Total de registros: {total}")
registrar_log(f"Enviados com sucesso: {enviados}")
registrar_log(f"Sem e-mail (pulados): {pulados}")
registrar_log(f"Erros: {erros}")
registrar_log(f"Duração: {duracao}")
registrar_log(f"Fim do envio: {fim.strftime('%d/%m/%Y %H:%M:%S')}")
registrar_log("========================\n")

print("\nEnvio concluído.")
print(f"Total: {total}")
print(f"Enviados: {enviados}")
print(f"Pulados (sem e-mail): {pulados}")
print(f"Erros: {erros}")