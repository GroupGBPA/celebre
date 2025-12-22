# Importando bibliotecas
import win32com.client
import subprocess
import psutil
import os
from utils.logger import logger
import time
from utils.database import db_conection
from psycopg2 import Binary
from pathlib import Path
import sys

# Funções
OUTLOOK_PATHS = [
    r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
]

# Verificando se o outlook está aberto antes de iniciar o processo
def outlook_is_open():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and proc.info['name'].lower() == 'outlook.exe':
            return True
    return False

# Func para abrir o outlook via process
def open_outlook():
    for path in OUTLOOK_PATHS:
        if os.path.exists(path):
            subprocess.Popen([path])
            logger.info("Executando uma nova sessão do outlook")
            return True 
    raise FileNotFoundError("Outlook não se encontra em nenhum dos caminhos")

def outlook_process():
    if not outlook_is_open():
        logger.info('Outlook encontra-se fechado...')
        open_outlook()
        time.sleep(25)
        outlook = win32com.client.Dispatch("Outlook.Application")
    else:
        outlook = win32com.client.Dispatch("Outlook.Application")
        logger.info('Outlook funcionando corretamente')

    # Cria a pasta para salvar os anexos
    project_path = Path(__file__).parent.parent
    download_folder = os.path.join(project_path, 'tmp')
    os.makedirs(download_folder, exist_ok=True)

    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
    messages = inbox.Items

    # Filtrando as mensagens por não lidas em ordem mais recente
    non_read_messages = messages.Restrict("[Unread] = true").Count

    # Caso não existam emails não lidos encerra o prcesso
    if non_read_messages == 0:
        logger.info("Nenhum e-mail não lido encontrado. Encerrando o processo.")
        sys.exit(0)

    messages = messages.Restrict("[Unread] = true")

    # Array dos attachments
    pdf_attachments = []

    for msg in messages:
        try:
            subject = msg.Subject
            sender_email = msg.SenderEmailAddress
            total_attachments = msg.Attachments.Count

            for i in range(1, total_attachments + 1):
                attachment = msg.Attachments.Item(i)
                filename = attachment.FileName or "undefined"

                if not filename.lower().endswith(".pdf"):
                    continue

                logger.info("Encontrei um PDF no email")

                file_path = os.path.join(download_folder, filename)
                attachment.SaveAsFile(file_path) # Aqui é onde o PDF é salvo na pasta tmp

                pdf_attachments.append({
                    "email_subject": subject,
                    "sender": sender_email,
                    "received_time": msg.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
                    "filename": filename,
                    "file_path": file_path
                })

                msg.Unread = False
                msg.Save()

        except Exception as e:
            logger.error(f"Erro ao processar email '{subject}': {e}")
    
    for pdf in pdf_attachments:
        with open(pdf["file_path"], "rb") as f:
            pdf["content_bytes"] = f.read()

    logger.info("Iniciando conexão com o banco de dados")
    try:
        conn = db_conection()
        cursor = conn.cursor()
        query = """
            INSERT INTO broker_emails (
                email_subject,
                sender_email,
                received_time,
                filename,
                file_data
            ) VALUES (%s, %s, %s, %s, %s)
        """
        for pdf in pdf_attachments:
            cursor.execute(
                query,
                (
                    pdf["email_subject"],
                    pdf["sender"],
                    pdf["received_time"],
                    pdf["filename"],
                    Binary(pdf["content_bytes"])
                )
            )
        conn.commit()
        cursor.close()
        conn.close()

    except Exception as e:
        if conn:
            conn.rollback()
            logger.error("Rollback realizado devido a erro no processo")
        logger.exception("Erro geral ao salvar anexos no banco de dados")



