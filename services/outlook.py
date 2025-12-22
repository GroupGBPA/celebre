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
from psycopg2 import errors
import pythoncom

# --- CONSTANTES ---
OUTLOOK_PATHS = [
    r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
]

# Query de criação da tabela (com IF NOT EXISTS para evitar erro se já existir)
CREATE_TABLE_QUERY = """
    CREATE TABLE IF NOT EXISTS broker_emails (
        id BIGSERIAL       PRIMARY KEY,
        email_subject      TEXT NOT NULL,
        sender_email       TEXT NOT NULL,
        received_time      TIMESTAMP NOT NULL,
        filename           TEXT NOT NULL,
        file_data          BYTEA NOT NULL,
        created_at         TIMESTAMP NOT NULL DEFAULT NOW()
    );
"""

# --- FUNÇÕES AUXILIARES ---

def outlook_is_open():
    try:
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and proc.info['name'].lower() == 'outlook.exe':
                return True
    except Exception as e:
        logger.warning(f"Não foi possível verificar processos: {e}")
    return False

def open_outlook():
    logger.info("Tentando abrir o executável do Outlook...")
    for path in OUTLOOK_PATHS:
        if os.path.exists(path):
            try:
                subprocess.Popen([path])
                logger.info(f"Executando uma nova sessão do outlook via: {path}")
                return True
            except Exception as e:
                logger.error(f"Erro ao tentar abrir o executável: {e}")
    
    logger.critical("Outlook não se encontra em nenhum dos caminhos configurados.")
    raise FileNotFoundError("Outlook não se encontra em nenhum dos caminhos")

def verify_db_structure():
    """Verifica e cria a tabela broker_emails se necessário."""
    logger.info("Verificando estrutura do banco de dados...")
    conn = None
    try:
        conn = db_conection()
        cursor = conn.cursor()
        cursor.execute(CREATE_TABLE_QUERY)
        conn.commit()
        logger.info("Tabela 'broker_emails' verificada/criada com sucesso.")
        cursor.close()
    except Exception as e:
        logger.critical(f"Erro fatal ao tentar criar a tabela no banco: {e}")
        if conn:
            conn.rollback()
        raise e # Levanta o erro para parar o processo, pois sem tabela não podemos seguir
    finally:
        if conn:
            conn.close()

# --- FUNÇÃO PRINCIPAL ---

def outlook_process():
    logger.info("=== Iniciando Processo de Extração de Emails ===")
    
    # 1. Garantir estrutura do Banco (Cria a tabela se não existir)
    try:
        verify_db_structure()
    except Exception:
        return # Para tudo se não conseguir criar a tabela

    # 2. Garantir Abertura do Outlook
    try:
        if not outlook_is_open():
            logger.info('Outlook encontra-se fechado. Iniciando abertura...')
            open_outlook()
            logger.info("Aguardando 25 segundos para carregamento do Outlook...")
            time.sleep(25)
        else:
            logger.info('Outlook já está aberto.')
            
        # Tenta conectar ao objeto COM
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
        messages = inbox.Items
        logger.info('Conexão com Outlook MAPI estabelecida com sucesso.')

    except Exception as e:
        logger.critical(f"Falha fatal ao conectar com a aplicação Outlook: {e}")
        return

    # 3. Configuração de Pastas
    try:
        project_path = Path(__file__).parent.parent
        download_folder = os.path.join(project_path, 'tmp')
        os.makedirs(download_folder, exist_ok=True)
        logger.info(f"Pasta de downloads configurada: {download_folder}")
    except Exception as e:
        logger.critical(f"Erro ao criar pasta temporária: {e}")
        return

    # 4. Filtragem de Mensagens
    try:
        filtered_messages = messages.Restrict("[Unread] = true")
        non_read_count = filtered_messages.Count
        
        logger.info(f"Total de emails não lidos encontrados: {non_read_count}")

        if non_read_count == 0:
            logger.info("Nenhum e-mail não lido encontrado. Encerrando o processo.")
            sys.exit(0)
            
        filtered_messages.Sort("[ReceivedTime]", True)

    except Exception as e:
        logger.error(f"Erro ao filtrar mensagens: {e}")
        return

    # Array dos attachments
    pdf_attachments = []
    
    logger.info("Iniciando iteração sobre as mensagens...")

    # 5. Iteração e Processamento
    for msg in filtered_messages:
        subject = "Desconhecido"
        try:
            subject = getattr(msg, 'Subject', 'Sem Assunto')
            
            if not hasattr(msg, 'SenderEmailAddress'):
                logger.warning(f"Item '{subject}' ignorado pois não possui remetente.")
                continue
                
            sender_email = msg.SenderEmailAddress
            total_attachments = msg.Attachments.Count
            
            logger.info(f"Processando email: '{subject}' | Anexos: {total_attachments}")
            
            email_processed_successfully = False

            # Loop nos anexos
            for i in range(1, total_attachments + 1):
                try:
                    attachment = msg.Attachments.Item(i)
                    filename = attachment.FileName or "undefined"

                    if not filename.lower().endswith(".pdf"):
                        continue

                    logger.info(f"  -> PDF Encontrado: {filename}")

                    file_path = os.path.join(download_folder, filename)
                    attachment.SaveAsFile(file_path)
                    
                    file_content = None
                    try:
                        with open(file_path, "rb") as f:
                            file_content = f.read()
                    except Exception as e_read:
                        logger.error(f"Erro ao ler binário do arquivo {filename}: {e_read}")
                        continue 

                    pdf_attachments.append({
                        "email_subject": subject,
                        "sender": sender_email,
                        "received_time": msg.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
                        "filename": filename,
                        "file_path": file_path,
                        "content_bytes": file_content
                    })
                    
                    email_processed_successfully = True

                except Exception as attach_err:
                    logger.error(f"Erro ao processar anexo {i} do email '{subject}': {attach_err}")

            if email_processed_successfully:
                try:
                    msg.Unread = False
                    msg.Save()
                    logger.info(f"Email '{subject}' marcado como lido.")
                except Exception as e_mark:
                    logger.error(f"Erro ao marcar email '{subject}' como lido: {e_mark}")
            else:
                logger.info(f"Email '{subject}' mantido como não lido (sem PDF ou erro).")

        except Exception as e:
            logger.error(f"Erro genérico ao processar email '{subject}': {e}")
            continue

    # 6. Inserção no Banco de Dados
    if not pdf_attachments:
        logger.info("Nenhum anexo PDF para gravar no banco.")
        return

    logger.info(f"Iniciando gravação de {len(pdf_attachments)} anexos no banco...")
    
    conn = None
    try:
        conn = db_conection()
        cursor = conn.cursor()
        
        insert_query = """
            INSERT INTO broker_emails (
                email_subject,
                sender_email,
                received_time,
                filename,
                file_data
            ) VALUES (%s, %s, %s, %s, %s)
        """
        
        items_saved = 0
        for pdf in pdf_attachments:
            try:
                cursor.execute(
                    insert_query,
                    (
                        pdf["email_subject"],
                        pdf["sender"],
                        pdf["received_time"],
                        pdf["filename"],
                        Binary(pdf["content_bytes"])
                    )
                )
                items_saved += 1
            except Exception as e_sql:
                logger.error(f"Erro ao inserir '{pdf['filename']}': {e_sql}")

        conn.commit()
        logger.info(f"Sucesso: {items_saved} registros salvos.")
        cursor.close()

    except Exception as e:
        logger.exception("Erro crítico na conexão com o banco.")
        if conn:
            conn.rollback()
    finally:
        if conn:
            conn.close()
            
    logger.info("=== Processo Finalizado ===")