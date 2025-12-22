import logging
import sys
from utils.database import db_conection

# --- QUERY DE CRIAÇÃO DA TABELA ---
CREATE_TABLE_QUERY = """
    CREATE TABLE IF NOT EXISTS process_logs (
        id SERIAL PRIMARY KEY,
        log_level VARCHAR(20),
        msg TEXT,
        module_name VARCHAR(50),
        created_at TIMESTAMP DEFAULT (NOW() AT TIME ZONE 'UTC' - INTERVAL '3 HOURS')
    );
"""

# --- FUNÇÃO PARA GARANTIR A TABELA ---
def _init_log_table():
    """Tenta criar a tabela de logs se ela não existir."""
    conn = None
    try:
        conn = db_conection()
        cursor = conn.cursor()
        cursor.execute(CREATE_TABLE_QUERY)
        conn.commit()
        # Não usamos o logger aqui para evitar recursão/loop infinito
        print("Tabela 'process_logs' verificada/criada com sucesso.")
    except Exception as e:
        print(f"ATENÇÃO: Não foi possível criar/verificar a tabela de logs: {e}")
    finally:
        if conn:
            conn.close()

# --- CLASSE HANDLER DO BANCO (OBRIGATÓRIO SER CLASSE P/ LIB LOGGING) ---
class DBHandler(logging.Handler):
    def emit(self, record):
        conn = None
        try:
            log_msg = self.format(record)
            conn = db_conection()
            cursor = conn.cursor()
            
            query = """
                INSERT INTO process_logs (log_level, msg, module_name) 
                VALUES (%s, %s, %s)
            """
            
            cursor.execute(query, (
                record.levelname, 
                log_msg, 
                record.module
            ))
            
            conn.commit()
            cursor.close()
            
        except Exception:
            # Se falhar ao gravar no banco, o logger trata o erro internamente
            self.handleError(record)
        finally:
            if conn:
                conn.close()

# --- CONFIGURAÇÃO GERAL ---
def _configure_logger():
    # 1. Primeiro, garante que a tabela existe
    _init_log_table()

    # 2. Configura o objeto Logger
    log_obj = logging.getLogger("RPA_Logger")
    log_obj.setLevel(logging.INFO)

    # Evita duplicar handlers se a função for chamada 2x
    if log_obj.handlers:
        return log_obj

    formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')

    # Handler Console
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(formatter)
    log_obj.addHandler(ch)

    # Handler Banco de Dados
    dbh = DBHandler()
    dbh.setFormatter(formatter)
    log_obj.addHandler(dbh)

    return log_obj

# Exporta o logger configurado
logger = _configure_logger()