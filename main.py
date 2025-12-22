# Importando bibliotecas
import sys
from utils.logger import logger
from services.outlook import outlook_process

def main():
    logger.info("=== RPA Iniciado ===")

    try:
        # Executa o processo principal de negócio
        outlook_process()

    except KeyboardInterrupt:
        # Caso você pare o robô manualmente no terminal
        logger.warning("Execução interrompida manualmente pelo usuário.")
        
    except Exception as e:
        # Última barreira de defesa: captura erros de importação, sintaxe ou falhas graves
        logger.critical(f"Erro fatal não tratado na execução principal: {e}")
        # Opcional: sys.exit(1) informaria ao Windows Task Scheduler que houve falha

    finally:
        # Executa sempre, com sucesso ou com erro, para fechar o bloco de logs
        logger.info("=== RPA Finalizado ===")

if __name__ == "__main__":
    main()