from utils.logger import logger
from core.outlook import outlook_process

def main():
    logger.info("=== RPA Iniciado ===")
    outlook_process()



if __name__ == "__main__":
    main()