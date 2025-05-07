import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s - %(message)s',
    handlers=[
        logging.FileHandler("main.log", encoding='utf-8'),
        logging.StreamHandler()
    ],
    force=True
)

import process_emails

process_emails.main()