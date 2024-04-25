import os
import logging
from datetime import datetime

today = datetime.today()


def get_today():
    return today.strftime('%d/%m/%Y')


def write_log(logs, level=None):
    log_file = os.path.join('logs', f'{today.strftime('%d-%m-%Y')}.log')

    logging.basicConfig(filename=log_file, encoding='utf-8', level=level,
                        format='%(asctime)s - %(name)s - %(levelname)s : %(message)s')

    logger = logging.getLogger(__name__)

    if level == logging.ERROR:
        logger.error(logs)
    elif level == logging.INFO:
        logger.info(logs)
    elif level == logging.CRITICAL:
        logger.critical(logs)
    else:
        logger.exception(logs)


def format_number(number):
    try:
        return '{:,.2f}'.format(float(number))
    except ValueError:
        return number  # Handle non-numeric values gracefully
