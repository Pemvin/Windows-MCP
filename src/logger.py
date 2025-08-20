import logging
import sys

def setup_logger():
    """
    Set up the logger for the application.
    为应用程序设置日志记录器。
    """
    logger = logging.getLogger('mcp_logger')
    logger.setLevel(logging.DEBUG)  # Set the lowest level to capture all logs

    # Prevent adding handlers multiple times
    if logger.hasHandlers():
        return logger

    # Formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - [%(module)s:%(lineno)d] - %(message)s'
    )

    # File Handler
    file_handler = logging.FileHandler('app.log', mode='a', encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)  # Log everything to the file
    file_handler.setFormatter(formatter)

    # Stream Handler (Console)
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setLevel(logging.INFO)  # Log INFO and above to the console
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    return logger