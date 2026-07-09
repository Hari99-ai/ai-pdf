import logging
import sys
from pathlib import Path

# Create a log file inside the workspace
LOG_FILE_PATH = Path(__file__).parent.parent / "extraction.log"

def setup_logger():
    logger = logging.getLogger("pdf_extractor")
    logger.setLevel(logging.INFO)
    
    # Avoid duplicate handlers if already configured
    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] [%(filename)s:%(lineno)d]: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # File handler
    try:
        file_handler = logging.FileHandler(LOG_FILE_PATH, encoding="utf-8")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as e:
        print(f"Warning: Could not create file handler for logging: {e}", file=sys.stderr)

    return logger

logger = setup_logger()
