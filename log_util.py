import logging.config


# can't configure queue_handler easily as of 3.11 python only available on 3.12
logger = logging.getLogger("Pipeline")


def setup_logging(log_lvl_stdout='INFO') -> None:
    logging_config: dict = {
        "version": 1,
        "disable_existing_loggers": False,
        "formatters": {
            "simple": {
                "format": "%(levelname)s: %(message)s",
            },
            "detailed": {
                "format": "[%(levelname)s|%(module)s|L%(lineno)d] %(asctime)s: %(message)s",
                "datefmt": "%Y-%m-%dT%H:%M:%S%z",
            },
        },
        "handlers": {
            "stdout": {
                "class": "logging.StreamHandler",
                "level": log_lvl_stdout,
                "formatter": "simple",
                "stream": "ext://sys.stdout",
            },
            "file": {
                "class": "logging.handlers.RotatingFileHandler",
                "level": "DEBUG",
                "formatter": "detailed",
                "filename": "logs/Pipeline.log",
                "maxBytes": 5000000, # 5 mb
                "backupCount": 5,
            },
        },
        "loggers": {
            "root": {
                "level": "DEBUG",
                "handlers": [
                    "stdout",
                    "file",
                ],
            },
        },
    }
    logging.config.dictConfig(logging_config)
