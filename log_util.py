import logging.config


# can't configure queue_handler easily as of 3.11 python only available on 3.12
logger = logging.getLogger("Pipeline")
logging_config: dict = {
    "version": 1,
    "disable_existing_loggers": False,
    "formatters": {
        "simple": {
            "format": "%(levelname)s: %(message)s",
        },
        "detailed": {
            "format": "[%(levelname)s|%(module)s|L%(lineno)d] %(asctime)s: %(message)",
            "datefmt": "%Y-%m-%dT%H:%M:%S%z",
        },
    },
    "handlers": {
        "stderr": {
            "class": "logging.StreamHandler",
            "level": "WARNING",
            "formatter": "simple",
            "stream": "ext://sys.stderr",
        },
        "file": {
            "class": "logging.handlers.RotatingFileHandler",
            "level": "DEBUG",
            "formatter": "detailed",
            "filename": "logs/Pipeline.log",
            "maxBytes": 150000,
            "backupCount": 5,
        },
    },
    "loggers": {
        "root": {
            "level": "DEBUG",
            "handlers": [
                "stderr",
                "file",
            ],
        },
    },
}
logging.config.dictConfig(logging_config)
