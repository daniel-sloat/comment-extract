import logging


def logger_init(func):
    def wrapper():
        logger_start()
        func()
        logger_quit()

    return wrapper


def logger_start():
    logging.basicConfig(
        filename="log.log",
        filemode="w",
        level=logging.INFO,
        datefmt=r"%Y-%m-%d %H:%M:%S",
        format="%(asctime)s.%(msecs)03d [%(levelname)s] %(message)s",
    )
    logging.info("Logging initialized.")


def logger_quit():
    logging.info("Logging ended.")
    logging.shutdown()
