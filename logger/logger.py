import logging
import functools


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


def log_filename(func):
    @functools.wraps(func)
    def wrapper(filename):
        logging.info(f"Reading file {filename}")
        result = func(filename)
        return result

    return wrapper


def log_comments(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        result = func(*args, **kwargs)
        logging.info(f"Reading file {len(result)}")
        return result

    return wrapper
