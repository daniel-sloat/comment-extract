"""Logging support functions."""
import functools
import logging


def logger_init(print_console=True):
    def inner(func):
        def wrapper():
            logger_start(print_console)
            func()
            logger_quit()

        return wrapper

    return inner


def logger_start(print_console=True):
    logging.basicConfig(
        filename="log.log",
        filemode="w",
        level=logging.INFO,
        datefmt=r"%Y-%m-%d %H:%M:%S",
        format="%(asctime)s.%(msecs)03d [%(levelname)s] %(message)s",
    )
    if print_console:
        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        formatter = logging.Formatter(
            "%(asctime)s (elapsed: %(relativeCreated)dms) [%(levelname)s] %(message)s"
        )
        console.setFormatter(formatter)
        logging.getLogger().addHandler(console)
    logging.info("Logging initialized.")


def logger_quit():
    logging.info("Logging ended.")
    logging.shutdown()


def log_filename(func):
    @functools.wraps(func)
    def wrapper(filename):
        logging.info("Reading file %s", filename)
        result = func(filename)
        return result

    return wrapper


def log_comments(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        result = func(*args, **kwargs)
        if (replies := len(result._all_comments) - len(result)) != 0:
            logging.info("Read %s comments (with %s replies).", len(result), replies)
        else:
            logging.info("Read %s comments.", len(result))
        return result

    return wrapper


def log_total_record(func):
    @functools.wraps(func)
    def wrapper(self, *args, **kwargs):
        result = func(self, *args, **kwargs)
        logging.info("Total of %s comments from %s documents.", self.total, len(self))
        return result

    return wrapper
