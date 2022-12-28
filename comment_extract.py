import tomli

from pathlib import Path
import logging

from comment_extract import docx_xml as dx, write_xlsx as xl


def initialize_logging():
    logging.basicConfig(
        filename="log.log",
        filemode="w",
        level=logging.INFO,
        datefmt=r"%Y-%m-%d %H:%M:%S",
        format="%(asctime)s.%(msecs)03d [%(levelname)s] %(message)s",
    )
    logging.info("Comment extract script initialized.")
    return None


def load_toml_config(
    config_filename: str = "config.toml",
) -> dict:
    with open(config_filename, "rb") as f:
        return tomli.load(f)


def get_files(
    folder_path: str,
    globex: str = "*.docx",
) -> list[Path]:
    logging.info(f"Reading files from: {folder_path}")
    p = Path(folder_path).glob(globex)
    return [x for x in p if x.is_file()]


def add_columns(comment_record):
    addl_columns = {"Heading 2": None, "Heading 3": None, "Response": None}
    new_comment_record = []
    for comment in comment_record:
        comment |= addl_columns

        keyorder = [
            "File Name",
            "Document Number",
            "Commenter Code",
            "Document Comment Number",
            "Comment Author",
            "Comment Author Initials",
            "Comment Date",
            "Comment Bubble",
            "Heading 1",
            "Heading 2",
            "Heading 3",
            "Comment Data",
            "Response",
        ]
        comment = {k: comment[k] for k in keyorder if k in comment}
        new_comment_record.append(comment)
    return new_comment_record


def extract_comments(
    files: list[Path],
    config_file: dict[str : str | dict[str:bool]],
) -> list[dx.CommentRecordData]:
    file_data = []
    total = len(files)
    for count, file in enumerate(files, 1):
        logging.info(f"Reading from {file.name}")
        comment_record = dx.read_comments(
            file,
            config_file["filename_delimiter"],
            config_file["comment_bubble_delimiter"],
            config_file["ignore_formatting"],
            config_file["include_replies"]
        )
        if config_file["response"]["add_columns"]:
            comment_record = add_columns(comment_record)
        file_data.extend(comment_record)
        print(f"Progress: {count}/{total}", end="\r")
        logging.info(f"Read {len(comment_record)} comments.")
    finish = f"Total of {len(file_data)} comments from {len(files)} files."
    print(finish)
    logging.info(finish)
    return file_data


def create_workbook(
    file_data: list[dx.CommentRecordData],
    config_file: dict,
) -> None:
    logging.info("Creating workbook...")
    filename = xl.create_workbook(
        file_data, add_columns=config_file["response"]["add_columns"]
    )
    logging.info(f"Workbook created: {filename}")
    return None


def quit_logging() -> None:
    logging.info("Finished.")
    logging.shutdown()
    return None


def main() -> None:
    initialize_logging()
    config_file = load_toml_config()
    files = get_files(config_file["folder_path"], globex="*.docx")
    file_data = extract_comments(files, config_file)
    create_workbook(file_data, config_file)
    quit_logging()
    return None


if __name__ == "__main__":
    main()
