from lxml import etree

from zipfile import ZipFile
from itertools import groupby
from datetime import datetime
import re

from .ooxml_ns import *
from .docx_styles import *
from .docx_aliases import *


def get_docx_xml_tree(
    document_path: str,
) -> dict[str : etree.Element]:
    """Gets dictionary of relevant Office Open XML root element nodes in docx.

    Args:
        document_path (str): Document location path.

    Returns:
        dict[str:etree.Element]: Returns dict with zipped filepath as keys
            and values of root etree element.
    """
    with ZipFile(document_path, "r") as z:
        docx_xml_tree = {}
        regex = r"^word/(?:document|styles|comments|commentsExtended|footnotes|endnotes)\.xml$"
        for xml_file in [name for name in z.namelist() if re.search(regex, name)]:
            docx_xml_tree[xml_file] = etree.fromstring(z.read(xml_file))
    return docx_xml_tree


def doc_comment_integrity_check(
    docx_xml: dict[str : etree.Element],
) -> None:
    """Check whether document and comment data are valid.

    Args:
        docx_xml (dict[str:etree.Element]):
            Keys of zipped path, values of etree.Element at root.

    Returns:
        None
    """
    count_equal = (
        docx_xml["word/document.xml"].xpath(
            "count(//w:commentRangeStart)", namespaces=ooXMLns
        )
        == docx_xml["word/document.xml"].xpath(
            "count(//w:commentRangeEnd)", namespaces=ooXMLns
        )
        == docx_xml["word/comments.xml"].xpath("count(//w:comment)", namespaces=ooXMLns)
    )
    assert count_equal == True, "Comment data within document appears corrupt."
    return None


def get_filename_info(
    filename: str,
    delimiter: str,
) -> dict[str : str | None]:
    stem = filename.rsplit(".")[0]
    if delimiter and delimiter in stem:
        split = stem.split(delimiter, maxsplit=1)
        codes = {
            "File Name": filename,
            "Document Number": int(split[0].strip()),
            "Commenter Code": split[1].strip(),
        }
    else:
        codes = {
            "File Name": filename,
            "Document Number": None,
            "Commenter Code": stem.strip(),
        }
    return codes


def get_comment_bounds(
    docx_xml: dict[str : etree.Element],
) -> dict[int : etree.Element]:
    """Gets the comment start and end nodes associated with a comment.

    Args:
        docx_xml (dict[str:etree.Element]):
            Keys of zipped path, values of etree.Element at root.

    Returns:
        dict[int:etree.Element]: Sorted and numbered dict with start/end nodes as values.
    """
    # Using only the commentRangeStart ids to get the comment ids.
    # This will grab comment ids in document order.
    comment_ids = [
        int(id)
        for id in docx_xml["word/document.xml"].xpath(
            "//w:commentRangeStart/@w:id",
            namespaces=ooXMLns,
        )
    ]
    comment_bounds = {}
    for id in comment_ids:
        start_and_end_nodes = docx_xml["word/document.xml"].xpath(
            rf"//w:commentRangeStart[@w:id={id}]|//w:commentRangeEnd[@w:id={id}]",
            namespaces=ooXMLns,
        )
        comment_bounds[id] = start_and_end_nodes
    return comment_bounds


def get_comment_paragraphs(
    comment_bounds: dict[int : etree.Element],
) -> dict[int : list[etree.Element]]:
    """Get comment paragraphs, return list of comments with paragraphs nested.

    Args:
        comment_bounds (dict[int:etree.Element]):
            Sorted and numbered dict with start/end nodes as values.

    Returns:
        dict[int:list]: [Comments[Paragraphs]]
    """
    comment_paragraphs = {}
    for id, data in comment_bounds.items():
        # commentRanges may be children of paragraphs, or siblings.
        # xpaths grab the parent paragraph if there is one, or the following/
        # preceding paragraph if there isn't.
        start_paragraph = data[0].xpath(
            "parent::w:p|following-sibling::w:p[1]",
            namespaces=ooXMLns,
        )[0]
        end_paragraph = data[1].xpath(
            "parent::w:p|preceding-sibling::w:p[1]",
            namespaces=ooXMLns,
        )[0]
        # xpath selects all paragraphs starting from start, and
        # returns only paragraphs that have text or a commentRangeEnd within them,
        # and text is not just whitespace.
        xpath = "(self::w:p|following-sibling::w:p)\
            [(w:r/w:t or w:commentRangeEnd) and not(re:test(string(.),'^\s+$'))]"
        # Generator reduces memory load (otherwise would get rest of document
        # after start paragraph).
        paragraphs = (
            x for x in start_paragraph.xpath(xpath, namespaces=ooXMLns | regexns)
        )
        paras = []
        for para in paragraphs:
            paras.append(para)
            if para == end_paragraph:
                break
        comment_paragraphs[id] = paras
    return comment_paragraphs


def remove_comment_replies(
    docx_xml: dict[str : etree.Element],
    comment_paragraphs: dict[int : list[etree.Element]],
) -> dict[int : list[etree.Element]]:
    """Remove comments that is an ancestor of another comment (a reply).

    Args:
        docx_xml (dict[str : etree.Element]): Must include "word/commentsExtended.xml"
            and "word/comments.xml"

    Returns:
        list[int]: Comment IDs that are replies.
    """
    # https://blogs.msmvps.com/wordmeister/2012/09/29/word2013comments-wordopenxml/
    # OOXML does not use comment ids to link replies, but instead uses a paraIdParent
    # field that has a value of the paraId of the last paragraph in the comment.
    # Get paraId that are replies. These paraIds will be removed from the comment set
    # because they have a 'parent' comment (ancestor).
    para_ids_parent = docx_xml["word/commentsExtended.xml"].xpath(
        "w15:commentEx/@w15:paraId[../@w15:paraIdParent]", namespaces=ooXMLns
    )
    # Take last paragraph in comment, which is used to compare to commentsExtended.xml.
    para_ids = docx_xml["word/comments.xml"].xpath(
        "w:comment/w:p[last()]/@w14:paraId", namespaces=ooXMLns
    )
    comment_ids = [
        int(x)
        for x in docx_xml["word/comments.xml"].xpath(
            "w:comment/@w:id", namespaces=ooXMLns
        )
    ]
    remove_comment_ids = []
    # Gets the comment ids that are ancestors of comments.
    if para_ids_parent:
        remove_comment_ids.extend(
            [
                comment_id
                for comment_id, para_id in zip(comment_ids, para_ids)
                for parent in para_ids_parent
                if parent in para_id
            ]
        )
    # Remove those ancestors from the comments.
    for comment_id in remove_comment_ids:
        comment_paragraphs.pop(comment_id)
    return comment_paragraphs


def get_paragraph_runs(
    paragraph: etree.Element,
    comment_no: int,
) -> list:
    """Gets runs in paragraph after the start of a specific comment number
    and before the end of a specific comment number. Paragraphs must be only those
    associated with comments.

    Args:
        paragraph (etree.Element): Paragraph element.
        comment_no (int): Comment number.

    Returns:
        list: Select run elements between start and end of a comment.
    """

    def get_runs(children):
        select_runs = []
        for child in children:
            # Selects runs that only include text or are a footnote/endnote reference.
            # Sometimes runs can include drawings and other things.
            # Stops selecting if commentRangeEnd encountered in paragraph.
            if child.xpath(
                "self::w:r[w:t|w:footnoteReference|w:endnoteReference]",
                namespaces=ooXMLns,
            ):
                select_runs.append(child)
            elif child.xpath(
                f"self::w:commentRangeEnd[@w:id={comment_no}]",
                namespaces=ooXMLns,
            ):
                break
        return select_runs

    # Selects all paragraph runs and commentRangeStart/Ends within the paragraph.
    # commentRangeStart/End can appear at paragraph level or at run level.
    # The generator is important for selecting runs after a start is found.
    para_elements = (
        el
        for el in paragraph.xpath(
            f"w:r|w:commentRangeStart[@w:id={comment_no}] | w:commentRangeEnd[@w:id={comment_no}]",
            namespaces=ooXMLns,
        )
    )
    comment_start_paragraph = paragraph.xpath(
        f"boolean(w:commentRangeStart[@w:id={comment_no}])",
        namespaces=ooXMLns,
    )

    if comment_start_paragraph:
        # Begin appending runs after the commentRangeStart (if present).
        for element in para_elements:
            if element.xpath(
                f"boolean(self::w:commentRangeStart[@w:id={comment_no}])",
                namespaces=ooXMLns,
            ):
                runs = get_runs(para_elements)
    else:
        runs = get_runs(para_elements)
    return runs


def get_text(element: etree.Element) -> str:
    return element.xpath("string(.)", namespaces=ooXMLns)


def clean_text(text: str) -> str:
    text = re.sub(r"\s{2,}", " ", text)  # Replace 2 spaces or more with single space
    text = re.sub(r"\s", " ", text)  # Replace any whitespace with regular space
    text = re.sub(r"“|”", '"', text)  # Replace curly quote characters
    return text


def encode_run_props(
    props: StyleProps,
    ignored_formats: dict[str:bool],
) -> str:
    """Encode formatting from docx type names and properties to string with each
    characters representing a character.

    Args:
        props (dict[str:dict]): Run properties ({'b': {'val': '1'}, 'i': {}})
        ignored_formats (dict): Formats to ignore ({'bold': false}).

    Returns:
        str: Encoded string of characters that represent formatting.
    """
    formats = []
    for k in props:
        # b, i, strike, dstrike are toggle formats. They can be 'on' based on
        # multiple different values (whether there is just the format name like 'b'
        # and no attributes/value associated), or if the value attribute indicates 'on'
        # (like '1', 'on', or 'true')
        toggle_on = (
            not props[k]
            or props[k].get("val") == "1"
            or props[k].get("val") == "on"
            or props[k].get("val") == "true"
        )
        match k:
            case "b" if toggle_on and not ignored_formats["bold"]:
                formats.append("b")
            case "i" if toggle_on and not ignored_formats["italic"]:
                formats.append("i")
            case "strike" if toggle_on and not ignored_formats["strikethrough"]:
                formats.append("s")
            case "dstrike" if toggle_on and not ignored_formats["double_strikethrough"]:
                formats.append("z")
            case "u":
                # There are many different types of underline in word.
                # Convert all complex underline (e.g., wavy, dotted, etc.) to basic
                # single underline, unless the value contains 'double'.
                underline_type = props["u"].get("val", "single")
                if (
                    re.search("[D|d]ouble", underline_type)
                    and not ignored_formats["double_underline"]
                ):
                    formats.append("w")
                elif underline_type == "none":
                    break
                elif not ignored_formats["underline"]:
                    formats.append("u")
            case "vertAlign":
                if (
                    props["vertAlign"].get("val") == "subscript"
                    and not ignored_formats["subscript"]
                ):
                    formats.append("v")
                elif (
                    props["vertAlign"].get("val") == "superscript"
                    and not ignored_formats["superscript"]
                ):
                    formats.append("x")
    # Sort to make sure all in order
    formats = "".join(sorted(formats))
    return formats


def footnote_or_endnote(element: etree.Element) -> bool:
    """Determine whether element is a footnote or endnote."""
    boolean = element.xpath(
        "boolean(w:footnoteReference/@w:id or w:endnoteReference/@w:id)",
        namespaces=ooXMLns,
    )
    return boolean


def get_footnotes_and_endnotes_data(
    docx_xml: dict[str : etree.Element],
    styles: Styles,
    para_props_def: StyleProps,
    run_props_def: StyleProps,
    ignored_formats: dict[str:bool],
) -> dict[str:DataLinkedToComment]:
    # Check if footnotes or endnotes even exist
    fn_en_exist = any(
        ["word/footnotes.xml" in docx_xml, "word/endnotes.xml" in docx_xml]
    )
    if not fn_en_exist:
        return {}

    # Predicated greater than zero because footnotes/endnotes begin at 1, and
    # there may be extraneous footnote/endnote data at id 0 or -1.
    footnote_data = docx_xml["word/footnotes.xml"].xpath(
        "w:footnote[@w:id>0]", namespaces=ooXMLns
    )
    endnote_data = docx_xml["word/endnotes.xml"].xpath(
        "w:endnote[@w:id>0]", namespaces=ooXMLns
    )
    d = {
        "Footnotes": footnote_data,
        "Endnotes": endnote_data,
    }

    footnotes_and_endnotes = {}
    for name, data in d.items():
        footnotes_and_endnotes[name] = {}
        for note in data:
            para_data = []
            paragraphs = note.xpath("w:p", namespaces=ooXMLns)
            note_no = int(note.xpath("string(@w:id)", namespaces=ooXMLns))
            for paragraph in paragraphs:
                run_data = []
                para_runs = paragraph.xpath("w:r[w:t]", namespaces=ooXMLns)
                for run in para_runs:
                    run_text = get_text(run)
                    run_props = get_format_props(
                        paragraph, run, para_props_def, run_props_def, styles
                    )
                    run_props = encode_run_props(run_props, ignored_formats)
                    run_data.append([run_text, run_props])
                para_data.append(run_data)
            footnotes_and_endnotes[name].update({note_no: para_data})
    return footnotes_and_endnotes


def renumber_footnotes_and_endnotes(
    comment_paragraphs: dict[int : list[etree.Element]],
    fn_en: dict[str:DataLinkedToComment],
) -> DataLinkedToComment:
    """Renumber footnotes and endnotes on a per-comment basis, so that each comment
    footnote or endnote begins at 1. Differences between footnote and endnote numbering
    are ignored, so if there is a footnote and endnote in a comment, they are numbered
    sequentially."""
    notes = {}
    for comment_id, paragraphs in comment_paragraphs.items():
        fnanden = []
        note_renumber = 0
        for run in [r for paragraph in paragraphs for r in paragraph]:
            xpath = r"w:footnoteReference | w:endnoteReference"
            note_nodes = (x for x in run.xpath(xpath, namespaces=ooXMLns))
            for node in note_nodes:
                id = int(node.xpath("string(@w:id)", namespaces=ooXMLns))
                note_renumber += 1
                if node.xpath("self::w:footnoteReference", namespaces=ooXMLns):
                    fn = fn_en["Footnotes"][id]
                    for paragraph in fn:
                        if paragraph == fn[0]:
                            paragraph.insert(0, [f"{note_renumber}", "x"])
                    fnanden.extend(fn)
                elif node.xpath("self::w:endnoteReference", namespaces=ooXMLns):
                    en = fn_en["Endnotes"][id]
                    for paragraph in en:
                        if paragraph == en[0]:
                            paragraph.insert(0, [f"{note_renumber}", "x"])
                    fnanden.extend(en)
                notes[comment_id] = fnanden
    return notes


def consolidate_runs(comment_data: NestedRunData) -> NestedRunData:
    """Many times, runs may have following runs with the same format, but the
    runs may be split arbitrarily. This groups those similar runs and formats,
    and can significantly reduce file-size and make the file render better when
    veiwing.

    [['Text ', ''], ['More text ', ''], ['Bold ', 'b'], ['Not bold ', '']] ->
    [['Text More text ', ''], ['Bold ', 'b'], ['Not bold ', '']]
    """
    grouped_comment = []
    for paragraph in comment_data:
        grouped_runs = []
        for key_format, group in groupby(paragraph, lambda x: x[1]):
            data = []
            for run in group:
                data.append(run[0])
            text = "".join(data)
            # This is a convenient time to remove extra whitespaces, etc.
            text = clean_text(text)
            grouped_runs.append([text, key_format])
        grouped_comment.append(grouped_runs)
    return grouped_comment


def insert_footnote_or_endnote(
    comment_data: NestedRunData,
    notes: DataLinkedToComment,
    comment_id: int,
) -> NestedRunData:
    # Add footnote/endnote text to end of comment
    if comment_id in notes:
        comment_data.extend(notes[comment_id])
    return comment_data


def if_empty_comment(comment_data: NestedRunData) -> NestedRunData:
    # If no paragraphs, make empty comment
    if not comment_data:
        comment_data = [[["(( Empty comment ))", ""]]]
    return comment_data


def get_comment_data(
    comment_paras: dict[int : list[etree.Element]],
    notes: DataLinkedToComment,
    styles: Styles,
    para_props_def: StyleProps,
    run_props_def: StyleProps,
    ignored_formats: dict[str:bool],
) -> list[CommentRecordData]:
    """Iterate over comments, paragraphs, and runs"""
    comment_data = []
    for no, (comment_id, paragraphs) in enumerate(comment_paras.items(), 1):
        para = []
        footnote_endnote_num = 0
        for paragraph in paragraphs:
            run_data = []
            runs = get_paragraph_runs(paragraph, comment_id)
            if runs:
                for run in runs:
                    run_text = get_text(run)

                    if run == runs[0]:
                        run_text = run_text.lstrip()
                        if not run_text:
                            continue
                    elif run == runs[-1]:
                        run_text = run_text.rstrip()
                        if not run_text:
                            continue

                    run_props = get_format_props(
                        paragraph, run, para_props_def, run_props_def, styles
                    )
                    run_props = encode_run_props(run_props, ignored_formats)

                    if footnote_or_endnote(run):
                        # Combining footnotes and endnotes may create issues with numbering,
                        # also it starts the notes at 1 for each comment.
                        footnote_endnote_num += 1
                        run_text = f"{footnote_endnote_num}"
                        run_props = "x"

                    run_data.append([run_text, run_props])
                para.append(run_data)
            else:
                break
        para = insert_footnote_or_endnote(para, notes, comment_id)
        para = if_empty_comment(para)
        para = consolidate_runs(para)
        comment_data.append(
            {
                "Document Comment Number": no,
                "Original Comment ID": comment_id,
                "Comment Data": para,
            }
        )
    return comment_data


def get_comment_bubble_data(
    docx_xml: dict[str : etree.Element],
    delimiter: str,
) -> dict[int:CommentRecordData]:
    """Get comment bubble data"""
    comment_bubbles = docx_xml["word/comments.xml"].xpath(
        "w:comment", namespaces=ooXMLns
    )
    comment_ids = [
        int(x)
        for x in docx_xml["word/comments.xml"].xpath(
            "w:comment/@w:id", namespaces=ooXMLns
        )
    ]
    comment_bubble_data = {}
    for comment_bubble, comment_id in zip(comment_bubbles, comment_ids):
        text = []
        comment_text_paragraphs = comment_bubble.xpath("w:p", namespaces=ooXMLns)
        for paragraph in comment_text_paragraphs:
            if paragraph != comment_text_paragraphs[0]:
                text.append("\n")
            text.append(paragraph.xpath("string(.)", namespaces=ooXMLns))
        comment_text = "".join(text)
        split = comment_text.split(delimiter, maxsplit=1)
        if delimiter and len(split) == 2:
            # Commenter Code updates Commenter Code from filename information.
            comment_bubble_data[comment_id] = {
                "Comment Bubble": comment_text.strip(),
                "Heading 1": split[0].strip(),
                "Commenter Code": split[1].strip(),
            }
        else:
            comment_bubble_data[comment_id] = {
                "Comment Bubble": comment_text.strip(),
                "Heading 1": comment_text.strip(),
            }
    return comment_bubble_data


def create_date_types(date: str) -> datetime:
    try:
        if date[-1] == "Z":
            return datetime.strptime(date, r"%Y-%m-%dT%H:%M:%SZ")
        else:
            return datetime.strptime(date, r"%Y-%m-%dT%H:%M:%S")
    except:
        return date


def get_comment_metadata(
    docx_xml: dict[str : etree.Element],
) -> dict[int:CommentRecordData]:
    """Get comment metadata associated with comment. This data includes the comment
    author, comment author initials, and comment date.

    Args:
        docx_xml (dict[str : etree.Element]): Must include "word/comments.xml"

    Returns:
        dict[int : dict[str : str | datetime]]: Comment Author, Comment Author
            Initials, Comment Date
    """
    comment_data = docx_xml["word/comments.xml"].xpath("w:comment", namespaces=ooXMLns)
    comments_metadata = {}
    for comment_node in comment_data:
        comment_id = int(comment_node.xpath("number(@w:id)", namespaces=ooXMLns))
        comment_author = comment_node.xpath("string(@w:author)", namespaces=ooXMLns)
        comment_date = comment_node.xpath("string(@w:date)", namespaces=ooXMLns)
        # Make string as datetime type, which is interpreted as date format by xlsxwriter
        comment_date = create_date_types(comment_date)
        comment_author_initials = comment_node.xpath(
            "string(@w:initials)", namespaces=ooXMLns
        )
        comments_metadata[comment_id] = {
            "Comment Author": comment_author,
            "Comment Author Initials": comment_author_initials,
            "Comment Date": comment_date,
        }
    return comments_metadata


def create_comment_record(
    comment_data: list[CommentRecordData],
    filename_info: CommentRecordData,
    comment_bubbles: dict[int:CommentRecordData],
    comment_metadata: dict[int:CommentRecordData],
) -> list[CommentRecordData]:
    """_summary_

    Returns:
        list[CommentRecordData]: _description_
    """
    comment_record = []
    for comment in comment_data:
        # Get data that requires looking up the original comment ID, then delete it
        comment_bubble = comment_bubbles[comment["Original Comment ID"]]
        comment_datum = comment_metadata[comment["Original Comment ID"]]
        comment.pop("Original Comment ID")

        
        comment = comment | filename_info | comment_bubble | comment_datum
        # Sort comment record
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
            "Comment Data",
        ]
        comment = {k: comment[k] for k in keyorder if k in comment}
        comment_record.append(comment)
    return comment_record


def read_comments(
    file: str,
    filename_delimiter: str,
    bubble_delimiter: str,
    ignored_styles: dict[str:bool],
    include_replies: bool
) -> list[CommentRecordData]:
    """_summary_

    Args:
        file (str): _description_
        filename_delimiter (str): _description_
        comment_bubble_delimiter (str): _description_
        ignored_styles (dict[str:bool]): _description_

    Returns:
        list[CommentRecordData]: _description_
    """
    docx_xml = get_docx_xml_tree(file)
    comment_record = []
    if "word/comments.xml" in docx_xml:
        doc_comment_integrity_check(docx_xml)
        comment_bounds = get_comment_bounds(docx_xml)
        comment_paragraphs = get_comment_paragraphs(comment_bounds)

        if not include_replies:
            comment_paragraphs = remove_comment_replies(docx_xml, comment_paragraphs)
            
        default_props = get_doc_default_style_props(docx_xml)
        style_props = get_styles(docx_xml)
        comment_bubble_data = get_comment_bubble_data(docx_xml, bubble_delimiter)
        comment_notes_data = get_footnotes_and_endnotes_data(
            docx_xml, style_props, *default_props, ignored_styles
        )
        comment_notes_data = renumber_footnotes_and_endnotes(
            comment_paragraphs, comment_notes_data
        )
        filename_info = get_filename_info(file.name, filename_delimiter)
        metadata = get_comment_metadata(docx_xml)
        comment_data = get_comment_data(
            comment_paragraphs,
            comment_notes_data,
            style_props,
            *default_props,
            ignored_styles,
        )
        comment_record = create_comment_record(
            comment_data,
            filename_info,
            comment_bubble_data,
            metadata,
        )
    return comment_record
