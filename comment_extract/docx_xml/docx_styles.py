from lxml import etree

from .ooxml_ns import *
from .docx_aliases import *

# STYLE INHERITANCE
#   PARAGRAPHS
#       Use default paragraph properties (docDefaults)
#       Append paragraph style properties
#           [Local paragraph properties are only used for
#           list formats and bullets, and are ignored.]
#
#   RUNS
#       Use default run properties (docDefaults)
#       Append run style properties
#       Append local run properties
#
#   COMBINE PARAGRAPHS AND RUN FORMATTING
#       Append result run properties over paragraph properties
#
# Styles can also be based on other styles, and 'inherit' those
# styles' format attributes. And that inherited style may itself
# be based on another style - and so on until the 'base style'.


def get_style_data(
    docx_xml: dict[str : etree.Element],
) -> Styles:
    """Get style names and associated style properties.

    Args:
        docx_xml (etree.Element):
            Keys of zipped path, values of etree.Element at root.

    Returns:
        Styles: Styles dictionary {style name: style properties}
    """
    # docx files have styles, and those styles may be based on other styles.
    style_nodes = docx_xml["word/styles.xml"].xpath(
        "w:style",
        namespaces=ooXMLns,
    )
    styles = {}
    for style_node in style_nodes:
        # Create dict of style names to style properties
        style_name = style_node.xpath(
            "string(@w:styleId)",
            namespaces=ooXMLns,
        )
        style_props = get_rPr_props(style_node)
        styles[style_name] = {"Styles": style_props}

        style_basedon = style_node.xpath(
            "string(w:basedOn/@w:val)",
            namespaces=ooXMLns,
        )
        styles[style_name].update({"BasedOn": style_basedon})
    return styles


def inherit_style_props(styles: Styles) -> Styles:
    """Using based-on style name, traverse style tree until all based-on styles have
    been inherited into the top-level style.

    Args:
        styles (Styles):
                    Styles dictionary with style data (properties dictionary and
                    the based-on style information).

    Returns:
        Styles: Style properties dict directly to style name.
    """
    for style_data in styles.values():
        # While there is a style names in "BasedOn", do:
        while style_data["BasedOn"]:
            # Inherit styles from based-on style (by updating based-on style
            # properties with current style properties.
            based_on_style_props = styles[style_data["BasedOn"]]["Styles"]
            style_data["Styles"] = based_on_style_props | style_data["Styles"]

            # Replace "BasedOn" style name with the name of the style we just
            # inherited with that based-on name (if available).
            # Once "BasedOn" is false (""), then stop while.
            following_style = styles[style_data["BasedOn"]]["BasedOn"]
            style_data["BasedOn"] = following_style
    # Clean up styles dictionary to simple {"StyleName": dict properties}
    styles = {k: v["Styles"] for k, v in [x for x in styles.items()]}
    return styles


def get_styles(docx_xml: dict[str : etree.Element]) -> Styles:
    styles = get_style_data(docx_xml)
    styles = inherit_style_props(styles)
    return styles


def get_doc_default_style_props(
    docx_xml: dict[str : etree.Element],
) -> tuple[StyleProps, StyleProps]:
    doc_defaults = docx_xml["word/styles.xml"].xpath(
        "w:docDefaults", namespaces=ooXMLns
    )
    run_props_default = {}
    para_props_default = {}
    for default in doc_defaults:
        run_doc_default = default.xpath("w:rPrDefault/w:rPr", namespaces=ooXMLns)
        para_doc_default = default.xpath("w:pPrDefault/w:pPr/w:rPr", namespaces=ooXMLns)
        if len(run_doc_default):
            run_props_default = get_rPr_props(run_doc_default[0])
        if len(para_doc_default):
            para_props_default = get_rPr_props(para_doc_default[0])
    return run_props_default, para_props_default


def get_para_style_props(paragraph: etree.Element, styles: Styles) -> StyleProps:
    """Find paragraph style properties.

    Args:
        paragraph: (etree.Element): Paragraph element node.
        styles (Styles): Dict of style names and formats.

    Returns:
        StyleProps: Paragraph style formats.
    """
    # The run properties of the paragraph are used to format the paragraph marker.
    # This formatting is used in Word to format list numbers and bullets.
    # https://stackoverflow.com/questions/41435869/why-isnt-this-ooxml-text-bold
    # Therefore, paragraph run properties are not applicable to extracting comments,
    # as this script ignores any list numbers or bullets.
    para_prop_node = paragraph.xpath("w:pPr", namespaces=ooXMLns)
    para_style = {}
    if para_prop_node:
        para_style_name = para_prop_node[0].xpath(
            "string(w:pStyle/@w:val)", namespaces=ooXMLns
        )
        para_style = styles.get(para_style_name, {})
        # Paragraph run formatting (not needed)
        # para_formats = get_rPr_props(para_prop_node[0])
    return para_style


def get_rPr_props(element: etree.Element) -> StyleProps:
    """Returns selected formatting elements under 'rPr' node as a dictionary.

    Args:
        element (etree.Element): Parent node of rPr node.

    Returns:
        StyleProps: Example: {"b": {}, "u": {"val": "single"}}
    """
    # http://officeopenxml.com/WPtextFormatting.php
    # Remove unwanted styles to clean up the format dictionary.
    # 'w:rStyle' must be included to apply run-specific style names.
    included_nodes = (
        "w:rStyle",
        "w:b",
        "w:i",
        "w:u",
        "w:strike",
        "w:dstrike",
        "w:vertAlign",
    )
    run_prop_nodes = element.xpath(
        f"w:rPr/{'|w:rPr/'.join(included_nodes)}", namespaces=ooXMLns
    )
    run_formats = {}
    for node in run_prop_nodes:
        # Remove namespace prefix (e.g., 'w:') for ease of use later
        name = etree.QName(node).localname
        formats = {}
        for attribute in node.attrib:
            attrib_tag = etree.QName(attribute).localname
            formats[attrib_tag] = node.attrib[attribute]
        run_formats[name] = formats
    return run_formats


def get_run_style_props(run_props: StyleProps, styles: Styles) -> StyleProps:
    """Update run props with run style properties, if stated.

    Args:
        run_props (StyleProps): _description_
        styles (Styles): _description_

    Returns:
        StyleProps
    """
    if "rStyle" in run_props:
        run_style_name = run_props["rStyle"]["val"]
        style_props = styles[run_style_name]
        return style_props
    return {}


def get_format_props(
    paragraph: etree.Element,
    run: etree.Element,
    para_props_def: StyleProps,
    run_props_def: StyleProps,
    styles: Styles,
) -> StyleProps:
    run_direct_props = get_rPr_props(run)
    run_style_props = get_run_style_props(run_direct_props, styles)
    run_props = run_props_def | run_style_props | run_direct_props

    para_style_props = get_para_style_props(paragraph, styles)
    para_props = para_props_def | para_style_props

    return para_props | run_props
