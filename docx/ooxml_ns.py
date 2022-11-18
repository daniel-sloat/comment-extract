import re


ns = {
    "namespaces": {
        # Office Open XML
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
        "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        # Regular expressions for xslt
        "re": "http://exslt.org/regular-expressions",
    },
}


def uri(tag, namespace="w"):
    return f"{{{ns['namespaces'][namespace]}}}{tag}"


def tag(uri_tag):
    return re.sub("{.*}", "", uri_tag)
