from collections import UserDict

from lxml.etree import QName


class attrib_dict(UserDict):
    """UserDict that removes URI from element tag string in key."""

    def __setitem__(self, key, value):
        qkey = QName(key).localname
        self.data[qkey] = value


def get_attrib(prop_element):
    d = attrib_dict()
    for prop in prop_element:
        if len(prop) > 0:
            d[prop.tag] = get_attrib(prop)
        else:
            d[prop.tag] = attrib_dict(prop.attrib)
    return d
