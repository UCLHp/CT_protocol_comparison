import xml.etree.ElementTree as ET


def read_xml(filepath):
    tree = ET.parse(filepath)
    root = tree.getroot()

    return root


def print_xml(root):
    for element in root.iter():
        print(
            "Element:", element.tag,
            "| Attributes:", element.attrib,
            "| Text:", element.text
        )