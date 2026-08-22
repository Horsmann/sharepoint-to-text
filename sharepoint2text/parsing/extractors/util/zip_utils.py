import zipfile
from typing import Dict, List

from sharepoint2text.parsing import _defused_xml as ET

RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"


def read_zip_text(zf: zipfile.ZipFile, path: str) -> str:
    """Read a text file from a ZIP archive using UTF-8 decoding.

    Args:
        zf: Open ZIP archive containing the requested package data.
        path: Package-relative path of the UTF-8 member.

    Returns:
        Member text decoded as UTF-8, with invalid bytes ignored.

    Raises:
        KeyError: If the requested member does not exist.
    """
    return zf.read(path).decode("utf-8", errors="ignore")


def read_zip_xml_root(zf: zipfile.ZipFile, path: str) -> ET.Element:
    """Parse an XML file from a ZIP archive and return its root element.

    Args:
        zf: Open ZIP archive containing the requested package data.
        path: Package-relative path of the XML member.

    Returns:
        Root element parsed with the safe XML parser.

    Raises:
        KeyError: If the requested member does not exist.
        ParseError: If the member does not contain well-formed XML.
    """
    element: ET.Element = ET.fromstring(zf.read(path))
    return element


def find_relationship_elements(rels_root: ET.Element) -> List[ET.Element]:
    """Return Relationship elements, handling namespace differences.

    Args:
        rels_root: Root element of an OOXML relationships part.

    Returns:
        Relationship elements in document order.
    """
    relationships = rels_root.findall(
        "rel:Relationship", {"rel": RELATIONSHIP_NAMESPACE}
    )
    if relationships:
        return relationships
    return rels_root.findall(f".//{{{RELATIONSHIP_NAMESPACE}}}Relationship")


def parse_relationships(rels_root: ET.Element) -> List[Dict[str, str]]:
    """Normalize Relationship elements into a list of dictionaries.

    Args:
        rels_root: Root element of an OOXML relationships part.

    Returns:
        Relationship dictionaries in source order. Each dictionary contains
        the relationship identifier, type, target, and target mode.
    """
    relationships: List[Dict[str, str]] = []
    for rel in find_relationship_elements(rels_root):
        relationships.append(
            {
                "id": rel.get("Id", ""),
                "type": rel.get("Type", ""),
                "target": rel.get("Target", ""),
                "target_mode": rel.get("TargetMode", ""),
            }
        )
    return relationships
