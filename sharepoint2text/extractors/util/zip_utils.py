import zipfile
from typing import Dict, List
from xml.etree.ElementTree import Element as XmlElement

from defusedxml import ElementTree as ET

RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"


def read_zip_text(zf: zipfile.ZipFile, path: str) -> str:
    """Read a text file from a ZIP archive using UTF-8 decoding."""
    return zf.read(path).decode("utf-8", errors="ignore")


def read_zip_xml_root(zf: zipfile.ZipFile, path: str) -> XmlElement:
    """Parse an XML file from a ZIP archive and return its root element."""
    return ET.fromstring(zf.read(path))


def find_relationship_elements(rels_root: XmlElement) -> List[XmlElement]:
    """Return Relationship elements, handling namespace differences."""
    relationships = rels_root.findall(
        "rel:Relationship", {"rel": RELATIONSHIP_NAMESPACE}
    )
    if relationships:
        return relationships
    return rels_root.findall(f".//{{{RELATIONSHIP_NAMESPACE}}}Relationship")


def parse_relationships(rels_root: XmlElement) -> List[Dict[str, str]]:
    """Normalize Relationship elements into a list of dictionaries."""
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
