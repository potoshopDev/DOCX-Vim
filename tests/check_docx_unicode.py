"""Early fixture check for Cyrillic integrity in the supplied DOCX files."""

from pathlib import Path
from zipfile import BadZipFile, ZipFile
import sys
import xml.etree.ElementTree as ET


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FIXTURES = Path(__file__).resolve().parents[1] / "e"


def check(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            xml = archive.read("word/document.xml")
        root = ET.fromstring(xml)
        content = "".join(node.text or "" for node in root.iter(W + "t"))
    except (BadZipFile, KeyError, ET.ParseError) as exc:
        print(f"FAIL {path.name}: invalid DOCX/XML: {exc}")
        return False

    cyrillic = sum("\u0400" <= char <= "\u052f" for char in content)
    replacement = content.count("\ufffd")
    valid = cyrillic > 0 and replacement == 0
    status = "OK" if valid else "FAIL"
    print(f"{status} {path.name}: Cyrillic={cyrillic}, replacement={replacement}")
    return valid


def main() -> int:
    paths = sorted(FIXTURES.glob("*.docx"))
    if not paths:
        print(f"No DOCX fixtures in {FIXTURES}")
        return 1
    return 0 if all(check(path) for path in paths) else 1


if __name__ == "__main__":
    sys.exit(main())
