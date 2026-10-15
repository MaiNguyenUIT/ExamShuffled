"""Operate on original OOXML; never round-trip exam content through plain text."""
from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
NS = {"w": W[1:-1], "r": R[1:-1]}
MAX_UPLOAD_BYTES = 30 * 1024 * 1024
MAX_PACKAGE_BYTES = 200 * 1024 * 1024


class DocumentError(ValueError):
    """Actionable input/structure error suitable for display to the user."""


def read_package(data: bytes) -> tuple[dict[str, bytes], etree._Element]:
    if len(data) > MAX_UPLOAD_BYTES:
        raise DocumentError("DOCX exceeds the 30 MiB upload limit.")
    try:
        with ZipFile(BytesIO(data)) as package:
            infos = package.infolist()
            if len(infos) > 10000 or sum(i.file_size for i in infos) > MAX_PACKAGE_BYTES:
                raise DocumentError("DOCX expands beyond the supported package limits.")
            if len({i.filename for i in infos}) != len(infos):
                raise DocumentError("DOCX contains duplicate ZIP entries.")
            parts = {i.filename: package.read(i) for i in infos}
        xml = parts["word/document.xml"]
        parser = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False)
        root = etree.fromstring(xml, parser)
        if root.getroottree().docinfo.doctype:
            raise DocumentError("DTD declarations are unsupported in DOCX XML.")
        if root.tag != W + "document" or root.find(W + "body") is None:
            raise DocumentError("Unsupported Word XML namespace or missing document body.")
        return parts, root
    except (BadZipFile, KeyError, etree.XMLSyntaxError, RuntimeError) as exc:
        raise DocumentError("Cannot read this DOCX. Upload an unencrypted Word .docx file.") from exc


def write_package(source: bytes, root: etree._Element) -> bytes:
    """Keep every package part and relationship unchanged except document.xml.

    All cloned nodes stay in the same document part, so their existing rIds remain
    valid. python-docx.save is deliberately not used for source exam packages.
    """
    output = BytesIO()
    replacement = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with ZipFile(BytesIO(source)) as original, ZipFile(output, "w", ZIP_DEFLATED) as result:
        result.comment = original.comment
        for info in original.infolist():
            result.writestr(info, replacement if info.filename == "word/document.xml"
                            else original.read(info))
    return output.getvalue()


def text_atoms(element: etree._Element):
    for node in element.iter():
        if node.tag == W + "t":
            yield node, node.text or ""
        elif node.tag == W + "tab":
            yield node, "\t"
        elif node.tag in {W + "br", W + "cr"}:
            yield node, "\n"


def block_text(element: etree._Element) -> str:
    return "".join(text for _, text in text_atoms(element))


def replace_text_span(element: etree._Element, start: int, end: int, replacement: str) -> None:
    """Replace a character span across runs, leaving run properties untouched."""
    if start < 0 or end <= start:
        raise DocumentError("Invalid text replacement span.")
    cursor = 0
    remaining = replacement
    touched = []
    for node, text in text_atoms(element):
        next_cursor = cursor + len(text)
        lo, hi = max(start, cursor), min(end, next_cursor)
        if lo < hi:
            if node.tag != W + "t":
                raise DocumentError("A label overlaps a tab or break; cannot safely relabel it.")
            touched.append((node, text, lo - cursor, hi - cursor))
        cursor = next_cursor
    if not touched or end > cursor:
        raise DocumentError("Could not locate the complete question/option label.")
    for i, (node, text, lo, hi) in enumerate(touched):
        count = len(remaining) if i == len(touched) - 1 else min(hi - lo, len(remaining))
        node.text = text[:lo] + remaining[:count] + text[hi:]
        remaining = remaining[count:]
        if node.text[:1].isspace() or node.text[-1:].isspace():
            node.set(XML_SPACE, "preserve")


def split_paragraph_fragments(paragraph: etree._Element, cuts: list[int]) -> list[list[etree._Element]]:
    """Split simple inline run content at offsets without flattening its styles.

    Non-text run children (drawing, OMML, fields) are intentionally rejected for
    inline shuffling. They remain fully supported in whole-block movements.
    """
    runs = [child for child in paragraph if child.tag != W + "pPr"]
    if any(run.tag != W + "r" or any(c.tag not in {W + "rPr", W + "t", W + "tab"}
                                     for c in run) for run in runs):
        raise DocumentError("Inline choices contain complex XML; disable option shuffling.")
    bounds = [0, *cuts, len(block_text(paragraph))]
    fragments: list[list[etree._Element]] = [[] for _ in range(len(bounds) - 1)]
    cursor = 0
    for run in runs:
        for child in run:
            if child.tag == W + "rPr":
                continue
            text = (child.text or "") if child.tag == W + "t" else "\t"
            finish = cursor + len(text)
            for index, (lo, hi) in enumerate(zip(bounds, bounds[1:])):
                left, right = max(cursor, lo), min(finish, hi)
                if left < right:
                    clone = etree.Element(W + "r", attrib=dict(run.attrib))
                    properties = run.find(W + "rPr")
                    if properties is not None:
                        clone.append(deepcopy(properties))
                    piece = deepcopy(child)
                    if child.tag == W + "t":
                        piece.text = text[left - cursor:right - cursor]
                        piece.set(XML_SPACE, "preserve")
                    clone.append(piece)
                    fragments[index].append(clone)
            cursor = finish
    return fragments


def colored_text(element: etree._Element) -> tuple[str, ...]:
    colors = set()
    for node in element.iter(W + "color"):
        value = node.get(W + "val", "auto").upper()
        if value not in {"AUTO", "000000", "FFFFFF"}:
            colors.add(value)
        elif node.get(W + "themeColor"):
            colors.add("theme:" + node.get(W + "themeColor"))
    return tuple(sorted(colors))


def inline_fragments(paragraph: etree._Element, label_starts: list[int]):
    """Return stem, choice contents and fixed-position whitespace separators."""
    text = block_text(paragraph)
    cuts = []
    ends = label_starts[1:] + [len(text)]
    for start, end in zip(label_starts, ends):
        content_end = start + len(text[start:end].rstrip())
        cuts.extend((start, content_end))
    fragments = split_paragraph_fragments(paragraph, cuts)
    return fragments[0], fragments[1::2], fragments[2::2]


def canonical(element: etree._Element) -> bytes:
    return etree.tostring(element, method="c14n")
