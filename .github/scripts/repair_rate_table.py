"""Apply one audited layout repair to the original TX corporation-tax deck.

Call this before any other PPTX repair: the complete original archive SHA-256
must match. The original archive is never modified on disk.
"""

from copy import deepcopy
from hashlib import sha256
from io import BytesIO
from pathlib import PurePosixPath
from zipfile import ZipFile

from lxml import etree


SOURCE_FILENAME = "22-E4.pptx"
SOURCE_SHA256 = "6a3b3c481722672738ab3625c03d6c80764fb4432e3fb1f1856931cfc4faabc1"
SLIDE_PART = "ppt/slides/slide6.xml"
SLIDE_SHA256 = "4da42a28d850d2d83f15791dfe5330cbd46b760c11e4b459a00cb3812ffb934f"
TARGET_TEXT = "Small profits rate（augmented profits≤lower limit）"
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def _a(local_name: str) -> str:
    return "{" + NS["a"] + "}" + local_name


def repair_rate_table(data: bytes, filename: str) -> tuple[bytes, list[dict]]:
    """Reduce only slide 6's long rate-label cell from 18 pt to 17 pt.

    All other filenames pass through unchanged. A matching filename with a
    different source hash fails closed, including previously repaired input.
    Text, line spacing, table geometry, and every other ZIP member are retained.
    """
    if PurePosixPath(filename.replace("\\", "/")).name != SOURCE_FILENAME:
        return data, []

    source_hash = sha256(data).hexdigest()
    if source_hash != SOURCE_SHA256:
        raise ValueError(
            f"Refusing rate-table repair: {filename} is not the audited original "
            f"(SHA-256 {source_hash}; expected {SOURCE_SHA256})"
        )

    with ZipFile(BytesIO(data)) as source:
        slide_xml = source.read(SLIDE_PART)
        if sha256(slide_xml).hexdigest() != SLIDE_SHA256:
            raise ValueError("Audited slide 6 XML hash does not match")

        parser = etree.XMLParser(resolve_entities=False, no_network=True)
        root = etree.fromstring(slide_xml, parser)
        before = deepcopy(root)
        frames = root.xpath(
            "./p:cSld/p:spTree/p:graphicFrame[p:nvGraphicFramePr/p:cNvPr[@id='6']]",
            namespaces=NS,
        )
        if len(frames) != 1:
            raise ValueError("Expected exactly one audited table shape with id 6")
        cells = frames[0].xpath(".//a:tbl/a:tr[2]/a:tc[1]", namespaces=NS)
        if len(cells) != 1:
            raise ValueError("Expected the audited second-row, first-column cell")
        cell = cells[0]
        paragraphs = cell.findall("a:txBody/a:p", NS)
        if len(paragraphs) != 1:
            raise ValueError("Audited rate cell must contain exactly one paragraph")
        paragraph = paragraphs[0]
        text = "".join(paragraph.itertext())
        if text != TARGET_TEXT or paragraph.findall("a:br", NS):
            raise ValueError("Audited rate-cell text or line-break structure changed")

        existing_properties = paragraph.xpath("./a:r/a:rPr | ./a:endParaRPr", namespaces=NS)
        if len(existing_properties) != 7 or any(
            prop.get("sz") != "1800" for prop in existing_properties
        ):
            raise ValueError("Audited rate cell must have seven explicit 18 pt properties")

        original_cell = deepcopy(cell)
        for prop in existing_properties:
            prop.set("sz", "1700")
        paragraph_properties = paragraph.find("a:pPr", NS)
        if paragraph_properties is None:
            paragraph_properties = etree.Element(_a("pPr"))
            paragraph.insert(0, paragraph_properties)
        default_properties = paragraph_properties.find("a:defRPr", NS)
        if default_properties is None:
            default_properties = etree.Element(_a("defRPr"))
            extension_list = paragraph_properties.find("a:extLst", NS)
            if extension_list is None:
                paragraph_properties.append(default_properties)
            else:
                paragraph_properties.insert(
                    paragraph_properties.index(extension_list), default_properties
                )
        default_properties.set("sz", "1700")

        # Verify the whole slide is unchanged outside the one audited cell.
        parent = cell.getparent()
        index = parent.index(cell)
        parent[index] = original_cell
        if etree.tostring(root, method="c14n") != etree.tostring(before, method="c14n"):
            raise ValueError("Rate-table repair changed content outside its target cell")
        parent[index] = cell
        if "".join(cell.itertext()) != "".join(original_cell.itertext()):
            raise ValueError("Rate-table repair changed the target cell's text")

        repaired_xml = etree.tostring(root, encoding="UTF-8", xml_declaration=True,
                                      standalone=True)
        output = BytesIO()
        with ZipFile(output, "w") as result:
            result.comment = source.comment
            for member in source.infolist():
                result.writestr(
                    member,
                    repaired_xml if member.filename == SLIDE_PART else source.read(member),
                )

    repaired = output.getvalue()
    return repaired, [{
        "repair": "corporation_rate_label_font_fit",
        "source": filename,
        "source_sha256": source_hash,
        "source_sha256_verified": True,
        "source_slide_sha256": SLIDE_SHA256,
        "repaired_sha256": sha256(repaired).hexdigest(),
        "repaired_slide_sha256": sha256(repaired_xml).hexdigest(),
        "slide": 6,
        "slide_part": SLIDE_PART,
        "shape_id": 6,
        "table_row": 2,
        "table_column": 1,
        "text": TARGET_TEXT,
        "old_font_pt": 18,
        "new_font_pt": 17,
        "updated_existing_font_properties": len(existing_properties),
        "paragraph_default_font_pt": 17,
        "text_spacing_and_geometry_preserved": True,
    }]
