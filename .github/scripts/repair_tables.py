"""Restore verified full-table PowerPoint fallback snapshots for LibreOffice.

Only the 18 explicitly reviewed tables in these original TX source decks are
eligible. Source bytes, media hashes, shape IDs, and geometry must match.
Non-table AlternateContent and all other ZIP members remain untouched.

The source snapshots already contain the complete correctly rendered tables.
Replacing each entire table AlternateContent with its snapshot avoids both
unsupported cell image-fill geometry and equation/text reflow. The snapshot
remains at the original table's position and size. It is raster content, so
searchable source text must be retained separately by the conversion pipeline.

Dependencies: lxml and Pillow. No files are written by repair_pptx().
"""
from copy import deepcopy
from io import BytesIO
import hashlib
import posixpath
import re
import zipfile

from lxml import etree as ET
from PIL import Image

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

# Visually verified against the original full-table PNGs on 2026-11-01.
# The explicit image and frame dimensions admit the source's own padding;
# aspect similarity alone is intentionally insufficient for eligibility.
VERIFIED = {'14-C2(c)~(e) C3.pptx': {'sha256': 'e175446b350e2502c02d51d3854e86a0117ace806ec035b433ab82e3d473e4af',
                          'tables': [{'slide': 20,
                                      'shape_id': '3',
                                      'media': 'ppt/media/image19.png',
                                      'image_sha256': '9f19b56ac0a4223d6b30f4b9ff895509ed21959eb7923c7630e61b9d5b3593e6',
                                      'image_size': [1246, 452],
                                      'frame_size': [7899400, 2856230]},
                                     {'slide': 7,
                                      'shape_id': '6',
                                      'media': 'ppt/media/image14.png',
                                      'image_sha256': '20cae2aa08ff4264116a9ae7393cffbfc53dc35b376a18947eaf3e7daf85b478',
                                      'image_size': [948, 377],
                                      'frame_size': [6003290, 2380615]}]},
 '15-C4.pptx': {'sha256': '6f1646c15e7cdb9379f38ab0dc06ea4ac7f7359e0de5be97a8612d718f7aeca5',
                'tables': [{'slide': 18,
                            'shape_id': '7',
                            'media': 'ppt/media/image15.png',
                            'image_sha256': '0acacaddb3015efd84f352297b04cd0628beff75cb0576b164bef424fe1fcc0d',
                            'image_size': [1048, 591],
                            'frame_size': [6636385, 3736340]}]},
 '17-A4~A6-individuals.pptx': {'sha256': '736a454c953b866e197bc4a642b0ace808b1d783eae42791d10ca751fb6d5acc',
                               'tables': [{'slide': 7,
                                           'shape_id': '3',
                                           'media': 'ppt/media/image15.png',
                                           'image_sha256': 'f424b118d019378e504e3fd6234b932ef489fd9fafa612a7ea2288abc48919ee',
                                           'image_size': [1133, 602],
                                           'frame_size': [7179945, 3810000]}]},
 '答案11 B3(g) .pptx': {'sha256': 'eeb15e66135b04068d53be410643d2fef866b719990b6d0bb81bca1db848b0d2',
                      'tables': [{'slide': 2,
                                  'shape_id': '6',
                                  'media': 'ppt/media/image15.png',
                                  'image_sha256': '46a5c37cc890d1a12bc02da24cf9c0c3ab4dedda3ac71f955aad8ee9df75f370',
                                  'image_size': [1696, 820],
                                  'frame_size': [8064500, 3888105]},
                                 {'slide': 3,
                                  'shape_id': '3',
                                  'media': 'ppt/media/image16.png',
                                  'image_sha256': 'e437ec0b3676c8f0ad2f108f982acf03f8b0a5ee2779889f1c1162d7b930e58e',
                                  'image_size': [1666, 617],
                                  'frame_size': [7920355, 2922270]}]},
 '答案14 C2(c)~(e) C3.pptx': {'sha256': 'ffaf0f360e6f1f30bbd921d24426dfbb44648fef468e7b9fd89271c3d3c22840',
                            'tables': [{'slide': 2,
                                        'shape_id': '6',
                                        'media': 'ppt/media/image16.png',
                                        'image_sha256': '9634ffad706adb19238bb4223a23deaa36bd785a0ba979cfcbe3ee943b6af9c7',
                                        'image_size': [504, 163],
                                        'frame_size': [6377940, 2045335]},
                                       {'slide': 3,
                                        'shape_id': '3',
                                        'media': 'ppt/media/image18.png',
                                        'image_sha256': 'c4cb71f8b3196a88b34034b0955ebb41b7fbfd82d86f946c0cf321b4a39ba3ee',
                                        'image_size': [417, 324],
                                        'frame_size': [5275580, 4096385]},
                                       {'slide': 5,
                                        'shape_id': '3',
                                        'media': 'ppt/media/image20.png',
                                        'image_sha256': '8d5212e52b96eb087e1562613ae5bc53d0413294a4536515eddeefc0c03d4100',
                                        'image_size': [602, 310],
                                        'frame_size': [7632700, 3918585]},
                                       {'slide': 6,
                                        'shape_id': '6',
                                        'media': 'ppt/media/image22.png',
                                        'image_sha256': 'acb4274bcfd3fdb7df19dc5d72178b7b6a49bd06351546c43b5f009a9d06e7b8',
                                        'image_size': [461, 136],
                                        'frame_size': [5832475, 1709420]}]},
 '答案15 C4.pptx': {'sha256': 'd60d627697ae196049f39021923c60aed39e73e1226a299f168b5c365665292b',
                  'tables': [{'slide': 5,
                              'shape_id': '3',
                              'media': 'ppt/media/image18.png',
                              'image_sha256': '9ff95079d83e0c71794eda07a7b1d8b159bf9e79662ae8e7332b53ed461f0d04',
                              'image_size': [558, 115],
                              'frame_size': [7066280, 1440815]}]},
 '答案16 C5 C6.pptx': {'sha256': '27e829b0dd3fea449387ac007763d22d6ee244099ec64bc8eec7483345ecfd2b',
                     'tables': [{'slide': 2,
                                 'shape_id': '6',
                                 'media': 'ppt/media/image16.png',
                                 'image_sha256': 'a247e3063081c20f7fece4aeb8c21d094357a81cb53eb8a611630238c218d36e',
                                 'image_size': [1468, 712],
                                 'frame_size': [8924349, 4187825]},
                                {'slide': 3,
                                 'shape_id': '3',
                                 'media': 'ppt/media/image18.png',
                                 'image_sha256': '41504c0cb7e2eaebfce5531bfab367075d2d8f3ad0c83ea30645bf0e74de5da4',
                                 'image_size': [615, 214],
                                 'frame_size': [5842635, 2023745]},
                                {'slide': 7,
                                 'shape_id': '3',
                                 'media': 'ppt/media/image20.png',
                                 'image_sha256': '26c0500561d37deae1c3e2035e5801f8e86c794a046865e813ccf3837b34f4b6',
                                 'image_size': [1197, 629],
                                 'frame_size': [7272655, 3642360]}]},
 '答案20 E2(d)~(g).pptx': {'sha256': '6feaa317fc1296852fd7b7cd3299348ca56018332fe2c14551aa8d225ced4e6f',
                         'tables': [{'slide': 2,
                                     'shape_id': '6',
                                     'media': 'ppt/media/image15.png',
                                     'image_sha256': '2b7853709fd4c0be178e7ef510b363066c2184803748197fc34c693dd586887c',
                                     'image_size': [1291, 691],
                                     'frame_size': [7848600, 3997754]},
                                    {'slide': 6,
                                     'shape_id': '3',
                                     'media': 'ppt/media/image16.png',
                                     'image_sha256': 'c0f49b492d776479a47dbc0cfd751e9e617e6094e7ce1f58bf0f012c73c2f4a7',
                                     'image_size': [1303, 724],
                                     'frame_size': [7920355, 4307840]}]},
 '答案24 A4~A6-companies.pptx': {'sha256': '6b8edbe5ffefda36721d50acba299bf85859513887fb68f8d36c1f7016174375',
                               'tables': [{'slide': 3,
                                           'shape_id': '3',
                                           'media': 'ppt/media/image17.png',
                                           'image_sha256': '1f2165dd8b620eccd364bb4d3fee4048c72ad9b87275f3998d18347c2998a6bb',
                                           'image_size': [602, 243],
                                           'frame_size': [7632700, 3070860]}]},
 '答案25 F.pptx': {'sha256': 'c280a8114c7320f6454d3653ffe1a4d4597bdc0ea40f99e827d8a30c0cfc3851',
                 'tables': [{'slide': 4,
                             'shape_id': '4',
                             'media': 'ppt/media/image17.png',
                             'image_sha256': 'a4ed99ac9ce022bb445b7238dddf00a5b7f56504c24b6b90ece64cd10beeead9',
                             'image_size': [1214, 477],
                           'frame_size': [7375515, 2719303]}]}}


def _tag(prefix, name):
    return f"{{{NS[prefix]}}}{name}"


def _require(condition, message):
    if not condition:
        raise ValueError(message)


def _image_picture(frame, rid):
    """Preserve identity and bounds while using the existing image relation."""
    picture = ET.Element(_tag("p", "pic"))
    nonvisual = ET.SubElement(picture, _tag("p", "nvPicPr"))
    original = frame.find("p:nvGraphicFramePr", NS)
    nonvisual.append(deepcopy(original.find("p:cNvPr", NS)))
    locks = ET.SubElement(ET.SubElement(nonvisual, _tag("p", "cNvPicPr")),
                          _tag("a", "picLocks"))
    locks.set("noChangeAspect", "1")
    nonvisual.append(deepcopy(original.find("p:nvPr", NS)))

    fill = ET.SubElement(picture, _tag("p", "blipFill"))
    ET.SubElement(fill, _tag("a", "blip"), {_tag("r", "embed"): rid})
    ET.SubElement(ET.SubElement(fill, _tag("a", "stretch")), _tag("a", "fillRect"))
    properties = ET.SubElement(picture, _tag("p", "spPr"))
    transform = deepcopy(frame.find("p:xfrm", NS))
    transform.tag = _tag("a", "xfrm")
    properties.append(transform)
    ET.SubElement(ET.SubElement(properties, _tag("a", "prstGeom"), prst="rect"),
                  _tag("a", "avLst"))
    ET.SubElement(ET.SubElement(properties, _tag("a", "ln")), _tag("a", "noFill"))
    return picture


def repair_pptx(data: bytes, filename: str) -> tuple[bytes, list[dict]]:
    """Return repaired ZIP bytes and a repair log; reject unverified candidates.

    Unaffected inputs return their original bytes. Repaired inputs are also
    idempotent: with the table AlternateContent removed, no candidates remain.
    This function never installs, renders, writes files, or accesses a network.
    """
    parser = ET.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    roots = {}
    candidates = []
    with zipfile.ZipFile(BytesIO(data)) as package:
        for part in package.namelist():
            match = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", part)
            if not match:
                continue
            root = ET.fromstring(package.read(part), parser)
            for alternate in root.findall(".//mc:AlternateContent", NS):
                fallback = alternate.find("mc:Fallback", NS)
                if fallback is None:
                    continue
                for frame in fallback.findall("p:graphicFrame", NS):
                    table = frame.find("a:graphic/a:graphicData/a:tbl", NS)
                    if table is None:
                        continue
                    fills = table.findall(".//a:tcPr/a:blipFill", NS)
                    if not fills:
                        continue
                    roots[part] = root
                    candidates.append((part, int(match.group(1)), alternate, fallback, frame, fills))
        if not candidates:
            return data, []

        verified = VERIFIED.get(filename)
        _require(verified is not None, f"Unverified table fallback in {filename!r}")
        _require(hashlib.sha256(data).hexdigest() == verified["sha256"],
                 f"Original PPTX hash mismatch: {filename}")
        expected = {(t["slide"], t["shape_id"]): t for t in verified["tables"]}
        seen = set()
        repairs = []
        for part, slide_number, alternate, fallback, frame, fills in candidates:
            identity = frame.find("p:nvGraphicFramePr/p:cNvPr", NS)
            key = (slide_number, identity.get("id"))
            _require(key in expected and key not in seen,
                     f"Unexpected/duplicate table {filename}:{key}")
            target = expected[key]
            _require(alternate.getparent().tag == _tag("p", "spTree"),
                     f"Unsupported nested table: {filename}:{key}")
            _require(len(fallback) == 1 and fallback[0] is frame,
                     f"Fallback contains extra objects: {filename}:{key}")
            _require(len(alternate.findall("mc:Choice", NS)) == 1,
                     f"Ambiguous table choices: {filename}:{key}")
            choice_frame = alternate.find("mc:Choice/p:graphicFrame", NS)
            _require(choice_frame is not None and
                     choice_frame.find("p:nvGraphicFramePr/p:cNvPr", NS).get("id") == key[1],
                     f"Choice/fallback identity mismatch: {filename}:{key}")
            extent = frame.find("p:xfrm/a:ext", NS)
            frame_size = [int(extent.get(n)) for n in ("cx", "cy")]
            _require(frame_size == target["frame_size"],
                     f"Table dimensions changed: {filename}:{key}")

            rel_part = posixpath.join(posixpath.dirname(part), "_rels",
                                     posixpath.basename(part) + ".rels")
            relationships = {r.get("Id"): r for r in
                             ET.fromstring(package.read(rel_part), parser)}
            rids = {fill.find("a:blip", NS).get(_tag("r", "embed")) for fill in fills}
            _require(len(rids) == 1 and None not in rids,
                     f"Table does not have one embedded image: {filename}:{key}")
            rid = next(iter(rids))
            relationship = relationships[rid]
            _require(relationship.get("TargetMode", "Internal") == "Internal" and
                     relationship.get("Type", "").endswith("/image"),
                     f"Not an internal image relation: {filename}:{key}")
            resource = posixpath.normpath(posixpath.join(posixpath.dirname(part),
                                                        relationship.get("Target")))
            _require(resource == target["media"],
                     f"Unexpected image resource: {filename}:{key}")
            image_data = package.read(resource)
            _require(hashlib.sha256(image_data).hexdigest() == target["image_sha256"],
                     f"Original table image hash mismatch: {filename}:{key}")
            with Image.open(BytesIO(image_data)) as image:
                image_size = list(image.size)
                _require(image.format == "PNG" and image_size == target["image_size"],
                         f"Unexpected table image format/dimensions: {filename}:{key}")

            # Replacing the entire AlternateContent makes LibreOffice's choice
            # deterministic. Unrelated choices, shapes and ink remain intact.
            alternate.getparent().replace(alternate, _image_picture(frame, rid))
            seen.add(key)
            repairs.append({
                "slide_number": slide_number,
                "shape_id": key[1],
                "reason": "Use verified original full-table PNG to avoid unsupported table-cell fallback fills",
                "image_resource": resource,
                "image_sha256": target["image_sha256"],
                "image_size": image_size,
                "frame_size_emu": frame_size,
                "image_to_frame_aspect_ratio": round(
                    (image_size[0] / image_size[1]) / (frame_size[0] / frame_size[1]), 6),
                "replaced_cell_image_fills": len(fills),
                "table_text_is_raster": True,
            })
        _require(seen == set(expected), f"Missing verified tables: {filename}: {set(expected) - seen}")

        changed = {part: ET.tostring(root, xml_declaration=True, encoding="UTF-8",
                                    standalone=True) for part, root in roots.items()}
        output = BytesIO()
        with zipfile.ZipFile(output, "w") as repaired:
            repaired.comment = package.comment
            for member in package.infolist():
                repaired.writestr(member, changed.get(member.filename, package.read(member.filename)))
        return output.getvalue(), sorted(repairs, key=lambda item: (item["slide_number"], item["shape_id"]))
