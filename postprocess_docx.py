from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}

ET.register_namespace("w", W_NS)


def w_tag(name: str) -> str:
    return f"{W}{name}"


def w_attr(name: str) -> str:
    return f"{W}{name}"


def child(parent: ET.Element, name: str) -> ET.Element | None:
    return parent.find(f"w:{name}", NS)


def ensure_child(parent: ET.Element, name: str) -> ET.Element:
    found = child(parent, name)
    if found is not None:
        return found
    found = ET.SubElement(parent, w_tag(name))
    return found


def remove_children(parent: ET.Element, name: str) -> None:
    tag = w_tag(name)
    for item in list(parent):
        if item.tag == tag:
            parent.remove(item)


def set_attr(element: ET.Element, name: str, value: str) -> None:
    element.set(w_attr(name), value)


def add_element(parent: ET.Element, name: str, attrs: dict[str, str] | None = None) -> ET.Element:
    element = ET.SubElement(parent, w_tag(name))
    for key, value in (attrs or {}).items():
        set_attr(element, key, value)
    return element


def paragraph_style_id(paragraph: ET.Element) -> str | None:
    p_pr = child(paragraph, "pPr")
    if p_pr is None:
        return None
    p_style = child(p_pr, "pStyle")
    if p_style is None:
        return None
    return p_style.get(w_attr("val"))


def set_paragraph_style_id(paragraph: ET.Element, style_id: str) -> None:
    p_pr = ensure_child(paragraph, "pPr")
    p_style = child(p_pr, "pStyle")
    if p_style is None:
        p_style = ET.Element(w_tag("pStyle"))
        p_pr.insert(0, p_style)
    set_attr(p_style, "val", style_id)


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(t.text or "" for t in paragraph.findall(".//w:t", NS)).strip()


def set_paragraph_text(paragraph: ET.Element, value: str) -> None:
    texts = paragraph.findall(".//w:t", NS)
    if not texts:
        run = add_element(paragraph, "r")
        text = add_element(run, "t")
        text.text = value
        return
    texts[0].text = value
    for text in texts[1:]:
        text.text = ""


def is_page_break_paragraph(paragraph: ET.Element) -> bool:
    br = paragraph.find(".//w:br", NS)
    return br is not None and br.get(w_attr("type")) == "page"


def make_page_break_paragraph() -> ET.Element:
    paragraph = ET.Element(w_tag("p"))
    run = add_element(paragraph, "r")
    add_element(run, "br", {"type": "page"})
    return paragraph


def paragraph_needs_leading_page_break(paragraph: ET.Element) -> bool:
    style_id = paragraph_style_id(paragraph)
    text = paragraph_text(paragraph)
    if style_id == "Heading1":
        return True
    if style_id == "AbstractTitle" and text in {"摘要", "目录"}:
        return True
    return False


def insert_explicit_page_breaks(parent: ET.Element) -> None:
    children = list(parent)
    offset = 0
    for index, item in enumerate(children):
        actual_index = index + offset
        if item.tag == w_tag("p") and paragraph_needs_leading_page_break(item):
            previous = parent[actual_index - 1] if actual_index > 0 else None
            if previous is not None and previous.tag == w_tag("p") and is_page_break_paragraph(previous):
                continue
            parent.insert(actual_index, make_page_break_paragraph())
            offset += 1
            continue
        insert_explicit_page_breaks(item)


def heading_level(style_id: str | None) -> int | None:
    if not style_id or not style_id.startswith("Heading"):
        return None
    raw = style_id.removeprefix("Heading")
    if raw.isdigit():
        level = int(raw)
        if 1 <= level <= 6:
            return level
    return None


def heading_prefix(level: int, counters: list[int]) -> str:
    if level == 1:
        return f"第{counters[0]}章 "
    if level == 2:
        return f"{counters[0]}.{counters[1]} "
    if level == 3:
        return f"{counters[0]}.{counters[1]}.{counters[2]} "
    if level == 4:
        return f"({counters[3]}) "
    if level == 5:
        return f"{counters[4]}) "
    letter = chr(ord("a") + max(counters[5] - 1, 0) % 26)
    return f"{letter}) "


def heading_already_prefixed(level: int, text: str, prefix: str) -> bool:
    if text.startswith(prefix):
        return True
    if level == 1 and text.startswith("第") and "章" in text[:8]:
        return True
    return False


def insert_text_prefix(paragraph: ET.Element, value: str) -> None:
    run = ET.Element(w_tag("r"))
    text = add_element(run, "t")
    text.text = value
    set_attr(text, "space", "preserve")

    insert_at = 0
    for index, item in enumerate(list(paragraph)):
        if item.tag == w_tag("pPr"):
            insert_at = index + 1
            break
    paragraph.insert(insert_at, run)


def apply_static_heading_numbers(root: ET.Element) -> None:
    counters = [0, 0, 0, 0, 0, 0]
    for paragraph in root.findall(".//w:p", NS):
        level = heading_level(paragraph_style_id(paragraph))
        if level is None:
            continue

        counters[level - 1] += 1
        for index in range(level, len(counters)):
            counters[index] = 0

        prefix = heading_prefix(level, counters)
        if not heading_already_prefixed(level, paragraph_text(paragraph), prefix):
            insert_text_prefix(paragraph, prefix)


def max_bookmark_id(root: ET.Element) -> int:
    values: list[int] = []
    for bookmark in root.findall(".//w:bookmarkStart", NS):
        raw = bookmark.get(w_attr("id"))
        if raw and raw.isdigit():
            values.append(int(raw))
    return max(values, default=0)


def add_bookmark_start(name: str, bookmark_id: int) -> ET.Element:
    bookmark = ET.Element(w_tag("bookmarkStart"))
    set_attr(bookmark, "id", str(bookmark_id))
    set_attr(bookmark, "name", name)
    return bookmark


def add_bookmark_end(bookmark_id: int) -> ET.Element:
    bookmark = ET.Element(w_tag("bookmarkEnd"))
    set_attr(bookmark, "id", str(bookmark_id))
    return bookmark


def add_run_text(paragraph: ET.Element, value: str, *, preserve: bool = False) -> ET.Element:
    run = add_element(paragraph, "r")
    text = add_element(run, "t")
    if preserve:
        set_attr(text, "space", "preserve")
    text.text = value
    return run


def add_run_tab(paragraph: ET.Element) -> None:
    run = add_element(paragraph, "r")
    add_element(run, "tab")


def add_pageref_field(paragraph: ET.Element, bookmark_name: str) -> None:
    begin_run = add_element(paragraph, "r")
    add_element(begin_run, "fldChar", {"fldCharType": "begin", "dirty": "true"})

    instr_run = add_element(paragraph, "r")
    instr = add_element(instr_run, "instrText")
    set_attr(instr, "space", "preserve")
    instr.text = f" PAGEREF {bookmark_name} \\h "

    separate_run = add_element(paragraph, "r")
    add_element(separate_run, "fldChar", {"fldCharType": "separate"})

    add_run_text(paragraph, "1")

    end_run = add_element(paragraph, "r")
    add_element(end_run, "fldChar", {"fldCharType": "end"})


def make_toc_entry(text: str, level: int, bookmark_name: str) -> ET.Element:
    paragraph = ET.Element(w_tag("p"))
    p_pr = add_element(paragraph, "pPr")
    add_element(p_pr, "pStyle", {"val": f"TOC{level}"})
    add_run_text(paragraph, text)
    add_run_tab(paragraph)
    add_pageref_field(paragraph, bookmark_name)
    return paragraph


def make_title_paragraph(text: str) -> ET.Element:
    paragraph = ET.Element(w_tag("p"))
    p_pr = add_element(paragraph, "pPr")
    add_element(p_pr, "pStyle", {"val": "AbstractTitle"})
    add_run_text(paragraph, text)
    return paragraph


def add_heading_bookmarks_and_collect_toc(root: ET.Element) -> list[tuple[int, str, str]]:
    entries: list[tuple[int, str, str]] = []
    next_id = max_bookmark_id(root) + 1

    def walk(parent: ET.Element) -> None:
        nonlocal next_id
        index = 0
        while index < len(parent):
            item = parent[index]
            if item.tag == w_tag("p"):
                level = heading_level(paragraph_style_id(item))
                if level is not None and level <= 3:
                    bookmark_name = f"_mdtoc_{len(entries) + 1}"
                    parent.insert(index, add_bookmark_start(bookmark_name, next_id))
                    parent.insert(index + 2, add_bookmark_end(next_id))
                    entries.append((level, paragraph_text(item), bookmark_name))
                    next_id += 1
                    index += 3
                    continue
            walk(item)
            index += 1

    walk(root)
    return entries


def replace_toc_with_static_entries(root: ET.Element, entries: list[tuple[int, str, str]]) -> None:
    for sdt in root.findall(".//w:sdt", NS):
        gallery = sdt.find("w:sdtPr/w:docPartObj/w:docPartGallery", NS)
        if gallery is None or gallery.get(w_attr("val")) != "Table of Contents":
            continue

        content = child(sdt, "sdtContent")
        if content is None:
            content = add_element(sdt, "sdtContent")
        for item in list(content):
            content.remove(item)

        content.append(make_page_break_paragraph())
        content.append(make_title_paragraph("目录"))
        for level, text, bookmark_name in entries:
            content.append(make_toc_entry(text, level, bookmark_name))
        return


def patch_document_xml(path: Path) -> None:
    tree = ET.parse(path)
    root = tree.getroot()

    for paragraph in root.findall(".//w:p", NS):
        style_id = paragraph_style_id(paragraph)
        text = paragraph_text(paragraph)

        if style_id == "AbstractTitle" and text in {"Abstract", "摘要"}:
            set_paragraph_text(paragraph, "摘要")

        if style_id == "TOCHeading" and text in {"Table of Contents", "Contents", "目录"}:
            set_paragraph_text(paragraph, "目录")
            # Reuse the front-matter title style so the TOC title starts a new page
            # but does not participate in chapter numbering.
            set_paragraph_style_id(paragraph, "AbstractTitle")

    body = root.find("w:body", NS)
    if body is not None:
        apply_static_heading_numbers(body)
        insert_explicit_page_breaks(body)

    tree.write(path, encoding="UTF-8", xml_declaration=True)


def ensure_style_ppr(style: ET.Element) -> ET.Element:
    p_pr = child(style, "pPr")
    if p_pr is not None:
        return p_pr
    p_pr = ET.Element(w_tag("pPr"))
    # Keep pPr near the top of the style definition for better Word compatibility.
    insert_at = 0
    for i, item in enumerate(list(style)):
        if item.tag in {w_tag("name"), w_tag("basedOn"), w_tag("next"), w_tag("link"), w_tag("uiPriority")}:
            insert_at = i + 1
    style.insert(insert_at, p_pr)
    return p_pr


def set_style_page_break(style: ET.Element, enabled: bool = True) -> None:
    p_pr = ensure_style_ppr(style)
    remove_children(p_pr, "pageBreakBefore")
    if enabled:
        add_element(p_pr, "pageBreakBefore")


def remove_style_pagination_marks(style: ET.Element) -> None:
    p_pr = child(style, "pPr")
    if p_pr is None:
        return
    for name in ("pageBreakBefore", "keepNext", "keepLines"):
        remove_children(p_pr, name)


def set_style_alignment(style: ET.Element, alignment: str) -> None:
    p_pr = ensure_style_ppr(style)
    remove_children(p_pr, "jc")
    add_element(p_pr, "jc", {"val": alignment})


def set_style_spacing(style: ET.Element, *, before: str, after: str, line: str, rule: str = "auto") -> None:
    p_pr = ensure_style_ppr(style)
    remove_children(p_pr, "spacing")
    add_element(p_pr, "spacing", {"before": before, "after": after, "line": line, "lineRule": rule})


def set_style_indent(style: ET.Element, *, left: str = "0", first_line: str | None = "0", hanging: str | None = None) -> None:
    p_pr = ensure_style_ppr(style)
    remove_children(p_pr, "ind")
    attrs = {"left": left}
    if first_line is not None:
        attrs["firstLine"] = first_line
    if hanging is not None:
        attrs["hanging"] = hanging
    add_element(p_pr, "ind", attrs)


def set_style_right_tab(style: ET.Element, *, pos: str = "8312", leader: str = "dot") -> None:
    p_pr = ensure_style_ppr(style)
    remove_children(p_pr, "tabs")
    tabs = add_element(p_pr, "tabs")
    add_element(tabs, "tab", {"val": "right", "leader": leader, "pos": pos})


def set_auto_spacing_off(p_pr: ET.Element) -> None:
    remove_children(p_pr, "autoSpaceDE")
    remove_children(p_pr, "autoSpaceDN")
    add_element(p_pr, "autoSpaceDE", {"val": "0"})
    add_element(p_pr, "autoSpaceDN", {"val": "0"})


def set_based_on(style: ET.Element, style_id: str) -> None:
    remove_children(style, "basedOn")
    based_on = ET.Element(w_tag("basedOn"))
    set_attr(based_on, "val", style_id)
    style.insert(1, based_on)


def ensure_paragraph_style(root: ET.Element, style_id: str, name: str) -> ET.Element:
    style = root.find(f"w:style[@w:styleId='{style_id}']", NS)
    if style is not None:
        return style
    style = ET.Element(w_tag("style"), {w_attr("type"): "paragraph", w_attr("styleId"): style_id})
    add_element(style, "name", {"val": name})
    add_element(style, "basedOn", {"val": "Normal"})
    add_element(style, "uiPriority", {"val": "39"})
    add_element(style, "unhideWhenUsed")
    root.append(style)
    return style


def set_style_run_font(style: ET.Element, *, latin: str, east_asia: str, size: str = "24", bold: bool = False) -> None:
    r_pr = child(style, "rPr")
    if r_pr is None:
        r_pr = add_element(style, "rPr")
    remove_children(r_pr, "rFonts")
    fonts = add_element(r_pr, "rFonts")
    for key, value in {
        "ascii": latin,
        "hAnsi": latin,
        "cs": latin,
        "eastAsia": east_asia,
    }.items():
        set_attr(fonts, key, value)
    remove_children(r_pr, "sz")
    remove_children(r_pr, "szCs")
    add_element(r_pr, "sz", {"val": size})
    add_element(r_pr, "szCs", {"val": size})
    remove_children(r_pr, "b")
    if bold:
        add_element(r_pr, "b")


def configure_toc_styles(root: ET.Element) -> None:
    # A4 width 11906 twips - left/right margins 1797 twips = 8312 twips text width.
    # The right tab gives Word the classic "title .... page" TOC layout.
    for level in range(1, 7):
        style = ensure_paragraph_style(root, f"TOC{level}", f"toc {level}")
        p_pr = ensure_style_ppr(style)
        remove_children(p_pr, "numPr")
        set_style_indent(style, left=str((level - 1) * 420), first_line="0")
        set_style_spacing(style, before="0", after="0", line="240", rule="auto")
        set_style_right_tab(style)
        set_style_run_font(style, latin="Times New Roman", east_asia="SimSun", size="24")


def patch_styles_xml(path: Path, num_id: str) -> None:
    tree = ET.parse(path)
    root = tree.getroot()

    for style in root.findall("w:style", NS):
        remove_style_pagination_marks(style)
        p_pr = child(style, "pPr")
        if p_pr is not None:
            remove_children(p_pr, "numPr")

    styles_by_id = {
        style.get(w_attr("styleId")): style
        for style in root.findall("w:style", NS)
        if style.get(w_attr("styleId"))
    }

    configure_toc_styles(root)

    abstract_title = styles_by_id.get("AbstractTitle")
    if abstract_title is not None:
        set_style_alignment(abstract_title, "center")
        set_style_spacing(abstract_title, before="480", after="360", line="240", rule="auto")

    toc_heading = styles_by_id.get("TOCHeading")
    if toc_heading is not None:
        set_based_on(toc_heading, "Normal")
        set_style_alignment(toc_heading, "center")
        set_style_spacing(toc_heading, before="480", after="360", line="240", rule="auto")

    for level in range(1, 7):
        style = styles_by_id.get(f"Heading{level}")
        if style is None:
            continue
        p_pr = ensure_style_ppr(style)
        set_auto_spacing_off(p_pr)
        remove_children(p_pr, "numPr")
        if level == 1:
            set_style_alignment(style, "center")

    tree.write(path, encoding="UTF-8", xml_declaration=True)


def max_id(root: ET.Element, element_name: str, attr_name: str) -> int:
    values: list[int] = []
    for element in root.findall(f"w:{element_name}", NS):
        raw = element.get(w_attr(attr_name))
        if raw and raw.isdigit():
            values.append(int(raw))
    return max(values, default=0)


def add_level(
    abstract: ET.Element,
    *,
    ilvl: int,
    fmt: str,
    text: str,
    style: str,
    align: str,
    size: str,
) -> None:
    lvl = add_element(abstract, "lvl", {"ilvl": str(ilvl)})
    add_element(lvl, "start", {"val": "1"})
    add_element(lvl, "numFmt", {"val": fmt})
    add_element(lvl, "pStyle", {"val": style})
    add_element(lvl, "lvlText", {"val": text})
    add_element(lvl, "suff", {"val": "space"})
    add_element(lvl, "lvlJc", {"val": align})

    p_pr = add_element(lvl, "pPr")
    set_auto_spacing_off(p_pr)
    add_element(p_pr, "ind", {"left": "0", "hanging": "0"})

    r_pr = add_element(lvl, "rPr")
    add_element(r_pr, "b")
    fonts = add_element(r_pr, "rFonts")
    for key in ("ascii", "hAnsi", "cs", "eastAsia"):
        set_attr(fonts, key, "SimHei")
    add_element(r_pr, "sz", {"val": size})
    add_element(r_pr, "szCs", {"val": size})


def patch_numbering_xml(path: Path) -> str:
    if path.exists():
        tree = ET.parse(path)
        root = tree.getroot()
    else:
        root = ET.Element(w_tag("numbering"))
        tree = ET.ElementTree(root)

    abstract_id = str(max_id(root, "abstractNum", "abstractNumId") + 1)
    num_id = str(max_id(root, "num", "numId") + 1)

    abstract = add_element(root, "abstractNum", {"abstractNumId": abstract_id})
    add_element(abstract, "nsid", {"val": "4D445750"})
    add_element(abstract, "multiLevelType", {"val": "hybridMultilevel"})
    add_element(abstract, "tmpl", {"val": "4D445750"})

    levels = [
        ("decimal", "第%1章", "Heading1", "center", "32"),
        ("decimal", "%1.%2", "Heading2", "left", "28"),
        ("decimal", "%1.%2.%3", "Heading3", "left", "26"),
        ("decimal", "(%4)", "Heading4", "left", "24"),
        ("decimal", "%5)", "Heading5", "left", "24"),
        ("lowerLetter", "%6)", "Heading6", "left", "24"),
    ]
    for ilvl, (fmt, text, style, align, size) in enumerate(levels):
        add_level(abstract, ilvl=ilvl, fmt=fmt, text=text, style=style, align=align, size=size)

    num = add_element(root, "num", {"numId": num_id})
    add_element(num, "abstractNumId", {"val": abstract_id})

    tree.write(path, encoding="UTF-8", xml_declaration=True)
    return num_id


def patch_settings_xml(path: Path) -> None:
    tree = ET.parse(path)
    root = tree.getroot()

    update_fields = child(root, "updateFields")
    if update_fields is None:
        update_fields = add_element(root, "updateFields")
    set_attr(update_fields, "val", "true")

    tree.write(path, encoding="UTF-8", xml_declaration=True)


def patch_docx(input_docx: Path, output_docx: Path) -> None:
    input_docx = input_docx.resolve()
    output_docx = output_docx.resolve()

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        with zipfile.ZipFile(input_docx) as zin:
            zin.extractall(tmp_path)

        document_xml = tmp_path / "word" / "document.xml"
        styles_xml = tmp_path / "word" / "styles.xml"
        numbering_xml = tmp_path / "word" / "numbering.xml"
        settings_xml = tmp_path / "word" / "settings.xml"

        num_id = patch_numbering_xml(numbering_xml)
        if styles_xml.exists():
            patch_styles_xml(styles_xml, num_id)
        if document_xml.exists():
            patch_document_xml(document_xml)
        if settings_xml.exists():
            patch_settings_xml(settings_xml)

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in tmp_path.rglob("*"):
                if item.is_file():
                    zout.write(item, item.relative_to(tmp_path))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Apply thesis-format DOCX fixes after Pandoc conversion.")
    parser.add_argument("input", help="Input DOCX")
    parser.add_argument("output", nargs="?", help="Output DOCX; defaults to overwriting input")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_docx = Path(args.input)
    output_docx = Path(args.output) if args.output else input_docx
    if not input_docx.exists():
        raise SystemExit(f"Input DOCX not found: {input_docx}")

    if input_docx.resolve() == output_docx.resolve():
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as fh:
            temp_output = Path(fh.name)
        try:
            patch_docx(input_docx, temp_output)
            shutil.move(str(temp_output), str(output_docx))
        finally:
            temp_output.unlink(missing_ok=True)
    else:
        patch_docx(input_docx, output_docx)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
