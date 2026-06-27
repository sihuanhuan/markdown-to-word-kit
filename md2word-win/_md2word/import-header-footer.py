from __future__ import annotations

import argparse
import posixpath
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
REL = f"{{{REL_NS}}}"
CT = f"{{{CT_NS}}}"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("", REL_NS)

HEADER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
FOOTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"


def qn(namespace_prefix: str, name: str) -> str:
    if namespace_prefix == "w":
        return f"{W}{name}"
    if namespace_prefix == "r":
        return f"{R}{name}"
    raise ValueError(namespace_prefix)


def unzip_docx(path: Path, out_dir: Path) -> None:
    with zipfile.ZipFile(path) as zf:
        zf.extractall(out_dir)


def zip_docx(folder: Path, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in folder.rglob("*"):
            if item.is_file():
                zf.write(item, item.relative_to(folder).as_posix())


def read_xml(path: Path) -> ET.ElementTree:
    return ET.parse(path)


def rels_path_for_part(part_name: str) -> str:
    directory, basename = posixpath.split(part_name)
    return posixpath.join(directory, "_rels", f"{basename}.rels")


def resolve_relationship_target(part_name: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part_name), target))


def relative_relationship_target(part_name: str, target_part_name: str) -> str:
    return posixpath.relpath(target_part_name, posixpath.dirname(part_name))


def relationship_is_external(rel: ET.Element) -> bool:
    return rel.get("TargetMode") == "External"


def rel_type(rel: ET.Element) -> str:
    return rel.get("Type", "")


def find_last_section_properties(document_root: ET.Element) -> ET.Element:
    sections = document_root.findall(f".//{W}sectPr")
    if not sections:
        body = document_root.find(f"{W}body")
        if body is None:
            raise SystemExit("Target document.xml has no body.")
        section = ET.SubElement(body, f"{W}sectPr")
        return section
    return sections[-1]


def copy_element(element: ET.Element) -> ET.Element:
    return ET.fromstring(ET.tostring(element, encoding="utf-8"))


def source_header_footer_refs(source_doc_root: ET.Element) -> list[ET.Element]:
    source_section = find_last_section_properties(source_doc_root)
    refs = [
        item
        for item in list(source_section)
        if item.tag in {qn("w", "headerReference"), qn("w", "footerReference")}
    ]
    if not refs:
        raise SystemExit("Source document has no header/footer references to import.")
    return refs


def remove_target_header_footer_parts(target_root: Path) -> None:
    word_dir = target_root / "word"
    for pattern in ("header*.xml", "footer*.xml"):
        for item in word_dir.glob(pattern):
            item.unlink(missing_ok=True)
    rels_dir = word_dir / "_rels"
    for pattern in ("header*.xml.rels", "footer*.xml.rels"):
        for item in rels_dir.glob(pattern):
            item.unlink(missing_ok=True)


def load_relationships(rels_file: Path) -> ET.ElementTree:
    if rels_file.exists():
        return ET.parse(rels_file)
    root = ET.Element(f"{REL}Relationships")
    return ET.ElementTree(root)


def next_relationship_id(rels_root: ET.Element) -> str:
    values: list[int] = []
    for rel in rels_root.findall(f"{REL}Relationship"):
        rel_id = rel.get("Id", "")
        if rel_id.startswith("rId") and rel_id[3:].isdigit():
            values.append(int(rel_id[3:]))
    return f"rId{max(values, default=0) + 1}"


def remove_document_header_footer_relationships(rels_root: ET.Element) -> None:
    for rel in list(rels_root):
        if rel.tag == f"{REL}Relationship" and rel_type(rel) in {HEADER_REL_TYPE, FOOTER_REL_TYPE}:
            rels_root.remove(rel)


def remove_header_footer_refs(section: ET.Element) -> None:
    for item in list(section):
        if item.tag in {qn("w", "headerReference"), qn("w", "footerReference"), qn("w", "titlePg")}:
            section.remove(item)


def rel_by_id(rels_root: ET.Element, rel_id: str) -> ET.Element | None:
    for rel in rels_root.findall(f"{REL}Relationship"):
        if rel.get("Id") == rel_id:
            return rel
    return None


def used_part_names(root: Path) -> set[str]:
    names: set[str] = set()
    for item in root.rglob("*"):
        if item.is_file():
            names.add(item.relative_to(root).as_posix())
    return names


def next_part_name(existing: set[str], prefix: str, suffix: str = ".xml") -> str:
    index = 1
    while True:
        candidate = f"word/{prefix}{index}{suffix}"
        if candidate not in existing:
            existing.add(candidate)
            return candidate
        index += 1


def copy_part_and_related_files(
    source_root: Path,
    target_root: Path,
    source_part: str,
    target_part: str,
    copied: dict[str, str],
) -> None:
    if copied.get(source_part) == target_part:
        return
    copied[source_part] = target_part

    source_file = source_root / source_part
    if not source_file.exists():
        raise SystemExit(f"Referenced source part not found: {source_part}")

    target_file = target_root / target_part
    target_file.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_file, target_file)

    source_rels_name = rels_path_for_part(source_part)
    source_rels = source_root / source_rels_name
    if not source_rels.exists():
        return

    target_rels_name = rels_path_for_part(target_part)
    target_rels = target_root / target_rels_name
    target_rels.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_rels, target_rels)

    rels_tree = ET.parse(source_rels)
    for rel in rels_tree.getroot().findall(f"{REL}Relationship"):
        if relationship_is_external(rel):
            continue
        target = rel.get("Target")
        if not target:
            continue
        source_child = resolve_relationship_target(source_part, target)
        target_child = resolve_relationship_target(target_part, target)
        copy_part_and_related_files(source_root, target_root, source_child, target_child, copied)


def content_type_maps(content_types_root: ET.Element) -> tuple[dict[str, str], dict[str, str]]:
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for item in content_types_root:
        if item.tag == f"{CT}Default":
            ext = item.get("Extension")
            content_type = item.get("ContentType")
            if ext and content_type:
                defaults[ext] = content_type
        elif item.tag == f"{CT}Override":
            part = item.get("PartName")
            content_type = item.get("ContentType")
            if part and content_type:
                overrides[part] = content_type
    return defaults, overrides


def remove_override(content_types_root: ET.Element, part_name: str) -> None:
    part = f"/{part_name.lstrip('/')}"
    for item in list(content_types_root):
        if item.tag == f"{CT}Override" and item.get("PartName") == part:
            content_types_root.remove(item)


def ensure_default(content_types_root: ET.Element, extension: str, content_type: str) -> None:
    for item in content_types_root.findall(f"{CT}Default"):
        if item.get("Extension") == extension:
            return
    ET.SubElement(content_types_root, f"{CT}Default", {"Extension": extension, "ContentType": content_type})


def ensure_override(content_types_root: ET.Element, part_name: str, content_type: str) -> None:
    remove_override(content_types_root, part_name)
    ET.SubElement(
        content_types_root,
        f"{CT}Override",
        {"PartName": f"/{part_name.lstrip('/')}", "ContentType": content_type},
    )


def patch_content_types(
    source_root: Path,
    target_root: Path,
    copied_parts: dict[str, str],
    header_footer_parts: dict[str, str],
) -> None:
    source_ct_tree = ET.parse(source_root / "[Content_Types].xml")
    target_ct_tree = ET.parse(target_root / "[Content_Types].xml")
    source_defaults, source_overrides = content_type_maps(source_ct_tree.getroot())
    target_root_el = target_ct_tree.getroot()

    for item in list(target_root_el):
        if item.tag == f"{CT}Override":
            part = item.get("PartName", "")
            if part.startswith("/word/header") or part.startswith("/word/footer"):
                target_root_el.remove(item)

    for source_part, target_part in copied_parts.items():
        source_override = source_overrides.get(f"/{source_part}")
        if source_override:
            ensure_override(target_root_el, target_part, source_override)

        suffix = Path(target_part).suffix.lstrip(".")
        if suffix in source_defaults:
            ensure_default(target_root_el, suffix, source_defaults[suffix])

    for target_part, kind in header_footer_parts.items():
        ensure_override(
            target_root_el,
            target_part,
            HEADER_CONTENT_TYPE if kind == "header" else FOOTER_CONTENT_TYPE,
        )

    target_ct_tree.write(target_root / "[Content_Types].xml", encoding="UTF-8", xml_declaration=True)


def sync_even_odd_setting(source_root: Path, target_root: Path) -> None:
    source_settings = source_root / "word" / "settings.xml"
    target_settings = target_root / "word" / "settings.xml"
    if not source_settings.exists() or not target_settings.exists():
        return

    source_tree = ET.parse(source_settings)
    target_tree = ET.parse(target_settings)
    source_has_even_odd = source_tree.getroot().find(f"{W}evenAndOddHeaders") is not None
    target_root_el = target_tree.getroot()

    for item in list(target_root_el):
        if item.tag == f"{W}evenAndOddHeaders":
            target_root_el.remove(item)
    if source_has_even_odd:
        target_root_el.insert(0, ET.Element(f"{W}evenAndOddHeaders"))

    target_tree.write(target_settings, encoding="UTF-8", xml_declaration=True)


def import_header_footer(source_docx: Path, target_docx: Path, output_docx: Path) -> None:
    with tempfile.TemporaryDirectory(prefix="hf-import-") as tmp:
        tmp_path = Path(tmp)
        source_root = tmp_path / "source"
        target_root = tmp_path / "target"
        unzip_docx(source_docx, source_root)
        unzip_docx(target_docx, target_root)

        source_doc_tree = ET.parse(source_root / "word" / "document.xml")
        target_doc_tree = ET.parse(target_root / "word" / "document.xml")
        source_refs = source_header_footer_refs(source_doc_tree.getroot())

        source_rels_tree = ET.parse(source_root / "word" / "_rels" / "document.xml.rels")
        source_rels_root = source_rels_tree.getroot()
        target_rels_file = target_root / "word" / "_rels" / "document.xml.rels"
        target_rels_tree = load_relationships(target_rels_file)
        target_rels_root = target_rels_tree.getroot()

        remove_target_header_footer_parts(target_root)
        remove_document_header_footer_relationships(target_rels_root)

        target_section = find_last_section_properties(target_doc_tree.getroot())
        remove_header_footer_refs(target_section)

        source_section = find_last_section_properties(source_doc_tree.getroot())
        source_title_pg = source_section.find(f"{W}titlePg")

        existing = used_part_names(target_root)
        copied_parts: dict[str, str] = {}
        header_footer_parts: dict[str, str] = {}
        new_refs: list[ET.Element] = []

        for source_ref in source_refs:
            source_rel_id = source_ref.get(qn("r", "id"))
            if not source_rel_id:
                continue
            source_rel = rel_by_id(source_rels_root, source_rel_id)
            if source_rel is None:
                raise SystemExit(f"Source relationship not found: {source_rel_id}")

            kind = "header" if rel_type(source_rel) == HEADER_REL_TYPE else "footer"
            source_part = resolve_relationship_target("word/document.xml", source_rel.get("Target", ""))
            target_part = next_part_name(existing, kind)
            copy_part_and_related_files(source_root, target_root, source_part, target_part, copied_parts)
            header_footer_parts[target_part] = kind

            new_rel_id = next_relationship_id(target_rels_root)
            ET.SubElement(
                target_rels_root,
                f"{REL}Relationship",
                {
                    "Id": new_rel_id,
                    "Type": rel_type(source_rel),
                    "Target": relative_relationship_target("word/document.xml", target_part),
                },
            )

            new_ref = copy_element(source_ref)
            new_ref.set(qn("r", "id"), new_rel_id)
            new_refs.append(new_ref)

        insert_at = 0
        for ref in new_refs:
            target_section.insert(insert_at, ref)
            insert_at += 1
        if source_title_pg is not None:
            target_section.insert(insert_at, copy_element(source_title_pg))

        patch_content_types(source_root, target_root, copied_parts, header_footer_parts)
        sync_even_odd_setting(source_root, target_root)

        target_doc_tree.write(target_root / "word" / "document.xml", encoding="UTF-8", xml_declaration=True)
        target_rels_tree.write(target_rels_file, encoding="UTF-8", xml_declaration=True)
        zip_docx(target_root, output_docx)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Copy headers and footers from a company DOCX into the Markdown template DOCX."
    )
    parser.add_argument("source", help="DOCX that contains the desired company headers/footers")
    parser.add_argument(
        "target",
        nargs="?",
        default=str(Path(__file__).resolve().parent / "template.docx"),
        help="Template DOCX to receive the headers/footers; defaults to _md2word/template.docx",
    )
    parser.add_argument("-o", "--output", help="Output DOCX; defaults to TARGET.with-header-footer.docx")
    parser.add_argument("--in-place", action="store_true", help="Replace TARGET in place")
    return parser.parse_args()


def default_output_path(target: Path) -> Path:
    return target.with_name(f"{target.stem}.with-header-footer{target.suffix}")


def main() -> int:
    args = parse_args()
    source_docx = Path(args.source).resolve()
    target_docx = Path(args.target).resolve()
    if not source_docx.exists():
        raise SystemExit(f"Source DOCX not found: {source_docx}")
    if not target_docx.exists():
        raise SystemExit(f"Target DOCX not found: {target_docx}")

    if args.in_place and args.output:
        raise SystemExit("Use either --in-place or --output, not both.")

    output_docx = target_docx if args.in_place else Path(args.output).resolve() if args.output else default_output_path(target_docx)

    if output_docx.resolve() == target_docx.resolve():
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as fh:
            temp_output = Path(fh.name)
        try:
            import_header_footer(source_docx, target_docx, temp_output)
            shutil.move(str(temp_output), str(target_docx))
        finally:
            temp_output.unlink(missing_ok=True)
    else:
        import_header_footer(source_docx, target_docx, output_docx)

    print(f"Created: {output_docx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
