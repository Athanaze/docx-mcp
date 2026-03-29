"""
Media inspection operations for .docx files.

Focuses on embedded images so agents can reason about binary payloads without
dumping raw ZIP contents.
"""
import json
import posixpath
import zipfile
from xml.etree import ElementTree as ET

from word_document_server.document import raw_docx_tool


_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_IMG_REL_SUFFIX = "/relationships/image"
def _normalize_part(path: str) -> str:
    return posixpath.normpath(path).lstrip("/")


def _resolve_rel_target(rels_path: str, target: str) -> str:
    base_dir = posixpath.dirname(rels_path).replace("_rels", "")
    return _normalize_part(posixpath.join(base_dir, target))


@raw_docx_tool(readonly=True)
def list_document_images(filename, include_usage=True):
    """List embedded images with size/type and optional usage statistics."""
    try:
        with zipfile.ZipFile(filename) as zf:
            names = set(zf.namelist())
            media_paths = sorted(
                n for n in names
                if n.startswith("word/media/") and not n.endswith("/")
            )
            images = []
            path_to_info = {}
            for media_path in media_paths:
                data = zf.read(media_path)
                ext = media_path.rsplit(".", 1)[-1].lower() if "." in media_path else ""
                info = {
                    "path": media_path,
                    "filename": posixpath.basename(media_path),
                    "extension": ext,
                    "size_bytes": len(data),
                    "size_kb": round(len(data) / 1024.0, 2),
                    "usage_count": 0,
                    "used_in": [],
                }
                images.append(info)
                path_to_info[media_path] = info

            if include_usage and media_paths:
                rel_sources = sorted(
                    n for n in names
                    if n.startswith("word/_rels/") and n.endswith(".rels")
                )
                for rels_path in rel_sources:
                    try:
                        rel_root = ET.fromstring(zf.read(rels_path))
                    except Exception:
                        continue
                    source_part = _normalize_part(
                        posixpath.join(
                            posixpath.dirname(rels_path).replace("_rels", ""),
                            posixpath.basename(rels_path)[:-5],
                        )
                    )
                    if source_part not in names:
                        continue
                    try:
                        source_xml = zf.read(source_part).decode("utf-8", errors="ignore")
                    except Exception:
                        continue
                    for rel in rel_root.findall(f"{_REL_NS}Relationship"):
                        rel_type = rel.attrib.get("Type", "")
                        if not rel_type.endswith(_IMG_REL_SUFFIX):
                            continue
                        rid = rel.attrib.get("Id")
                        target = rel.attrib.get("Target", "")
                        target_part = _resolve_rel_target(rels_path, target)
                        info = path_to_info.get(target_part)
                        if not info:
                            continue
                        hits = 0
                        if rid:
                            # precise count for this relationship id
                            hits = source_xml.count(f'r:embed="{rid}"') + source_xml.count(
                                f'r:link="{rid}"'
                            )
                        if hits <= 0:
                            hits = 1
                        info["usage_count"] += hits
                        if source_part not in info["used_in"]:
                            info["used_in"].append(source_part)

            total_bytes = sum(i["size_bytes"] for i in images)
            return json.dumps(
                {
                    "images": images,
                    "count": len(images),
                    "total_size_bytes": total_bytes,
                    "total_size_kb": round(total_bytes / 1024.0, 2),
                },
                indent=2,
            )
    except Exception as e:
        return f"Failed to inspect media: {e}"
