from __future__ import annotations

import json
from pathlib import Path

from build_ansichten_pages import HELP_ROOT, Node, assign_paths, is_page, read_nodes


OUTPUT = HELP_ROOT / "ansichten-tree.js"
JSON_OUTPUT = HELP_ROOT / "ansichten-tree.json"


def public_path(node: Node) -> str:
    if not node.public_url:
        return "/de/help/Ansichten/"
    return node.public_url.replace("https://x-erp.de", "")


def nearest_page_parent(node: Node) -> Node | None:
    current = node.parent
    while current:
        if is_page(current):
            return current
        current = current.parent
    return None


def build_page_children(page_nodes: list[Node]) -> dict[int, list[Node]]:
    page_set = {id(node) for node in page_nodes}
    result: dict[int, list[Node]] = {id(node): [] for node in page_nodes}
    for node in page_nodes:
        parent = nearest_page_parent(node)
        if parent and id(parent) in page_set:
            result[id(parent)].append(node)
    return result


def to_tree(node: Node, children_by_parent: dict[int, list[Node]]) -> dict[str, object]:
    children = [to_tree(child, children_by_parent) for child in children_by_parent.get(id(node), [])]
    result: dict[str, object] = {
        "title": node.topic,
        "path": public_path(node),
        "sourceRow": node.row,
        "sourceLevel": node.level,
    }
    if children:
        result["children"] = children
    return result


def main() -> None:
    nodes = read_nodes()
    page_nodes = assign_paths(nodes)
    root = page_nodes[0]
    children_by_parent = build_page_children(page_nodes)
    tree = to_tree(root, children_by_parent)
    payload = json.dumps(tree, ensure_ascii=False, indent=2)
    JSON_OUTPUT.write_text(payload + "\n", encoding="utf-8")
    OUTPUT.write_text(
        "// Generated from X-ERP-HELP.xlsx. Do not edit manually.\n"
        f"window.XERP_ANSICHTEN_TREE = {payload};\n",
        encoding="utf-8",
    )
    print(f"Wrote {OUTPUT} and {JSON_OUTPUT} with {len(page_nodes)} page nodes")


if __name__ == "__main__":
    main()
