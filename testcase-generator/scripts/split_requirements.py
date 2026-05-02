"""
按模块切分需求文档，降低后续生成测试用例时的上下文压力。

用法：
    python split_requirements.py <input.txt|input.md> <output_dir>
"""

import json
import re
import sys
from pathlib import Path


HEADING_PATTERNS = [
    re.compile(r"^\s*功能模块([一二三四五六七八九十百千万]+)[、\.．]\s*(.+?)\s*$"),
    re.compile(r"^\s*([一二三四五六七八九十百千万]+)[、\.．]\s*(.+?)\s*$"),
]


def slugify(text):
    text = re.sub(r"[\\/:*?\"<>|]+", "_", text.strip())
    text = re.sub(r"\s+", "_", text)
    return text[:80] or "module"


def is_heading(line):
    stripped = line.strip()
    if not stripped:
        return None
    for pattern in HEADING_PATTERNS:
        match = pattern.match(stripped)
        if match:
            return {
                "raw": stripped,
                "order": match.group(1),
                "title": match.group(2).strip(),
            }
    return None


def split_modules(text):
    modules = []
    current = None

    for line_no, line in enumerate(text.splitlines(), start=1):
        heading = is_heading(line)
        if heading:
            if current:
                current["content"] = current["content"].strip()
                current["line_end"] = line_no - 1
                modules.append(current)
            current = {
                "module_order": heading["order"],
                "module_name": heading["title"],
                "heading": heading["raw"],
                "line_start": line_no,
                "content": "",
            }
            continue

        if current:
            current["content"] += line + "\n"

    if current:
        current["content"] = current["content"].strip()
        current["line_end"] = len(text.splitlines())
        modules.append(current)

    return modules


def summarize_module(content):
    lines = [line.strip() for line in content.splitlines() if line.strip()]
    preview = lines[:8]
    return {
        "line_count": len(lines),
        "preview": preview,
    }


def write_outputs(input_path, output_dir, modules):
    output_dir.mkdir(parents=True, exist_ok=True)
    modules_dir = output_dir / "modules"
    modules_dir.mkdir(parents=True, exist_ok=True)

    index = {
        "source": str(input_path.resolve()),
        "module_count": len(modules),
        "modules": [],
    }

    for idx, module in enumerate(modules, start=1):
        summary = summarize_module(module["content"])
        filename = f"{idx:02d}_{slugify(module['module_name'])}.txt"
        file_path = modules_dir / filename
        file_path.write_text(module["content"] + "\n", encoding="utf-8")

        index["modules"].append({
            "index": idx,
            "module_order": module["module_order"],
            "module_name": module["module_name"],
            "heading": module["heading"],
            "line_start": module["line_start"],
            "line_end": module["line_end"],
            "line_count": summary["line_count"],
            "preview": summary["preview"],
            "path": str(file_path.resolve()),
        })

    index_path = output_dir / "module_index.json"
    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")

    return {
        "status": "ok",
        "source": str(input_path.resolve()),
        "module_count": len(modules),
        "index_path": str(index_path.resolve()),
        "modules_dir": str(modules_dir.resolve()),
    }


def main():
    if len(sys.argv) != 3:
        print(json.dumps({
            "status": "error",
            "message": "用法: python split_requirements.py <input.txt|input.md> <output_dir>",
        }, ensure_ascii=False))
        sys.exit(1)

    input_path = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])

    if not input_path.exists():
        print(json.dumps({
            "status": "error",
            "message": f"输入文件不存在: {input_path}",
        }, ensure_ascii=False))
        sys.exit(1)

    text = input_path.read_text(encoding="utf-8-sig")
    modules = split_modules(text)
    if not modules:
        print(json.dumps({
            "status": "error",
            "message": "未识别到模块标题，无法自动切分需求文档",
        }, ensure_ascii=False))
        sys.exit(1)

    print(json.dumps(write_outputs(input_path, output_dir, modules), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
