"""
合并模块级测试用例 JSON。

用法：
    python merge_cases.py <modules_dir> <output_all_cases.json>
"""

import json
import sys
from pathlib import Path


def load_cases(json_path):
    data = json.loads(json_path.read_text(encoding="utf-8-sig"))
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        return data.get("test_cases", data.get("cases", []))
    return []


def main():
    if len(sys.argv) != 3:
        print(json.dumps({
            "status": "error",
            "message": "用法: python merge_cases.py <modules_dir> <output_all_cases.json>",
        }, ensure_ascii=False))
        sys.exit(1)

    modules_dir = Path(sys.argv[1])
    output_path = Path(sys.argv[2])

    if not modules_dir.exists() or not modules_dir.is_dir():
        print(json.dumps({
            "status": "error",
            "message": f"目录不存在: {modules_dir}",
        }, ensure_ascii=False))
        sys.exit(1)

    all_cases = []
    seen_ids = {}
    warnings = []
    module_stats = []

    for json_file in sorted(modules_dir.glob("*.json")):
        if json_file.name.lower() == "all_cases.json":
            continue
        try:
            cases = load_cases(json_file)
        except json.JSONDecodeError as exc:
            warnings.append(f"JSON 解析失败 [{json_file.name}]: {exc}")
            continue

        module_stats.append({"file": json_file.name, "count": len(cases)})
        for case in cases:
            case_id = case.get("用例编号", "") or case.get("id", "")
            if case_id and case_id in seen_ids:
                warnings.append(f"重复用例编号 [{case_id}]，首次出现于 {seen_ids[case_id]}")
                continue
            if case_id:
                seen_ids[case_id] = json_file.name
            all_cases.append(case)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(all_cases, ensure_ascii=False, indent=2), encoding="utf-8")

    result = {
        "status": "ok",
        "path": str(output_path.resolve()),
        "count": len(all_cases),
        "modules": module_stats,
    }
    if warnings:
        result["warnings"] = warnings
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
