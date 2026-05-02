# 输出格式说明

## 最终交付规则

最终交付格式必须是 Excel。

- Markdown 只能作为中间预览
- JSON 只能作为脚本输入输出和模块级暂存
- 最后必须生成 `.xlsx` 文件，并返回文件绝对路径

## Markdown

适合在对话中直接展示或评审，但不是最终交付格式。

推荐结构：

1. 需求概述
2. 模块清单
3. 按模块展示测试用例表格
4. 最后给出总数统计

## JSON

推荐使用扁平数组，键名与 Excel 列头一致，仅作为中间格式，例如：

```json
[
  {
    "功能模块": "登录模块",
    "用例编号": "LOGIN_001",
    "功能点": "账号密码登录",
    "用例标题": "验证登录成功",
    "优先级": "High",
    "前置条件": "已存在可登录账号",
    "测试步骤": "1. 打开登录页\n2. 输入正确账号和密码\n3. 点击登录",
    "预期结果": "登录成功并跳转首页",
    "设计人": "",
    "执行结果": "",
    "执行人": "",
    "备注": ""
  }
]
```

## Excel

这是最终交付格式。

通过以下命令导出：

```bash
python "<skill_dir>/scripts/generate_xlsx.py" "<all_cases.json>" "<output.xlsx>"
```

如果模块 JSON 分散存放：

```bash
python "<skill_dir>/scripts/generate_xlsx.py" --merge "<modules_dir>" "<output.xlsx>"
```

## 合并总表

需要先合并模块 JSON 时，执行：

```bash
python "<skill_dir>/scripts/merge_cases.py" "<modules_dir>" "<output_dir>/all_cases.json"
```

返回结果中的 `count` 代表总用例数，`path` 代表输出文件绝对路径。

推荐标准收口流程：

1. 各模块分别生成 `*.json`
2. 合并为 `all_cases.json`
3. 基于 `all_cases.json` 导出最终 `output.xlsx`
4. 向用户返回 `output.xlsx` 的绝对路径
