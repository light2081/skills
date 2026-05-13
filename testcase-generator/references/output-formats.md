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

## 模块级工作计划

生成用例前必须先保存或展示模块级工作计划。计划用于分发子 agent、串行回退和最终校验。

推荐结构：

```json
[
  {
    "module": "订单提交",
    "source": "modules/order_submit.txt",
    "functions": ["商品校验", "库存校验", "价格计算", "提交订单"],
    "dependencies": ["购物车", "库存", "优惠券"],
    "methods": ["等价类", "边界值", "判定表", "错误猜测"],
    "output": "modules/order_submit.json",
    "parallel": true
  }
]
```

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

模块级 JSON 文件仍使用同一扁平数组格式。数组中每条用例的 `功能模块` 必须与模块计划中的 `module` 一致，避免一个模块文件混入其他模块用例。`用例编号` 默认使用模块前缀，合并前必须保证全局唯一，避免并行生成后出现重复编号。

合并前应检查：

- JSON 可以正常解析
- 顶层结构是数组
- 每条用例包含固定列头字段
- `用例编号` 不为空且全局唯一
- 同一模块文件没有混入其他模块名称

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

1. 输出模块识别结果
2. 输出模块级工作计划
3. 各模块分别生成 `*.json`
4. 校验模块 JSON，失败模块按计划串行重生成
5. 合并为 `all_cases.json`
6. 基于 `all_cases.json` 导出最终 `output.xlsx`
7. 向用户返回 `output.xlsx` 的绝对路径和用例总数
