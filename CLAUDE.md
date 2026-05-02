# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个面向 AI 编程协作的技能集合，当前包含三类能力：

- `docx-cn`：Word 文档（`.docx`）创建、读取、编辑与分析
- `jmeter-generator`：根据接口文档生成 JMeter `.jmx` 压测脚本
- `testcase-generator`：根据需求文档生成 Excel 测试用例

## 技术栈

**主要技术栈：**

- **docx-cn**：基于 Python 的 Word 文档处理，使用 `python-docx` 库
- **jmeter-generator**：基于 XML 的 JMeter 脚本生成，使用 Python 进行 XML 处理
- **testcase-generator**：基于 Python 的测试用例生成，使用 `openpyxl` 库生成 Excel

**依赖关系：**

- `python-docx`：用于 Word 文档读取和处理
- `openpyxl`：用于 Excel 文件生成
- `pandoc`：用于文本提取
- `LibreOffice`：用于 PDF 转换和文档处理
- `docx`（Node.js）：用于新建 Word 文档

## 项目结构

项目采用扁平化的目录结构，主要包含三个技能目录：

```
├── docx-cn/                     # Word 文档处理技能
│   ├── SKILL.md                 # 技能说明文档
│   ├── scripts/                # Python 脚本
│   │   ├── office/             # Office 处理工具
│   │   │   ├── helpers/        # 辅助工具
│   │   │   ├── schemas/        # XML Schema 定义
│   │   │   ├── soffice.py      # LibreOffice 集成
│   │   │   ├── unpack.py       # 解包 DOCX
│   │   │   └── validate.py     # 验证 DOCX
│   │   └── styles.js           # 共享样式定义
│   └── templates/              # 模板文件
├── jmeter-generator/            # JMeter 脚本生成技能
│   └── SKILL.md                # 技能说明文档
└── testcase-generator/          # 测试用例生成技能
    ├── SKILL.md                # 技能说明文档
    ├── references/             # 参考文档
    │   ├── analysis-guide.md   # 分析指南
    │   ├── output-formats.md   # 输出格式
    │   └── test-case-fields.md # 测试用例字段
    └── scripts/               # Python 脚本
        ├── generate_xlsx.py    # 生成 Excel
        ├── merge_cases.py     # 合并用例
        └── split_requirements.py # 拆分需求文档
```

## 核心架构

**技能架构：**

- 每个技能都有独立的 `SKILL.md` 文件，定义了技能的触发条件、执行流程和详细说明
- 技能通过脚本实现具体功能，脚本位于 `scripts/` 目录下
- 共享样式和配置通过 `styles.js` 等文件统一管理

**主要组件：**

1. **docx-cn 技能**：

   - 文档创建：使用 `docx` JavaScript 库
   - 文档编辑：通过解包、修改 XML、重新打包的方式
   - 文档分析：使用 `pandoc` 进行文本提取
   - 修订处理：支持接受修订、添加评论等功能
2. **jmeter-generator 技能**：

   - 接口解析：支持多种输入格式（自然语言、cURL、HTTP 报文、Swagger、Markdown、Word）
   - JMX 生成：基于 XML 模板生成 JMeter 测试脚本
   - 环境检查：自动检测和安装依赖
3. **testcase-generator 技能**：

   - 需求解析：支持多种文档格式
   - 模块划分：自动识别功能模块
   - 测试用例生成：使用等价类、边界值、判定表等方法
   - Excel 输出：使用 `openpyxl` 生成标准化测试用例表格

## 开发指南

**环境要求：**

- Python 3.x 环境
- Node.js 环境（用于 docx-js）
- LibreOffice（用于文档转换）
- Pandoc（用于文本提取）

**常见开发任务：**

- 添加新技能：创建新的技能目录，包含 SKILL.md 和 scripts 目录
- 修改现有技能：编辑对应技能目录下的脚本文件
- 更新依赖：检查并更新 requirements.txt 或 package.json 文件

**注意事项：**

- 每个技能都是独立的，修改时请确保不影响其他技能
- 脚本文件应包含适当的错误处理和日志记录
- 文档输出应保持标准化格式
