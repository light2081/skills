---
name: jmeter-generator
description: 根据 HTTP 接口文档自动生成 JMeter 压测脚本（.jmx 格式）。支持自然语言描述、cURL 命令、HTTP 报文、Swagger/OpenAPI 片段、Markdown 文件（.md）和 Word 文档（.docx）作为输入。当用户提到以下任意场景时主动触发：生成 JMeter 脚本、生成压测脚本、生成 jmx、接口性能测试、负载测试、并发测试、压力测试、需要生成测试脚本，即使用户没有明确说"JMeter"也要触发。
---

# jmeter-generator

根据用户提供的 HTTP 接口信息，生成可直接导入 JMeter 的 `.jmx` 测试脚本文件。

## 触发条件

用户输入包含以下任意关键词时触发本 Skill：
- "生成 jmx"、"生成 jmeter 脚本"、"生成压测脚本"
- "接口性能测试"、"负载测试"、"并发测试"、"压力测试"
- "jmeter script"、"jmx script"
- 提供接口文档并需要生成测试脚本
- 直接调用 `/jmeter-generator`

---

## 执行流程

### 第零步：环境依赖检查

**仅当用户输入包含 `.docx` 文件路径时，才需执行此步骤。其他输入格式可直接跳至第一步。**

#### 检测 Python 是否存在

使用 Bash 工具执行：

```bash
python --version 2>&1
```

- **输出包含版本号**（如 `Python 3.x.x`）：Python 可用，继续检测依赖。
- **输出包含 `not found` 或命令不存在**：告知用户：

  > 未检测到 Python 环境。请先安装 Python 3.x（https://www.python.org/downloads/），安装后重新执行本命令。

  **终止当前流程，不再继续。**

#### 检测 python-docx

使用 Bash 工具执行：

```bash
python -c "import docx; print('ok')" 2>&1
```

- **输出为 `ok`**：依赖就绪，跳过安装，直接进入第一步。
- **输出含 `ModuleNotFoundError` 或其他错误**：进入自动安装流程。

#### 自动安装 python-docx

```bash
pip install python-docx 2>&1
```

#### 安装后验证

```bash
python -c "import docx; print('ok')" 2>&1
```

- **输出为 `ok`**：告知用户"python-docx 已安装完成"，继续第一步。
- **失败**：告知用户：

  > python-docx 自动安装失败，请手动执行：`pip install python-docx`
  > 安装完成后重新执行本命令。

  **终止当前流程，不再继续。**

---

### 第一步：解析用户输入

从用户输入中提取接口信息，支持以下格式：

- **自然语言**：如"POST /api/login，参数 username 和 password"
- **cURL 命令**：解析 `-X`、`-H`、`-d`、`--data`、`-u` 等参数
- **HTTP 报文**：解析请求行、请求头、请求体
- **Swagger/OpenAPI 片段**：解析 path、method、parameters、requestBody
- **Markdown 文件**：用户提供 `.md` 文件路径时，使用 Read 工具读取文件内容后按上述规则解析
- **Word 文档**：用户提供 `.docx` 文件路径时，使用 Bash 工具调用 python-docx 提取文本后解析

**解析时同步提取以下所有字段**（包括性能参数，避免第三步重复询问）：
```
host, port, protocol, path, method,
headers[], params[], body, content_type,
num_threads（线程数/并发数）, ramp_up（Ramp-up秒）, loop_count（循环次数）
```

#### 文件输入处理规则

1. **识别文件路径**：若用户消息中包含以 `.md` 或 `.docx` 结尾的路径（绝对路径或相对路径均可），视为文件输入。
2. **读取文件**：
   - `.md` 文件：使用 Read 工具直接读取全部内容。
   - `.docx` 文件：使用 Bash 工具调用 python-docx 提取段落文本：
     ```bash
     python -c "
import docx, sys
doc = docx.Document(sys.argv[1])
for p in doc.paragraphs:
    if p.text.strip():
        print(p.text)
for t in doc.tables:
    for row in t.rows:
        print('\t'.join(c.text.strip() for c in row.cells))
" "path/to/file.docx" 2>&1
     ```
     输出为纯文本，再按下述策略解析。
3. **内容解析**：将读取到的文本按以下策略提取接口信息：
   - 查找 HTTP 方法关键词（`GET`、`POST`、`PUT`、`DELETE`、`PATCH`）及其附近的路径
   - 查找 URL、域名、BaseURL、Host 等字段
   - 查找请求头描述（`Headers`、`Authorization`、`Content-Type` 等）
   - 查找请求体示例（JSON 代码块、表格参数、示例数据）
   - 查找响应示例中的预期状态码
4. **解析失败处理**：若文件内容无法提取到任何接口信息，告知用户"未能从文件中识别出接口信息"，并请用户补充说明或直接粘贴关键片段。
5. **Word 文件读取失败处理**：若 python-docx 执行出错（如文件损坏、加密保护等），提示用户将文档另存为 `.md` 或直接粘贴接口文本。

解析时提取以下字段：（已移至上方，见"同步提取"说明）

---

### 第二步：多接口检测

**若检测到用户描述了 2 个及以上接口**，停止解析，询问：

> 检测到您描述了 **N 个接口**：
> 1. [METHOD] [path] — [简短描述]
> 2. [METHOD] [path] — [简短描述]
> ...
>
> ⚠️ **若上述接口属于不同的域名/IP**，请注意：选 A 或 C 方案时，每个接口将保留各自独立的 Host 配置。
>
> 请问这些接口如何处理？
> - **A. 合并到一个 Test Plan**（同一线程组，按顺序执行）
> - **B. 每个接口生成独立的 `.jmx` 文件**
> - **C. 每个接口一个线程组，但在同一 `.jmx` 文件中**

等待用户选择后继续。

---

### 第三步：信息完整性校验

对每个接口逐一校验。**必填项缺失时立即追问，不可跳过：**

| 必填字段 | 追问话术 |
|----------|----------|
| `host` | "请提供接口的域名或 IP 地址（如 `api.example.com` 或 `192.168.1.1`）" |
| `path` | "请提供接口路径（如 `/api/v1/login`）" |
| `method` | "请提供请求方法：GET / POST / PUT / DELETE / PATCH ？" |

**若有多个必填项同时缺失，合并为一次提问，列出所有缺失字段。**

**选填项缺失时，列出默认值，询问是否调整（一次性批量）。已在第一步中提取到的性能参数（如用户已提供"50个并发"、"循环10次"等）直接使用，无需再次询问：**

> 以下配置将使用默认值，如需修改请告知：
> - 协议：`https`
> - 端口：`443`（https 默认）/ `80`（http 默认）
> - 线程数（并发用户数）：`1`
> - Ramp-up 时间（秒）：`1`
> - 循环次数：`1`
> - 响应断言（HTTP 状态码）：`200`
> - 测试计划名称：`[接口路径末段]-test`
> - 输出文件名：`[测试计划名称].jmx`
>
> 是否需要从 **CSV 文件**读取参数？（适用于批量加载不同用户名/密码、订单 ID 等测试数据）
> - **是**：请告知 CSV 文件路径及字段名，将在脚本中插入 CSV Data Set Config 组件
> - **否**：跳过，使用固定参数值

---

### 第四步：生成 JMX 文件

收集完所有信息后，按照以下模板生成 JMX XML，并使用 Write 工具写入 `.jmx` 文件。

#### JMX 主体模板

```xml
<?xml version="1.0" encoding="UTF-8"?>
<jmeterTestPlan version="1.2" properties="5.0" jmeter="5.6.3">
  <hashTree>
    <TestPlan guiclass="TestPlanGui" testclass="TestPlan" testname="${TEST_PLAN_NAME}" enabled="true">
      <stringProp name="TestPlan.comments"></stringProp>
      <boolProp name="TestPlan.functional_mode">false</boolProp>
      <boolProp name="TestPlan.serialize_threadgroups">false</boolProp>
      <elementProp name="TestPlan.user_defined_variables" elementType="Arguments" guiclass="ArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true">
        <collectionProp name="Arguments.arguments"/>
      </elementProp>
    </TestPlan>
    <hashTree>
      <ThreadGroup guiclass="ThreadGroupGui" testclass="ThreadGroup" testname="${THREAD_GROUP_NAME}" enabled="true">
        <stringProp name="ThreadGroup.on_sample_error">continue</stringProp>
        <elementProp name="ThreadGroup.main_controller" elementType="LoopController" guiclass="LoopControlPanel" testclass="LoopController" testname="Loop Controller" enabled="true">
          <boolProp name="LoopController.continue_forever">false</boolProp>
          <intProp name="LoopController.loops">${LOOP_COUNT}</intProp>
        </elementProp>
        <stringProp name="ThreadGroup.num_threads">${NUM_THREADS}</stringProp>
        <stringProp name="ThreadGroup.ramp_time">${RAMP_UP}</stringProp>
        <boolProp name="ThreadGroup.same_user_on_next_iteration">true</boolProp>
        <boolProp name="ThreadGroup.scheduler">false</boolProp>
        <stringProp name="ThreadGroup.duration"></stringProp>
        <stringProp name="ThreadGroup.delay"></stringProp>
      </ThreadGroup>
      <hashTree>
        <!-- HTTP 信息头管理器（有请求头时插入此块） -->
        ${HEADER_MANAGER}

        <!-- HTTP 请求采样器（每个接口一个，按顺序排列） -->
        ${HTTP_SAMPLERS}

        <!-- 查看结果树 -->
        <ResultCollector guiclass="ViewResultsFullVisualizer" testclass="ResultCollector" testname="View Results Tree" enabled="true">
          <boolProp name="ResultCollector.error_logging">false</boolProp>
          <objProp>
            <name>saveConfig</name>
            <value class="SampleSaveConfiguration">
              <time>true</time>
              <latency>true</latency>
              <timestamp>true</timestamp>
              <success>true</success>
              <label>true</label>
              <code>true</code>
              <message>true</message>
              <threadName>true</threadName>
              <dataType>true</dataType>
              <encoding>false</encoding>
              <assertions>true</assertions>
              <subresults>true</subresults>
              <responseData>false</responseData>
              <samplerData>false</samplerData>
              <xml>false</xml>
              <fieldNames>true</fieldNames>
              <responseHeaders>false</responseHeaders>
              <requestHeaders>false</requestHeaders>
              <responseDataOnError>false</responseDataOnError>
              <saveAssertionResultsFailureMessage>true</saveAssertionResultsFailureMessage>
              <assertionsResultsToSave>0</assertionsResultsToSave>
              <bytes>true</bytes>
              <sentBytes>true</sentBytes>
              <url>true</url>
              <threadCounts>true</threadCounts>
              <idleTime>true</idleTime>
              <connectTime>true</connectTime>
            </value>
          </objProp>
          <stringProp name="filename"></stringProp>
        </ResultCollector>
        <hashTree/>

        <!-- 汇总报告 -->
        <ResultCollector guiclass="SummaryReport" testclass="ResultCollector" testname="Summary Report" enabled="true">
          <boolProp name="ResultCollector.error_logging">false</boolProp>
          <objProp>
            <name>saveConfig</name>
            <value class="SampleSaveConfiguration">
              <time>true</time>
              <latency>true</latency>
              <timestamp>true</timestamp>
              <success>true</success>
              <label>true</label>
              <code>true</code>
              <message>true</message>
              <threadName>true</threadName>
              <dataType>true</dataType>
              <encoding>false</encoding>
              <assertions>true</assertions>
              <subresults>true</subresults>
              <responseData>false</responseData>
              <samplerData>false</samplerData>
              <xml>false</xml>
              <fieldNames>true</fieldNames>
              <responseHeaders>false</responseHeaders>
              <requestHeaders>false</requestHeaders>
              <responseDataOnError>false</responseDataOnError>
              <saveAssertionResultsFailureMessage>true</saveAssertionResultsFailureMessage>
              <assertionsResultsToSave>0</assertionsResultsToSave>
              <bytes>true</bytes>
              <sentBytes>true</sentBytes>
              <url>true</url>
              <threadCounts>true</threadCounts>
              <idleTime>true</idleTime>
              <connectTime>true</connectTime>
            </value>
          </objProp>
          <stringProp name="filename"></stringProp>
        </ResultCollector>
        <hashTree/>

        <!-- 聚合报告（分析 TPS、响应时间分位数等关键性能指标） -->
        <ResultCollector guiclass="StatVisualizer" testclass="ResultCollector" testname="Aggregate Report" enabled="true">
          <boolProp name="ResultCollector.error_logging">false</boolProp>
          <objProp>
            <name>saveConfig</name>
            <value class="SampleSaveConfiguration">
              <time>true</time>
              <latency>true</latency>
              <timestamp>true</timestamp>
              <success>true</success>
              <label>true</label>
              <code>true</code>
              <message>true</message>
              <threadName>true</threadName>
              <dataType>true</dataType>
              <encoding>false</encoding>
              <assertions>true</assertions>
              <subresults>true</subresults>
              <responseData>false</responseData>
              <samplerData>false</samplerData>
              <xml>false</xml>
              <fieldNames>true</fieldNames>
              <responseHeaders>false</responseHeaders>
              <requestHeaders>false</requestHeaders>
              <responseDataOnError>false</responseDataOnError>
              <saveAssertionResultsFailureMessage>true</saveAssertionResultsFailureMessage>
              <assertionsResultsToSave>0</assertionsResultsToSave>
              <bytes>true</bytes>
              <sentBytes>true</sentBytes>
              <url>true</url>
              <threadCounts>true</threadCounts>
              <idleTime>true</idleTime>
              <connectTime>true</connectTime>
            </value>
          </objProp>
          <stringProp name="filename"></stringProp>
        </ResultCollector>
        <hashTree/>
      </hashTree>
    </hashTree>
  </hashTree>
</jmeterTestPlan>
```

---

#### 子模板：GET 请求采样器

```xml
<HTTPSamplerProxy guiclass="HttpTestSampleGui" testclass="HTTPSamplerProxy" testname="${SAMPLER_NAME}" enabled="true">
  <stringProp name="HTTPSampler.domain">${HOST}</stringProp>
  <stringProp name="HTTPSampler.port">${PORT}</stringProp>
  <stringProp name="HTTPSampler.protocol">${PROTOCOL}</stringProp>
  <stringProp name="HTTPSampler.path">${PATH}</stringProp>
  <stringProp name="HTTPSampler.method">GET</stringProp>
  <boolProp name="HTTPSampler.follow_redirects">true</boolProp>
  <boolProp name="HTTPSampler.use_keepalive">true</boolProp>
  <boolProp name="HTTPSampler.postBodyRaw">false</boolProp>
  <elementProp name="HTTPsampler.Arguments" elementType="Arguments" guiclass="HTTPArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true">
    <collectionProp name="Arguments.arguments">
      ${QUERY_PARAMS}
    </collectionProp>
  </elementProp>
</HTTPSamplerProxy>
<hashTree>
  ${ASSERTION}
</hashTree>
```

#### 子模板：POST/PUT/PATCH/DELETE JSON Body 采样器

> **适用范围**：`POST`、`PUT`、`PATCH`、`DELETE`（带请求体时）均使用此模板，将 `${METHOD}` 替换为对应方法名即可。`DELETE` 无请求体时参考 GET 模板并将 method 改为 `DELETE`。

```xml
<HTTPSamplerProxy guiclass="HttpTestSampleGui" testclass="HTTPSamplerProxy" testname="${SAMPLER_NAME}" enabled="true">
  <stringProp name="HTTPSampler.domain">${HOST}</stringProp>
  <stringProp name="HTTPSampler.port">${PORT}</stringProp>
  <stringProp name="HTTPSampler.protocol">${PROTOCOL}</stringProp>
  <stringProp name="HTTPSampler.path">${PATH}</stringProp>
  <stringProp name="HTTPSampler.method">${METHOD}</stringProp>
  <boolProp name="HTTPSampler.follow_redirects">true</boolProp>
  <boolProp name="HTTPSampler.use_keepalive">true</boolProp>
  <boolProp name="HTTPSampler.postBodyRaw">true</boolProp>
  <elementProp name="HTTPsampler.Arguments" elementType="Arguments" guiclass="HTTPArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true">
    <collectionProp name="Arguments.arguments">
      <elementProp name="" elementType="HTTPArgument">
        <boolProp name="HTTPArgument.always_encode">false</boolProp>
        <stringProp name="Argument.value">${JSON_BODY}</stringProp>
        <stringProp name="Argument.metadata">=</stringProp>
      </elementProp>
    </collectionProp>
  </elementProp>
</HTTPSamplerProxy>
<hashTree>
  ${ASSERTION}
</hashTree>
```

#### 子模板：POST 表单（application/x-www-form-urlencoded）采样器

```xml
<HTTPSamplerProxy guiclass="HttpTestSampleGui" testclass="HTTPSamplerProxy" testname="${SAMPLER_NAME}" enabled="true">
  <stringProp name="HTTPSampler.domain">${HOST}</stringProp>
  <stringProp name="HTTPSampler.port">${PORT}</stringProp>
  <stringProp name="HTTPSampler.protocol">${PROTOCOL}</stringProp>
  <stringProp name="HTTPSampler.path">${PATH}</stringProp>
  <stringProp name="HTTPSampler.method">POST</stringProp>
  <boolProp name="HTTPSampler.follow_redirects">true</boolProp>
  <boolProp name="HTTPSampler.use_keepalive">true</boolProp>
  <boolProp name="HTTPSampler.postBodyRaw">false</boolProp>
  <elementProp name="HTTPsampler.Arguments" elementType="Arguments" guiclass="HTTPArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true">
    <collectionProp name="Arguments.arguments">
      ${FORM_PARAMS}
    </collectionProp>
  </elementProp>
</HTTPSamplerProxy>
<hashTree>
  ${ASSERTION}
</hashTree>
```

---

#### 子模板：Query/表单参数（逐个展开）

```xml
<elementProp name="${PARAM_NAME}" elementType="HTTPArgument">
  <boolProp name="HTTPArgument.always_encode">true</boolProp>
  <stringProp name="Argument.name">${PARAM_NAME}</stringProp>
  <stringProp name="Argument.value">${PARAM_VALUE}</stringProp>
  <stringProp name="Argument.metadata">=</stringProp>
  <boolProp name="HTTPArgument.use_equals">true</boolProp>
</elementProp>
```

---

#### 子模板：请求头管理器（有自定义请求头时插入）

```xml
<HeaderManager guiclass="HeaderPanel" testclass="HeaderManager" testname="HTTP Header Manager" enabled="true">
  <collectionProp name="HeaderManager.headers">
    ${HEADERS}
  </collectionProp>
</HeaderManager>
<hashTree/>
```

每个请求头条目：
```xml
<elementProp name="${HEADER_NAME}" elementType="Header">
  <stringProp name="Header.name">${HEADER_NAME}</stringProp>
  <stringProp name="Header.value">${HEADER_VALUE}</stringProp>
</elementProp>
```

**注意**：有 JSON Body 时自动补充 `Content-Type: application/json`；有表单参数时自动补充 `Content-Type: application/x-www-form-urlencoded`，无需用户重复声明。

---

#### 子模板：响应状态码断言（默认插入每个采样器的 hashTree 内）

```xml
<ResponseAssertion guiclass="AssertionGui" testclass="ResponseAssertion" testname="Response Assertion" enabled="true">
  <collectionProp name="Assertion.test_strings">
    <stringProp name="49586">${EXPECTED_STATUS_CODE}</stringProp>
  </collectionProp>
  <stringProp name="Assertion.custom_message"></stringProp>
  <stringProp name="Assertion.test_field">Assertion.response_code</stringProp>
  <boolProp name="Assertion.assume_success">false</boolProp>
  <intProp name="Assertion.test_type">8</intProp>
</ResponseAssertion>
<hashTree/>
```

> **注意**：`${ASSERTION}` 占位符在有断言时替换为上述 XML；若不需要断言，采样器的 `<hashTree>` 块须写为 `<hashTree/>` 空标签，不可省略，否则 JMeter XML 结构不合法。

---

#### 多线程组模式（用户选 C 时）

在同一 TestPlan 的 `<hashTree>` 下，为每个接口各生成一个独立的 `<ThreadGroup>` + `<hashTree>` 块，每个线程组可独立配置线程数/Ramp-up/循环次数，单独询问。

---

### 第五步：XML 校验

文件写入完成后，使用 Bash 工具执行 Python XML 解析校验，确认生成的文件格式正确：

```bash
python -c "
import xml.etree.ElementTree as ET, sys
try:
    ET.parse(sys.argv[1])
    print('XML valid')
except ET.ParseError as e:
    print(f'XML error: {e}')
    sys.exit(1)
" "[filename].jmx" 2>&1
```

- **输出 `XML valid`**：继续输出反馈。
- **输出 XML 错误**：告知用户"生成的 JMX 文件 XML 格式有误"并显示错误位置，不输出正常反馈。

---

### 第六步：输出反馈

文件写入成功后，输出以下信息：

> **JMX 文件已生成：** `./[filename].jmx`
>
> **测试计划概览：**
> - 测试计划名称：[名称]
> - 线程数：N  |  Ramp-up：Xs  |  循环：Y 次
> - 接口列表：
>   1. [METHOD] [protocol]://[host]:[port][path]
>   2. ...
>
> **使用方式：**
> - GUI：JMeter → File → Open 导入此文件
> - 命令行：`jmeter -n -t [filename].jmx -l result.jtl`
>
> ℹ️ *Windows 用户若在命令行中使用反斜杠路径，请将文件名替换为实际路径，建议使用正斜杠 `/` 或将路径用引号括起。*

---

## 注意事项

- XML 特殊字符必须转义：`&` → `&amp;`，`<` → `&lt;`，`>` → `&gt;`，`"` → `&quot;`
- 端口为 443 且协议为 https，或端口为 80 且协议为 http 时，`port` 字段写空字符串（JMeter 默认行为）
- 所有元素的 `enabled="true"` 属性必须保留
- 文件编码统一 UTF-8
- 输出文件默认放在用户当前工作目录，用户指定路径时按用户要求
