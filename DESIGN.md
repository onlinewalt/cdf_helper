# CDF Helper 设计文档

> 版本：0.3.0
> 技术栈：Python 3.13 + Flask 3 + openpyxl + xlrd

---

## 1. 项目概述

CDF Helper（报关清单生成器）根据备件来源 Excel（.xls / .xlsx）自动生成海关报关清单（Customs Declaration Form，报关清单）。用户提供模板文件与一个或多个备件来源文件，程序自动定位表头、解析备件（名称/规格/数量/单位），并可选调用 DeepSeek 估算缺失的“重量 / 单价”，最终按模板样式输出可打印的 Excel 清单。

## 2. 目录结构

```
D:\CDF_helper\
├── main.py                  # 入口：无参启动 Web 界面 / generate 子命令（CLI）
├── webapp.py                # Flask Web 应用（上传、生成、下载）
├── requirements.txt         # 依赖：flask / openpyxl / xlrd
├── 启动CDF助手.bat          # Windows 双击启动（纯 ASCII，避免编码问题）
├── test_web.py              # Web 端到端测试（Flask test client）
├── test_ai.py               # DeepSeek 模块单元测试（mock 后端）
├── templates/
│   ├── index.html           # 首页表单（模板/来源/清单信息/DeepSeek）
│   └── result.html          # 生成结果页（统计 + 下载链接）
├── static/
│   └── style.css            # 页面样式
├── cdf_helper/
│   ├── __init__.py
│   ├── parser.py            # 来源文件解析（表头自动识别）
│   ├── generator.py         # 模板填充 / 生成工作簿
│   ├── ai.py                # DeepSeek 估算 + 本地缓存
│   └── config.py            # API Key 配置读写
├── uploads/                 # 上传的模板/来源文件（gitignore）
├── generated/               # 生成的报关清单（gitignore）
├── config.json              # 本地配置（API Key，gitignore）
├── ai_cache.json            # DeepSeek 结果缓存（gitignore）
└── 模板文件/来源文件示例       # 用户数据，不入库
```

## 3. 核心流程

```
来源文件(.xls/.xlsx) ──► parser.parse_sources ──► List[Part]
                                                    │
                      ┌─────────────────────────────┤
                      │ 可选：DeepSeek              │
                      │ ai.enrich_parts（只补缺失）  │
                      └─────────────────────────────┤
                                                    ▼
模板文件(.xlsx) ──► generator.generate ──► 生成清单(.xlsx)
                                                    │
                                    Web: result.html → /download/<file>
                                    CLI: 打印输出路径
```

## 4. 模板结构（模板文件布局约定）

生成器依赖模板的固定布局（`generator.TEMPLATE_SHEET = "Sheet1"`）：

| 位置 | 内容 |
| --- | --- |
| A1（合并 A1:G1） | 标题 `船名：<船名>`（16pt 微软雅黑加粗，居中） |
| 第 2 行 | 表头：序号 / 备件名称 / 数量 / 单位 / 重量(KG) / 单价/RMB / 金额RMB |
| 第 3..N 行 | 每条备件一行：序号 `=ROW()-2`，金额 `=C{r}*F{r}` |
| N+1 行 | 合计行：黄色底纹，`=SUM(C$3:INDEX(C:C,ROW()-1))` 等公式 |
| 模板预留数据容量 | 4 行（第 3..6 行），超出时自动 `insert_rows` 扩容 |

样式策略：从模板采样行的单元格复制 `_style`（字体、边框、对齐、底纹）应用到新行，保证生成结果与模板完全一致。

## 5. 模块设计

### 5.1 `cdf_helper/parser.py` — 来源解析

- 统一封装 `_Sheet` / `_Cell`，同时支持 xlrd（.xls）与 openpyxl（.xlsx）；`parse_source` 遍历**全部工作表**，每张表独立解析（多 sheet 的装箱单各自为一组备件）。
- **表头自动识别**（`_find_header_row`）：
  1. 逐行扫描，先按 `HEADER_ALIASES` 精确匹配（规范化后：去空格、小写）。"item" 不再映射为名称列（避免将 `ITEM` 序号列误识别为名称），改为通过 `description`/`particulars` 别名识别名称列。
  2. 未精确命中时按 `_CONTAINS_RULES` 包含匹配（优先级：名称 > 规格/型号 > 数量 > 单位 > 重量 > 单价）。新增 `description`→名称、`q'ty`→数量、`part no`→规格的包含规则。
  3. 当某行同时识别到 `name` 与 `qty` 即视为表头行。
- **多行单元格解析**（`_parse_multiline_row`）：当名称列单元格包含换行符且**至少 2 行以序号前缀**（`_SEQ_PREFIX_RE` 匹配 `\d+\s` 或 `✓/工 + 空格 + 数字`）开头时，自动按换行分割多行，逐行创建备件 —— 适配如 Sheet8 格式（列 A 含 `1  PLASTIC SHELL...\n2  1PLASTIC...`，列 F 含数量，列 H 含单位 PC/SET）。
  - 数量/单位各列换行按行号对齐；若单位列缺失或与数量列相同，`_find_unit_column` 自动扫描数据列寻找非数值短文本（如 PC/SET/MTR）。
  - 单值列在多行时自动重复。
  - 非多行格式的多行名称（如 `thremostatic\nexpansion valve`）仍作为单一备件处理，换行被清洗为空格。
- **序号清洗**（`_strip_seq`）：对单个单元格的名称执行以下清洗——
  1. 去除前导的复选框字符（`工 ✓ ☑` 等）；
  2. 去除前导序号 + 2 空格以上（如 `2    INTERMEDIATE RELAY` → `INTERMEDIATE RELAY`）；
  3. 去除尾随序号（如 `INTERMEDIATE RELAY   1` → `INTERMEDIATE RELAY`）。
- 数据行按列读入；缺少数量按 1 处理并回调 `warn()` 提示；`_to_number` 支持千分位/全角逗号/提取首个数（正则公底）；名称/规格/类型等文本通过 `_clean` 规范（折叠换行为空格）。
- **英文 Receipt/Packing List**（如远通海事，表头含 `Item/Quantity(Unit)/Particulars` 或 `Description`，或仅 `Item/Quantity(Unit)`，数量形如 `N PCE`，以 `** End of Listing **` 结尾）：表头行内找不到 `name+qty` 时自动回退到 `_parse_packing_sheet` 启发式解析——按行抓数量（`_QTY_RE`，仅识别英文单位，避免中文描述如"每条总长11米"误判）、从同行/后续行收集名称与 `Type:` 规格，直至下一数量行或清单末尾；`Part No / Serial No / Dwg.*` 列按表头标注自动排除，不混入名称；缺名称的条目记占位符 `(未填写名称)` 并告警；签名/网址/分隔线等页脚碎片会被 `_trim_footer`/`_clean_name` 剔除。`_find_packing_header` 现在也识别 `qty`/`q'ty` 而非仅 `quantity`。
- **页脚过滤**（`_SIGNATURE_KEYWORDS` + `_is_signature`）：扩展关键词表，新增英文页脚词（`signed`、`received`、`place of supply`、`date of supply`、`please check`）、中文页脚词（`交货地址`、`请核对`、`核对`、`日 期`）；匹配前统一折叠空白并小写，避免 `"place  of  Supply:"` 因双空格失效。
- **船名尾部清洗**（`_split_zh_en`）：英文船名末尾的非字母数字字符（如 `YINNIAN-` 中的连字符）会被剔除。
- 中文签收单/清单：数据区中 `签字/签收/盖章/供船` 等签名行或页脚文件名会被跳过，不会当成备件。
- `detect_vessel(paths)`：逐行匹配 `船名[:：]xxx`、`船名/单位` 右邻单元格、或 `Vessel Name` 行（含跨格值，优先取括号内中文船名）。
- **中英文船名对照表**：根目录 `中英文船名25-5-14.xls` 的 sheet 内按列对排列 `中文船名：/英文名：`（含 `英文名是拼音：` 列）。`load_vessel_names` 解析出 `(zh2en, en2zh)` 双向映射；`detect_vessel_pair` 从来源返回 `(chinese, english)`（`Vessel Name` 等行的 `COSMERRY LAKE(遠怡湖)` 或 `船名` 行拆中英文）；`bilingual_vessel` 合成 `中文 英文`（优先用英文名反查，规避来源繁体 vs 对照表简体的差异，`_zh_key` 做轻量繁→简转换）。
- 生成入口（webapp / main）在船名为空时先 `detect_vessel_pair`，再无结果回退 `detect_vessel`；识别到的船名一律经 `bilingual_vessel` 富化为 `中文 英文`（对照表内未收录则保持原样）。该对照表文件不会被当作备件来源/模板候选。
- 数据模型：

```python
@dataclass
class Part:
    name: str              # 名称 / 物资名称（必填）
    qty: float = 1.0       # 数量
    unit: str = "个"       # 单位
    type: Optional[str]    # 规格 / 型号
    weight: Optional[float]# 重量(KG)
    price: Optional[float] # 单价(RMB)
```

### 5.2 `cdf_helper/generator.py` — 生成

- `generate(...)`：加载模板 → 写标题 → 清除样例数据 → 按需 `insert_rows` 扩容 → 逐行写入（公式 序号/金额）→ 写合计行 → 清理模板遗留单元格（O7/P7）→ 保存。
- `_display_name`：`include_spec=True` 时拼 `名称 + 空格 + 规格`，否则仅名称。
- `sanitize_filename`：剔除 Windows 非法字符 `\/:*?"<>|` 等，用于输出文件名 `<船名>-<港>-报关清单-<日期>.xlsx`。

### 5.3 `cdf_helper/ai.py` — DeepSeek 智能估算

- 端点：`https://api.deepseek.com/chat/completions`，模型 `deepseek-chat`，`response_format={"type":"json_object"}`。
- **只填补缺失字段**（`weight is None or price is None`），不覆盖来源已有值。
- 分批（`BATCH_SIZE=50`）请求，避免 500 条物料产生 500 次调用。
- 缓存 `ai_cache.json`：key = `sha1(名称|规格)`，命中即免调用、免计费；失败项也会写缓存（存 None）防止重复请求。
- 容错：HTTP 429/5xx 重试一次（间隔 5s）；URL 错误/超时/解析失败返回 None 并计数，相关项保持空白。
- `_parse_json` 兼容模型偶尔输出 ` ```json ... ``` ` 包裹。
- 结果 `stats = {requested, filled, from_cache, errors}`。

### 5.3 `cdf_helper/translator.py` — 英文名称翻译

- **始终开启**：在 `parse_sources` 后、AI 估算前自动执行，将英文备件名称翻译为中文。
- **语言判断**：`_is_english` 检测名称中是否含有 CJK 字符——仅无中文的字符段（如 `INTERMEDIATE RELAY`）会被翻译；已含中文的名称原样保留。
- **API**：调用牛转翻译 API `POST https://api.niutrans.com/v2/text/translate`，参数 `from=en/to=zh/appId/timestamp/srcText/authStr`；`authStr` 为 MD5(sorted(params + apikey))。
- **凭据**：优先环境变量 `TRANSLATE_APP_ID` / `TRANSLATE_APIKEY`，其次 `config.json` 的 `trans_app_id` / `trans_apikey`；Web 表单可填并「保存到本地 config.json」。
- **缓存**：结果缓存在 `translate_cache.json`（key = `sha1(name)`），命中直接返回不再请求；失败结果（None）也写入缓存防止重复请求。
- **批量**：`BATCH_SIZE=20`，API 失败/超时 HTTP 429/5xx 自动重试一次（间隔 5s）。
- **容错**：凭据缺失或 API 抛错均为非致命，原英文名称原样保留，并通过 `warn()` 提示。

### 5.5 `cdf_helper/config.py` — 配置

- `config.json` 存 API Key 等本地配置。
- 优先级：环境变量 `DEEPSEEK_API_KEY` > `config.json`。
- `save_config()` 只更新传入字段，保留其他配置。

### 5.6 `webapp.py` — Flask Web

- `GET /`：首页，列出服务器根目录可选的模板与来源文件，回填配置中的 API Key。
- `POST /generate`：处理上传（`template_upload` / `sources_upload`，存入 `uploads/`）或服务器文件选择（`template_path` / `source_paths`）→ 解析 → 翻译英文名称（自动）→ 可选 DeepSeek → 生成到 `generated/` → 渲染 result 页。
- `GET /download/<file>`：仅允许从 `generated/` 目录内下载（防目录穿越）。
- 启动时清理 `uploads/`、`generated/` 中超过 7 天的旧文件（`_cleanup_old_files`）。
- 表单字段：`vessel`（留空自动识别）、`port`、`date`、`include_spec`、`use_ai`、`api_key`、`save_key`、`trans_app_id`、`trans_apikey`、`save_trans_key`。

### 5.7 `main.py` — 入口

- 无子命令：`_serve()` 自动挑选空闲端口（默认 5000 起），启动 Flask 并打开浏览器。
- `generate` 子命令：完整 CLI 流程，新增 `--ai` / `--api-key`。
- 交互兜底：未传 `--template/--source/--vessel` 时逐项询问。

## 6. DeepSeek 提示词设计

- System：强调角色（船舶备件采购/海关申报助理）、估算口径（重量 KG、单价 RMB、保守合理）、无法判断返回 null。
- User：逐条列出 `编号. 名称；规格；单位`，要求仅输出 `{"items":[{"id":1,"weight_kg":…,"unit_price":…}]}`。
- 参数：`temperature=0.2`（求稳定，避免估算漂移）。

## 7. 测试

| 文件 | 覆盖点 |
| --- | --- |
| `test_ai.py` | 只补缺失值；缓存命中不再调 API；API 失败保持空白；JSON 包裹兼容 |
| `test_web.py` | 首页渲染、服务器文件生成、下载完整性、上传路径、无来源报错、AI 流程、无 Key 报错 |
| `test_packing.py` | 英文 Packing List 多行/多 Sheet 解析、Item/Quantity(Unit)/Particulars 表头识别、`End of Listing` 结束标记 |
| `test_translator.py` | 英文名称识别、翻译执行、缓存命中、API 失败保留原文、缓存文件持久化 |

运行：`python test_ai.py && python test_translator.py && python test_web.py && python test_packing.py`

## 8. 已知限制与未来扩展

- 重量/单价为估算值，正式申报前需人工核对。
- 来源中同名不同规格（如 M22/M24 螺栓）取消“拼接规格”后会丢失区分信息，默认保持拼接。
- 表头识别依赖中英文别名表，遇到新表头格式需扩充 `HEADER_ALIASES` / `_CONTAINS_RULES`。
- 可选扩展：多页分页/打印区域设置、历史记录管理、批量任务队列、其他 LLM 提供商接入、表单重量/单价人工复核界面。