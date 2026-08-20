# CDF Helper 设计文档

> 版本：0.2.0（对应提交 `676bd7c`）
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

- 统一封装 `_Sheet` / `_Cell`，同时支持 xlrd（.xls）与 openpyxl（.xlsx）。
- **表头自动识别**（`_find_header_row`）：
  1. 逐行扫描，先按 `HEADER_ALIASES` 精确匹配（规范化后：去空格、小写）。
  2. 未精确命中时按 `_CONTAINS_RULES` 包含匹配（优先级：名称 > 规格/型号 > 数量 > 单位 > 重量 > 单价）。
  3. 当某行同时识别到 `name` 与 `qty` 即视为表头行。
- 数据行按列读入；缺少数量按 1 处理并回调 `warn()` 提示；`_to_number` 支持千分位/全角逗号/提取首个数（正则兜底）。
- `detect_vessel(paths)`：正则匹配 `船名[:：]xxx` 或 `船名/单位` 右邻单元格，返回船名。
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

### 5.4 `cdf_helper/config.py` — 配置

- `config.json` 存 API Key 等本地配置。
- 优先级：环境变量 `DEEPSEEK_API_KEY` > `config.json`。
- `save_config()` 只更新传入字段，保留其他配置。

### 5.5 `webapp.py` — Flask Web

- `GET /`：首页，列出服务器根目录可选的模板与来源文件，回填配置中的 API Key。
- `POST /generate`：处理上传（`template_upload` / `sources_upload`，存入 `uploads/`）或服务器文件选择（`template_path` / `source_paths`）→ 解析 → 可选 DeepSeek → 生成到 `generated/` → 渲染 result 页。
- `GET /download/<file>`：仅允许从 `generated/` 目录内下载（防目录穿越）。
- 启动时清理 `uploads/`、`generated/` 中超过 7 天的旧文件（`_cleanup_old_files`）。
- 表单字段：`vessel`（留空自动识别）、`port`、`date`、`include_spec`、`use_ai`、`api_key`、`save_key`。

### 5.6 `main.py` — 入口

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

运行：`python test_ai.py && python test_web.py`

## 8. 已知限制与未来扩展

- 重量/单价为估算值，正式申报前需人工核对。
- 来源中同名不同规格（如 M22/M24 螺栓）取消“拼接规格”后会丢失区分信息，默认保持拼接。
- 表头识别依赖中英文别名表，遇到新表头格式需扩充 `HEADER_ALIASES` / `_CONTAINS_RULES`。
- 可选扩展：多页分页/打印区域设置、历史记录管理、批量任务队列、其他 LLM 提供商接入、表单重量/单价人工复核界面。