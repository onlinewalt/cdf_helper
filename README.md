# CDF Helper · 报关清单生成器

根据备件来源 Excel，自动生成海关报关清单（Customs Declaration Form）。

## 特性

- **智能解析**：自动识别多种Excel表头格式（中文签收单/物料清单 + 英文Packing List）
- **船名识别**：从来源文件自动提取船名，并匹配中英文对照表
- **AI估算**：可选通过 DeepSeek API 估算缺失的重量/单价（带本地缓存）
- **双界面**：Web 界面 + 命令行模式
- **模板驱动**：基于Excel模板生成，保持与模板一致的样式

## 快速开始

```bash
pip install -r requirements.txt
python main.py  # 启动 Web 界面，自动打开浏览器
```

Web 界面访问：http://127.0.0.1:5000

## 使用说明

### Web 界面
1. 上传报关清单模板（.xlsx）
2. 上传一个或多个备件来源文件（.xls/.xlsx）
3. 填写船名（留空则自动识别）、目的港、日期
4. 可选：勾选「DeepSeek 智能填写」估算缺失的重量/单价
5. 点击「生成报关清单」，完成后下载

### 命令行
```bash
python main.py generate \
  --template template.xlsx \
  --source parts1.xls parts2.xlsx \
  --vessel "远怡湖 COSMERRY LAKE" \
  --port 上海 \
  --date 2026-08-20 \
  --ai --api-key YOUR_DEEPSEEK_KEY
```

### AI Key 配置
DeepSeek API Key 优先级：
1. `--api-key` 参数
2. 环境变量 `DEEPSEEK_API_KEY`
3. `config.json`（通过 Web 界面「保存 API Key」保存）

## 模板格式要求

模板文件 Sheet1 的固定布局：

| 位置 | 内容 |
|------|------|
| A1（合并 A1:G1） | 标题 `船名：<船名>` |
| 第2行 | 表头：序号 / 备件名称 / 数量 / 单位 / 重量(KG) / 单价/RMB / 金额RMB |
| 第3..N行 | 每条备件一行：序号`=ROW()-2`，金额`=C{r}*F{r}` |
| N+1行 | 合计行（黄色底纹，SUM公式） |

## 测试

```bash
python test_ai.py      # AI模块测试（模拟，不调用真实API）
python test_web.py     # Web 端到端测试
python test_packing.py # Packing List解析测试
```

运行所有测试：
```bash
python test_ai.py && python test_web.py && python test_packing.py
```

## 目录结构

```
├── main.py              # 入口：Web 界面 / CLI
├── webapp.py            # Flask Web 应用
├── templates/
│   ├── index.html       # 首页表单
│   └── result.html      # 生成结果页
├── static/style.css     # 页面样式
├── cdf_helper/
│   ├── parser.py        # 来源文件解析
│   ├── generator.py     # 模板填充/生成
│   ├── ai.py            # DeepSeek估算
│   └── config.py        # API Key配置
├── uploads/             # 上传文件目录（gitignore）
├── generated/           # 生成文件目录（gitignore）
├── config.json          # 本地配置（gitignore）
└── ai_cache.json        # DeepSeek结果缓存（gitignore）
```

## 依赖

- Python 3.9+
- Flask 3.0+
- openpyxl 3.1+
- xlrd 2.0+

安装依赖：
```bash
pip install -r requirements.txt
```

## 注意事项

- 重量/单价为 AI 估算值，正式申报前请人工核对
- `uploads/` 和 `generated/` 目录下的文件超过7天将自动清理
- Web 界面无 CSRF 保护，仅限本地使用
- xlrd 2.0+ 仅支持 `.xls`；`.xlsx` 使用 openpyxl 处理

## 许可证

本项目仅供内部使用。