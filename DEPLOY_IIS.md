# IIS 部署指南（Windows Server）

## 前置要求

### 1. 服务器角色
- **Web Server (IIS)** — 包含 **CGI** 功能
- **Application Request Routing (ARR)** — 可选，用于反向代理
- **FastCGI Module** — IIS 安装 Web 服务器角色时默认包含

通过服务器管理器安装：
```
管理 → 添加角色和功能 → 服务器角色 → Web 服务器 (IIS)
  → 安全性 → 身份验证（按需启用）
  → 应用程序开发 → CGI
```

### 2. Python 安装
下载并安装 **Python 3.9+**（建议 3.13）：
- 下载地址：https://www.python.org/downloads/
- **勾选 "Add Python to PATH"**（添加到 PATH）
- 安装路径建议：`C:\Python313\`

### 3. 安装依赖
在部署文件夹中（包含 `web.config` 的目录）运行：
```bat
setup_iis.bat
```
或手动执行：
```bat
pip install -r requirements.txt
pip install wfastcgi
python -m wfastcgi-enable
```

## 部署步骤

### 1. 复制文件
将整个项目文件夹复制到服务器，例如：
```
C:\inetpub\wwwroot\cdf_helper\
```

### 2. 配置 web.config
编辑 `web.config`，确保 `scriptProcessor` 中的 Python 路径正确。运行 `setup_iis.bat` 会自动检测并更新路径：
```xml
<scriptProcessor="C:\Python313\python.exe|C:\Python313\lib\site-packages\wfastcgi.py" />
```

如果 `setup_iis.bat` 未能自动更新，请手动修改。

### 3. 创建 IIS 站点
1. 打开 **IIS 管理器**
2. 右击 **网站** → **添加网站**
   - 站点名称：`CDF Helper`
   - 物理路径：`C:\inetpub\wwwroot\cdf_helper\`
   - 绑定：`http` / 端口 `80`（或自定义端口）
3. **应用程序池**设置：
   - 右击站点 → 管理应用程序池
   - 选中对应的应用程序池 → 高级设置
   - **.NET CLR 版本**：无（None）
   - **托管管道模式**：经典（Classic）

### 4. 文件夹权限
确保 IIS 有以下文件夹的 **写入** 权限：
```
C:\inetpub\wwwroot\cdf_helper\uploads\
C:\inetpub\wwwroot\cdf_helper\generated\
C:\inetpub\wwwroot\cdf_helper\logs\
```

添加 `apphost.config` 中的应用程序池身份（通常是 `IIS APPPOOL\<应用程序池名称>`）到这些文件夹的写入权限。

### 5. 环境变量（可选）
如需配置 DeepSeek API Key 或翻译 API Key：
- 设置系统环境变量 `DEEPSEEK_API_KEY` 和 `TRANSLATE_APP_ID` / `TRANSLATE_APIKEY`
- 或在 `config.json` 中配置（首次在 Web 界面保存即可）

### 6. 访问
浏览到 `http://your-server/` 即可访问 CDF Helper。

## 故障排除

| 问题 | 解决方法 |
|------|----------|
| 500 错误 | 检查 `logs\wfastcgi.log` 日志 |
| wfastcgi 未注册 | 以管理员身份运行 `python -m wfastcgi-enable` |
| 权限错误 | 确保应用程序池身份对 `uploads/` 和 `generated/` 有写入权限 |
| 模块导入错误 | 确认 `pip install -r requirements.txt` 已在服务器 Python 中执行 |
| 中文乱码 | 确保 Python 安装路径中不包含中文字符 |

## 文件说明

| 文件 | 用途 |
|------|------|
| `web.config` | IIS wfastcgi 配置 |
| `setup_iis.bat` | 一键安装依赖 + 注册 wfastcgi + 更新 web.config |
| `webapp.py` | Flask 应用入口（`WSGI_HANDLER=webapp.app`） |
| `requirements.txt` | Python 依赖 |
