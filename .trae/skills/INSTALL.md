# 项目环境与 MCP 工具下载指南

## 1. 安装 uv（包管理工具）

### 方案 A：下载到项目本地（推荐）

```powershell
# 从 GitHub 下载 uv 的 Windows 二进制文件
# 访问 https://github.com/astral-sh/uv/releases/latest
# 下载 uv-x86_64-pc-windows-msvc.zip

# 解压后将 uv.exe、uvx.exe、uvw.exe 放入项目 .trae/ 目录
```

验证安装：

```powershell
.trae\uv --version
# 输出示例: uv 0.11.15 (3cffe97c2 2026-05-18 x86_64-pc-windows-msvc)
```

### 方案 B：通过官方脚本安装

```powershell
# 需要管理员权限，会自动添加到 PATH
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
```

---

## 2. 安装 Office-Word-MCP-Server

### 方式一：从 PyPI 安装（推荐）

```powershell
# 设置工具目录指向项目本地
$env:UV_TOOL_DIR = ".uv\tools"
$env:UV_LINK_MODE = "copy"

# 从 PyPI 安装
.trae\uv tool install office-word-mcp-server
```

### 方式二：从源码安装

```powershell
# 克隆仓库
git clone https://github.com/GongRzhe/Office-Word-MCP-Server.git
cd Office-Word-MCP-Server

# 安装到项目本地
$env:UV_TOOL_DIR = ".uv\tools"
$env:UV_LINK_MODE = "copy"
.trae\uv tool install --from . office-word-mcp-server
```

### 验证安装

```powershell
# 检查工具是否安装成功
Get-ChildItem .uv\tools\office-word-mcp-server\Scripts\word_mcp_server.exe

# 测试启动
.trae\uv tool run office-word-mcp-server --help
# 应显示 FastMCP 服务器启动信息
```
---

## 3. 卸载旧工具

### 卸载旧 docx-mcp

```powershell
# 删除旧工具环境
Remove-Item -Recurse -Force ".uv\tools\docx-mcp"
```

### 卸载 Office-Word-MCP-Server

```powershell
# 从 .uv/tools 中删除
Remove-Item -Recurse -Force ".uv\tools\office-word-mcp-server"
```

### 清理 PyPI 缓存

```powershell
# 查看缓存位置
.trae\uv cache dir

# 清理所有缓存
.trae\uv cache clean
```

---

## 4. 依赖包说明

| 包名 | 版本 | 用途 |
|:---|:---:|:---|
| `office-word-mcp-server` | 1.1.11+ | Word 文档 MCP 服务器（填表、排版、格式化） |
| `python-docx` | 1.1.2+ | Word 文档读写（MCP 底层依赖） |
| `uv` | 0.11.15+ | Python 包管理器（项目本地工具） |

---

## 5. 目录结构参考

安装完成后，项目目录结构如下：

```
Machine Learning\
  ├── .trae\
  │   ├── uv.exe          ← uv 主程序
  │   ├── uvx.exe         ← uvx（临时运行工具）
  │   ├── uvw.exe         ← uvw（等待工具）
  │   ├── uv.bat          ← uv 包装脚本
  │   ├── word-mcp.bat    ← Word MCP 包装脚本
  │   └── skills\
  │       ├── SKILL.md    ← 实验流程 skill
  │       └── INSTALL.md      ← 安装说明
  ├── .uv\
  │   ├── tools\
  │   │   └── office-word-mcp-server\  ← MCP 服务器环境
  │   └── cache\          ← 缓存
  └── .gitignore          ← 已忽略 .trae/ 和 .uv/
```
