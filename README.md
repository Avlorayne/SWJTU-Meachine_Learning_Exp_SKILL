# 西南交通大学 机器学习实验

本项目用于西南交通大学机器学习课程的实验任务，提供了完整的实验工作流和自动化工具。

## 前置条件

- **Python 3.10+** — 所有实验代码运行环境
- **Microsoft Word** — 实验报告模板填写和 PDF 导出
- **[Trae IDE](https://www.trae.ai/)**（推荐） — 自动执行实验工作流（可选）
其实你用deepseek-tui这种也没有什么问题，要是有问题就向我提 issue 或 rp

## 快速开始

### 1. 克隆项目

```bash
git clone <仓库地址>
cd Machine Learning
```

### 2. 配置 MCP 服务器

运行以下命令，将生成 MCP 配置 JSON 文件，将其复制到 TRAE 或其它 MCP 服务器的配置中。
```bashhell
# 获取当前文件夹下的 uvx.exe 并生成 MCP 配置 JSON
$uvxPath = (Get-ChildItem -Path ".\" -Filter "uvx.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1).FullName; if ($uvxPath) { $uvxPath = $uvxPath -replace '\\', '\\'; @"
{
  "mcpServers": {
    "office-word": {
      "command": "$uvxPath",
      "args": [
        "--from",
        "office-word-mcp-server",
        "word_mcp_server"
      ]
    }
  }
}
"@ } else { Write-Host "当前文件夹未找到 uvx.exe" }
```


例如（这个不要用，是我本地的配置）：
```json
{
  "mcpServers": {
    "office-word": {
      "command": "F:\\Project\\SWJTU-Meachine_Learning_Exp_SKILL\\.trae\\uvx.exe",
      "args": [
        "--from",
        "office-word-mcp-server",
        "word_mcp_server"
      ]
    }
  }
}
```

如果你的 Agent 在过程中忽略了 MCP，那我没招了，这个 AI 模型就是个 **。

### 3. 目录结构说明

```
Machine Learning/
├── .trae/skills/       # Trae IDE 自动化技能
│   ├── 西南交通大学-机器学习实验.skill.md   # 实验全流程自动化 skill
│   └── exp-report-filler/
│       └── docx_utils.py                    # Word 模板填充工具库
├── .uv/                # MCP 服务器配置
└── README.md
```

每个实验的标准目录结构：

```
Exp_i/
├── scr/
│   ├── experiment_0i.py        # 实验代码
│   └── experiment_0i_*.csv     # 数据集（训练集/测试集）
├── report/
│   ├── exp_i.md                # 实验指导文档
│   └── 机器学习_第i次实验报告模板.docx  # 实验报告模板
└── requirements.txt            # 依赖清单
```

## 使用方式

把你从老师那里接收到的实验文件夹直接丢给 Agent，让它干活就可以了。

### 首次运行

Prompt: 根据项目skill完成`第i次实验目录`

1. Agent 会向你询问相关信息，请留意回复。所有的信息只会记录在你的本地，请放心。
2. 首次运行时，MCP 服务器会从网络下载，可能需要一些时间。
3. Agent 会自动安装实验依赖，无需手动操作。

### 方式一：Trae IDE 自动执行（推荐）

在 Trae IDE 中打开本项目，Trae 会自动加载 [SKILL.md](.trae/skills/SKILL.md) 中的工作流。

详细的自动化流程见 [skill 文档](.trae/skills/SKILL.md)，主要步骤包括：

1. **Phase 1** — 读取 docx 模板 → 生成 `exp_i.md` 工作区（Agent 可编辑）
2. **Phase 2-4** — 补全代码 → 运行验证 → 调优改进
3. **Phase 5** — 在 `exp_i.md` 上撰写完整实验报告（原理、结果、分析、结论）
4. **Phase 6** — 用户确认报告内容，不满意则继续修改
5. **Phase 7** — 用户确认后，从 `exp_i.md` 填入 Word 模板
6. **Phase 8** — 打包为 `{学号}+{姓名}+第{i}次实验.zip`
7. **Phase 9** — 清理临时文件

### 方式二：手动执行

直接运行单个实验的代码：

```bash
cd Exp_7/scr
python experiment_07.py
```

### 方式三：手动填充 Word 报告

```bash
# 复制工具库
Copy-Item ".trae/skills/exp-report-filler/docx_utils.py" "Exp_i/scr/" -Force

# 编写 fill_report.py 后运行
python fill_report.py
```

工具库 [docx_utils.py](.trae/skills/exp-report-filler/docx_utils.py) 提供以下函数：

| 函数 | 功能 |
|------|------|
| `fill_table(doc, table_index, data, start_row, ...)` | 批量填充表格 |
| `set_code_font(doc, start_idx, end_idx)` | 设置代码段为 Consolas 11pt |
| `insert_text(doc, para_index, text, ...)` | 替换指定段落内容 |
| `append_text(paragraph, text, ...)` | 在段落末尾追加文本 |
| `add_picture_to_paragraph(doc, para_index, image_path, ...)` | 在段落中插入图片 |
| `validate_tables(doc, filename)` | 验证文档表格结构 |
| `backup_template(docx_path)` | 备份模板文件 |
| `docx_to_pdf(docx_path, pdf_path)` | docx 转 PDF |

## 实验报告提交

所有实验任务完成后，最终提交产物为 zip 压缩包，包含三个文件：

- `{学号}+{姓名}+第{i}次实验.docx` — 已填充的实验报告
- `{学号}+{姓名}+第{i}次实验.py` — 实验代码
- `{学号}+{姓名}+第{i}次实验.pdf` — 实验报告 PDF 版本

### 学号与姓名配置

学号和姓名记录在 `.trae/student_info.json`，初始为空。首次使用 Trae 自动化流程时 agent 会询问并写入，后续实验自动读取，无需重复输入。

```json
{"student_id": "", "student_name": ""}
```

该文件已加入 `.gitignore`，不会提交到仓库。

## 注意事项

- 实验代码中**不得导入除原有包以外的额外包**
- 代码注释请使用中文
- 模板已有文字内容不得修改，仅填充空白单元格或在末尾追加
- 代码段字体统一为 Consolas 11pt
- 中文正文用宋体，中文标题用黑体，英文正文用 Times New Roman
