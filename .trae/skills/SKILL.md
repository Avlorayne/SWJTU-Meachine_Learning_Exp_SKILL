---
name: "swjtu-ml-experiment"
description: "西南交通大学机器学习实验全流程助手。从 docx 实验模板读取要求 → 生成 md 工作区 → 实验编码验证 → 在 md 上撰写报告 → 用户确认后填入 Word → 打包提交。Invoke when user asks to conduct a machine learning experiment."
---

# 西南交通大学-机器学习实验

## 概述

每一次对话都是一个独立的机器学习实验任务。用户提供原始实验文件夹（如 `\第X次实验`），其中包含：

- 实验报告模板 `机器学习_第i次实验报告模板.docx`（包含实验要求与格式）
- CSV 格式的训练集与测试集
- 可能存在的无用文件 `2026机器学习实验要求.docx`

你将原始文件夹中的内容组织到标准目录结构 `\Exp_i\` 下，完成全部实验工作。**整个过程中，`exp_i.md`** **是 Agent 可直接编辑的工作区域**，所有实验原理、代码分析、结果数据、误差分析、结论等内容先在 md 上撰写和完善。**Word 文档仅在最终写入阶段被填充**，不做中间编辑。

## 初始步骤：获取项目根目录与配置

在开始任何操作前，先获取当前项目根目录，记为 `$PROJECT_ROOT`。后续所有路径均以此为基础进行拼接。

```powershell
$PROJECT_ROOT = (Get-Location).Path
```

### MCP 服务器配置与安装

本实验流程需要使用 MCP Word 工具操作 docx 文件。首次运行时需要安装 `office-word-mcp-server`，并配置 MCP 服务器。

#### Step 1: 安装 MCP 服务器

在 PowerShell 中执行以下命令，`uvx` 会自动从 PyPI 下载并缓存 `office-word-mcp-server` 包：

```powershell
& "$PROJECT_ROOT\.trae\uvx.exe" --from office-word-mcp-server word_mcp_server --help
```

下载完成后即可正常使用，后续启动将直接使用缓存，无需重复下载。

#### Step 2: 配置 Trae MCP 服务器

在 Trae IDE 的 **设置 → MCP 服务器 → 添加** 中，选择"添加手动"，将以下 JSON 粘贴到输入框中（请将 `$PROJECT_ROOT` 替换为实际路径）：

```json
{
  "mcpServers": {
    "word-mcp": {
      "command": "$PROJECT_ROOT\\.trae\\word-mcp.bat",
      "args": []
    }
  }
}
```
  
### 学号与姓名

读取 `.trae/student_info.json`，检查 `student_id` 和 `student_name`。

- 若为空：向用户询问学号和姓名，写入该文件。
- 若已填写：直接使用。

此后所有需要姓名学号的地方（填写 Word 报告、打包文件）均从此文件读取，不再重复询问。

## 初始清理

检查用户的原始实验文件夹中是否存在无用文件 `2026机器学习实验要求.docx`，若存在则删除：

```powershell
Remove-Item -Path "$PROJECT_ROOT\第X次实验\2026机器学习实验要求.docx" -ErrorAction SilentlyContinue
```

***

## 工作流程

### Phase 1: 创建目录、移动资源、生成 md 工作区

#### Step 1.1: 读取 docx 模板获取实验要求

使用 MCP Word 工具读取原始实验文件夹中的 `机器学习_第i次实验报告模板.docx`：

```
mcp_word-mcp_get_document_info(filename)
mcp_word-mcp_get_document_outline(filename)
mcp_word-mcp_get_document_text(filename)
```

获取：文档段落结构、表格布局、实验任务描述、模型要求、评估指标等信息。

#### Step 1.2: 创建标准目录结构

```
$PROJECT_ROOT\Exp_i\
  ├── scr\          # 代码与数据集
  └── report\       # 实验报告与图表
```

#### Step 1.3: 移动实验资源到标准目录

- CSV 数据集 → `\Exp_i\scr\`
- 实验报告模板 docx → `\Exp_i\report\`（保留原文件名，**此后不再直接编辑此 docx**）

#### Step 1.4: 从 docx 提取代码框架

使用 `mcp_word-mcp_get_document_text` 获取完整文本，从返回内容中识别代码块（等宽字体段落），提取后创建标准命名的代码文件 `\Exp_i\scr\experiment_0i.py`。

#### Step 1.5: \*\* 生成 md 工作区文档 `exp_i.md` \*\*

\*\* 这一步是核心：\*\* 将 docx 模板的结构和内容转换为 `\Exp_i\report\exp_i.md`，作为本次实验的**可编辑工作文档**。包含：

- 实验名称与目的
- 实验内容描述
- **待填写的实验原理**（占位，后续 Phase 5 完善）
- **待填写的实验结果表格**（占位，Phase 3 运行后填充数据）

`exp_i.md` 是本次实验唯一需要 Agent 直接编辑修改的文档，后续所有内容都在此文件中撰写和完善。

***

### Phase 2: 数据准备与代码实现

1. 检查虚拟环境 `\venv` 是否可用，有缺失则补充包。
2. 使用 `\Exp_i\scr\` 下的训练集与测试集，完成数据预处理、特征工程及模型训练。
3. 在 `experiment_0i.py` 基础上补全所有缺失功能。**不得改变已有代码结构与函数接口**。除了代码中原有的包，**不得导入和使用其它的包**。
4. 全程在虚拟环境 `\venv` 中运行代码；导出依赖至 `\Exp_i\requirements.txt`。
5. 所有代码注释使用中文。

***

### Phase 3: 端到端验证

1. 运行代码，使用测试集进行验证。
2. 计算所有关键评估指标，确保达到基准线。
3. 若需要绘图（pyplot），将生成的图表保存至 `\Exp_i\report\` 目录下。

***

### Phase 4: 迭代改进

1. 检查依赖包，对实验结果进行验证。
2. 根据效果改进代码，若缺失则安装机器学习相关依赖包。

***

### Phase 5: 在 md 工作区上撰写完整实验报告

这是**在 md 上撰写报告**的阶段。将 Phase 3 的实验结果和代码写入 `\Exp_i\report\exp_i.md`，完善各个章节：

1. **实验原理**：撰写算法原理描述（朴素贝叶斯 / 逻辑回归 / 决策树等）
2. **实验结果表格**：将运行输出的指标数值、模型参数填入 md 的表格中
3. **代码段落**：关键代码片段（函数实现部分）写入 md
4. **结果分析**：指标对比、图表解读
5. **误差分析**：分析模型误差来源、改进方向
6. **实验结论**：总结实验发现

**此阶段 exp\_i.md 的内容已完整，格式为 Markdown，Agent 可直接编辑修改。**

***

### Phase 6: 用户确认

报告撰写完成后，主动询问用户是否对内容满意：

- 将 `exp_i.md` 的关键内容（实验结果、分析、结论等）呈现给用户
- 用户确认前，Agent 可根据用户反馈继续修改 `exp_i.md`
- **用户确认满意后**，进入 Phase 7

***

### Phase 7: 从 md 填写 Word 模板

用户确认满意后，开始将 `exp_i.md` 中的内容填充到 Word 模板。在编写Word过程中， **不可以更改文件原有的格式或内容** 。

#### Step 7.1: 读取 md 内容

从 `exp_i.md` 中提取需要填入 Word 的结构化数据（学号姓名从 `student_info.json` 读取）：

- 学生信息（姓名、学号）
- 实验原理文本
- 实验数据表格
- 代码片段
- 结果分析、结论等文本

#### Step 7.2: 备份模板

```
mcp_word-mcp_copy_document(source_filename, destination_filename)
```

#### Step 7.3: 编写填充脚本

复制项目预置的 `docx_utils.py` 到实验目录，编写 `fill_report.py` 将数据填入 Word。

```powershell
Copy-Item ".trae/skills/exp-report-filler/docx_utils.py" "Exp_i/scr/" -Force
```

`docx_utils.py` 提供的主要函数见下方。

运行脚本：

```bash
python fill_report.py
```

#### Step 7.4: 验证输出

检查表格填充是否正确，段落内容是否完整。

#### Step 7.5: 插入图表

若实验生成了图表，使用 MCP 工具插入到 docx 中：

```
mcp_word-mcp_add_picture(filename, image_path, width)
```

***

### Phase 8: 打包提交文件

从 `student_info.json` 读取学号姓名，将三个文件打包为 zip：

| 文件        | 来源                                  | 命名格式                     |
| --------- | ----------------------------------- | ------------------------ |
| Word 实验报告 | `\Exp_i\report\机器学习_第i次实验报告模板.docx` | `{学号}+{姓名}+第{i}次实验.docx` |
| 实验代码      | `\Exp_i\scr\experiment_0i.py`       | `{学号}+{姓名}+第{i}次实验.py`   |
| PDF 实验报告  | docx 转换                             | `{学号}+{姓名}+第{i}次实验.pdf`  |

```python
# 创建 package_files.py 运行打包（含 docx→PDF 转换）
```

若 `pywin32` 未安装：

```bash
pip install pywin32
```

zip 输出到 `\Exp_i\report\`。

***

### Phase 9: 清理临时文件

清理临时脚本，确保目录只保留核心产物：

```
Exp_i/
├── scr/
│   ├── experiment_0i.py        # 实验代码
│   ├── experiment_0i_*.csv     # 数据集
├── report/
│   ├── 机器学习_第i次实验报告模板.docx  # 已填充的报告模板
│   ├── {学号}+{姓名}+第{i}次实验.zip    # 提交包
│   ├── 机器学习_第i次实验报告模板_备份.docx
│   └── exp_i.md                        # 实验工作文档
└── requirements.txt
```

清理：`docx_utils.py`、`fill_report.py`、`package_files.py`、`temp_submit\`。

***

## docx\_utils.py 函数参考

| 函数                                                           | 功能                  |
| ------------------------------------------------------------ | ------------------- |
| `fill_table(doc, table_index, data, start_row, ...)`         | 批量填充表格              |
| `set_code_font(doc, start_idx, end_idx)`                     | 设置代码段 Consolas 11pt |
| `insert_text(doc, para_index, text, ...)`                    | 替换指定段落内容            |
| `append_text(paragraph, text, ...)`                          | 段落末尾追加文本            |
| `add_picture_to_paragraph(doc, para_index, image_path, ...)` | 段落中插入图片             |
| `validate_tables(doc, filename)`                             | 验证表格结构              |
| `backup_template(docx_path)`                                 | 备份模板                |

## 填写内容要求

- **学生信息**：姓名、学号（填入学生信息表）
- **实验代码**：核心代码片段（Consolas 11pt）
- **实验结果**：评估指标数值、模型参数表格
- **结果图表**：曲线图等
- **实验原理、结果分析、结论**：按学术规范撰写

## 规则

- 不修改模板中已有的文字内容，只填充空白单元格或在末尾追加
- 表格数据保留原表头样式
- 代码一律用 Consolas 字体 11pt
- 文本格式（字体和大小等）与附近文字格式相同，否则默认中文正文用宋体，中文标题用黑体，英文用 Times New Roman
- 实验指导文档 `exp_i.md` 中的既有内容与格式不得随意删除
- 所有代码注释使用中文

