# Doc Processor

一个用于处理 MS Word (.docx) 文档的综合工具，专门设计用于英语教案生成。支持结构解析、模板清理、内容生成和文档重构。

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 📋 功能特点

- **🔍 结构解析**: 提取 Word 文档的层次化 JSON 表示（章节、段落、表格）
- **🧹 模板清理**: 创建空白版本的文档模板
- **📝 内容生成**: 基于解析结构和用户主题智能生成内容
- **🏗️ 文档构建**: 将生成的内容注入模板生成最终文档
- **📚 题库集成**: 支持本地题库索引和按需加载，自动标注题目来源
- **🎯 专为教案设计**: 针对英语教学场景优化的模板和内容生成逻辑

## 🚀 快速开始

### 环境要求

- Python 3.8+
- python-docx 库

### 安装

```bash
# 克隆项目
git clone https://github.com/xielk/doc-processor.git
cd doc-processor

# 安装依赖
pip install python-docx
```

### 基本使用流程

一个完整的教案生成工作流包含以下步骤：

#### 1️⃣ 解析原始模板

```bash
python scripts/parser.py input.docx > structure.json
```

这会提取文档结构并保存为 JSON 格式。

#### 2️⃣ 创建清理模板

```bash
python scripts/cleaner.py input.docx template_clean.docx
```

生成一个空白模板，保留格式但清空内容。

#### 3️⃣ 生成内容（关键步骤）

内容生成需要结合学生信息和教学目标。以下是标准流程：

**A. 准备学生信息**
- 学生姓名：[学生姓名]
- 年级：初中三年级
- 当前水平：[区域]一模 98 分
- 目标分数：130 分
- 薄弱点：阅读 C/D 篇、写作
- 上课日期：2025-03-05

**B. 搜索题库（自动）**

系统会自动从本地题库搜索相关真题：

```python
from scripts.searcher import search_question_bank

# 搜索阅读材料（优先嘉定区2025年）
results, passages = search_question_bank(
    topic="阅读B篇",
    district="嘉定",
    year="2025"
)

# 搜索语法题目
results, questions = search_question_bank(
    topic="非谓语动词",
    district="嘉定", 
    year="2024"
)
```

**C. 创建内容 JSON**

基于模板结构和题库结果，创建 `content.json`：

```json
{
  "p_1": "[培训机构]英语学科",
  "p_2": "[学生姓名]专属辅导讲义",
  "t_1": [
    ["学生姓名", "年级", "上课日期", "上课时间", "时长"],
    ["[学生姓名]", "初中三年级", "2025-03-05", "10:00:00", "2h"],
    ["任课教师", "班主任", "教学组长", "教学主管", "学生签字"],
    ["[教师名]", "[班主任名]", "[教学组长]", "[教学主管]", ""]
  ],
  "p_32": "【学员诊断】\n[学生姓名]同学...",
  ...
}
```

#### 4️⃣ 构建最终文档

```bash
python scripts/builder.py template_clean.docx content.json output.docx
```

## 📖 详细使用教程

### 场景示例：生成一篇阅读 C 篇突破教案

#### 步骤 1：准备模板

确保你有一个 Word 格式的教案模板。模板应包含表格结构：

- 基本信息表（学生姓名、年级、日期等）
- 教学内容表（教学目标、重难点等）
- 教学过程表（知识详解、真题在线等）

**⚠️ 重要提示**：
- 避免使用 `/tmp/` 目录保存模板（文件可能被清理）
- 建议保存到用户目录：`~/Documents/templates/`

#### 步骤 2：解析模板结构

```bash
python scripts/parser.py "~/Documents/templates/lesson_template.docx" > structure.json
```

查看生成的 `structure.json`，了解模板的段落 ID 和表格结构。

#### 步骤 3：搜索相关题目

```bash
# 确保题库索引已创建
python scripts/indexer.py

# 在 Python 中搜索题目
python << 'PYEOF'
from scripts.searcher import search_question_bank

# 搜索特定区域 C 篇阅读
results, passages = search_question_bank(
    topic="C篇",
    district="[区域]",
    year="2025",
    max_docs=3
)

for p in passages[:3]:
    print(f"来源: {p['source']}")
    print(f"内容: {p['content'][:200]}...")
    print()
PYEOF
```

#### 步骤 4：创建内容

基于搜索到的题目和学员信息，创建 `content.json`：

```json
{
  "p_1": "[培训机构]英语学科",
  "p_2": "[学生姓名]专属辅导讲义",
  "t_1": [
    ["学生姓名", "年级", "上课日期", "上课时间", "时长"],
    ["[学生姓名]", "初中三年级", "2025-03-05", "10:00:00", "2h"],
    ["任课教师", "班主任", "教学组长", "教学主管", "学生签字"],
    ["[教师名]", "[班主任名]", "[教学组长]", "[教学主管]", ""]
  ],
  "t_2": [
    ["教学内容", "阅读理解C篇首字母填空突破"],
    ["教学目标", "1. 掌握C篇首字母填空三步法..."],
    ...
  ],
  "p_32": "【学员诊断与提分策略】\n\n[学生姓名]同学现状分析：...",
  "p_33": "【C篇首字母填空真题精讲】\n\nPassage 1（[年份] [区域]一模）...",
  ...
}
```

#### 步骤 5：生成文档

```bash
python scripts/builder.py \
    "~/Documents/templates/lesson_template_clean.docx" \
    "content.json" \
    "~/Documents/output/output_C篇突破.docx"
```

## 📁 项目结构

```
doc-processor/
├── SKILL.md                    # 技能说明文档（强制阅读）
├── PROMPT.md                   # 对话提示词
├── README.md                   # 本文件
├── scripts/                    # Python 脚本
│   ├── parser.py              # 文档结构解析器
│   ├── cleaner.py             # 模板清理器
│   ├── builder.py             # 文档构建器
│   ├── indexer.py             # 题库索引生成器
│   ├── searcher.py            # 题库搜索器
│   ├── smart_builder.py       # 智能构建器（含异常处理）
│   └── generator.py           # 内容生成器（示例）
└── [generated files]          # 生成的文件
    ├── structure.json         # 解析的结构
    ├── content.json           # 生成的内容
    └── output.docx            # 最终文档
```

## 🎯 核心概念

### 1. 结构解析 (Parser)

提取 Word 文档的层次结构：

- **段落 (p_X)**: 普通段落，按顺序编号
- **表格 (t_X)**: 表格结构，按顺序编号
- **章节 (sec_X)**: 标题段落

### 2. 内容映射

`content.json` 使用 ID 映射到文档位置：

```json
{
  "p_1": "第一段内容",
  "p_2": "第二段内容",
  "t_1": [
    ["表头1", "表头2"],
    ["数据1", "数据2"]
  ]
}
```

### 3. 题库集成

**索引系统**：
- 首次使用需创建索引：`python scripts/indexer.py`
- 索引位置：`[你的题库路径]/index.json`

**搜索策略**：
- 优先搜索学员所在区（如[区域名]）
- 优先搜索最新年份（2025 > 2024 > 2023）
- 限制加载文件数（3-5个），控制 Token 消耗

**来源标注**：
- 格式：`([年份] [区域]一模)`、`([年份] [区域]二模)`
- 每道题目必须标注来源

## ⚠️ 常见问题

### Q1: 模板文件找不到（/tmp/ 目录文件丢失）

**问题**：在新 session 中运行时报错 `FileNotFoundError: /tmp/template.docx`

**原因**：`/tmp/` 目录下的文件在 session 结束后会被清理

**解决方案**：

1. **将模板移动到永久目录**：
```bash
cp "/tmp/template.docx" "~/Documents/templates/"
```

2. **使用绝对路径**：
```bash
python scripts/builder.py \
    "~/Documents/templates/template.docx" \
    "content.json" \
    "output.docx"
```

3. **重新上传模板**：如果文件已丢失，需要重新上传

### Q2: 生成的文档格式错乱

**可能原因**：
- 模板文件损坏
- content.json 中的 ID 与模板不匹配
- 使用了错误的模板版本

**解决方法**：
1. 重新解析模板生成 structure.json
2. 检查 content.json 的 ID 是否正确
3. 确保使用清理后的模板（clean template）

### Q3: 题库搜索不到题目

**可能原因**：
- 索引未创建
- 搜索关键词不匹配
- 题库中没有该区/年份的题目

**解决方法**：
```bash
# 重新创建索引
python scripts/indexer.py

# 扩大搜索范围（不限定年份）
results, questions = search_question_bank(
    topic="非谓语",
    district="[区域]",
    year=None  # 搜索所有年份
)
```

### Q4: 如何更新题库索引？

当题库有新文件时：

```bash
python scripts/indexer.py
```

这会重新扫描整个题库目录并更新索引。

## 🛠️ 开发计划

- [ ] Web UI 界面
- [ ] 支持更多文档格式（.doc, .pdf）
- [ ] AI 自动内容生成集成
- [ ] 云端题库同步

## 📝 贡献指南

欢迎提交 Issue 和 PR！

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

## 📄 许可证

本项目基于 MIT 许可证开源 - 详见 [LICENSE](LICENSE) 文件

## 👤 作者

**xielk** - [GitHub](https://github.com/xielk)

---

⭐ 如果这个项目对你有帮助，请给个 Star！
