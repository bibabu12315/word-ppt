# Word 转 PPT 自动化工具

本项目旨在实现从 Word 文档 (.docx) 到 PowerPoint 演示文稿 (.pptx) 的自动化转换。项目采用中间格式 (Markdown) 的方式，允许用户在生成 PPT 之前对内容进行人工校验和调整。

## 项目结构与文件说明

```text
word转ppt/
├── generate_md.py          # [入口] Word 转 Markdown 脚本
├── generate_ppt.py         # [入口] Markdown 转 PPT 脚本
├── main.py                 # [入口] 命令行主程序，串联整个流程
├── app.py                  # [入口] Streamlit Web 应用程序
├── requirements.txt        # 项目依赖库列表
├── README.md               # 项目说明文档
├── input/                  # 输入文件目录 (Word 文档)
├── output/                 # 输出文件目录 (Markdown, PPT)
├── parser/                 # 文档解析模块
│   ├── word_parser.py      # Word 文档解析逻辑
│   ├── markdown_parser.py  # Markdown 解析逻辑
│   ├── json_to_md.py       # JSON 转 Markdown 工具
│   └── data_structs.py     # 数据结构定义 (SlideData 等)
├── ppt/                    # PPT 生成模块
│   └── generator.py        # PPT 生成核心逻辑
├── utils/                  # 通用工具模块
│   ├── ppt_utils.py        # PPT 操作底层工具 (复制幻灯片/形状)
│   ├── create_template.py  # 创建默认 PPT 模板
│   └── create_dummy_docx.py# 生成测试用 Word 文档
└── ai/                     # AI 辅助模块
    └── llm_client.py       # LLM API 客户端 (如通义千问)
```

## 模块详细说明

### 1. 根目录 (入口脚本)
- **`generate_md.py`**: 负责将 Word 文档解析并转换为 Markdown 格式。这是转换流程的第一步。
- **`generate_ppt.py`**: 负责读取 Markdown 文件，并根据模板生成最终的 PPT 文件。这是转换流程的第二步。
- **`main.py`**: 命令行工具，可以一次性执行完整的转换流程，也可以单独执行某一步骤。
- **`app.py`**: 提供了一个基于 Streamlit 的 Web 界面，方便非技术人员上传文档、预览内容并下载 PPT。

### 2. Parser (解析模块)
- **`parser/word_parser.py`**: 使用 `python-docx` 读取 Word 文档，提取标题、正文、列表等信息，生成中间 JSON 数据。
- **`parser/markdown_parser.py`**: 解析特定格式的 Markdown 文件，将其转换为程序内部使用的 `PresentationData` 对象。
- **`parser/json_to_md.py`**: 辅助工具，将 `word_parser` 生成的 JSON 数据格式化为易读的 Markdown 文本。
- **`parser/data_structs.py`**: 定义了项目中使用的数据类（Dataclasses），如 `SlideData`（幻灯片数据）、`ContentBlock`（内容块）等，确保各模块间数据交互的规范性。

### 3. PPT (生成模块)
- **`ppt/generator.py`**: 核心生成器。它加载 PPT 模板，根据 `PresentationData` 中的数据，克隆模板页，填充文本，并处理样式（字体、颜色、大小）的迁移。

### 4. Utils (工具模块)
- **`utils/ppt_utils.py`**: 包含底层的 PPT 操作函数，例如 `duplicate_slide`（复制幻灯片）、`duplicate_shape`（复制形状）、样式复制等。
- **`utils/create_template.py`**: 用于生成一个基础的 PPT 模板文件，防止因缺少模板导致程序运行失败。
- **`utils/create_dummy_docx.py`**: 开发测试用，生成一个包含标准层级结构的 Word 文档。

### 5. AI (AI 模块)
- **`ai/llm_client.py`**: 封装了与大语言模型（如 Qwen）交互的客户端，用于未来扩展 AI 辅助生成、摘要提取等功能。

## 使用流程

1.  **准备环境**: 安装依赖 `pip install -r requirements.txt`
2.  **准备模板**: 确保 `input/template.pptx` 存在（或使用默认模板）。
3.  **转换 Word**: 运行 `python generate_md.py` 将 Word 转换为 Markdown。
4.  **编辑 Markdown**: (可选) 在 `output/` 目录下修改生成的 Markdown 文件。
5.  **生成 PPT**: 运行 `python generate_ppt.py` 生成最终 PPT。
