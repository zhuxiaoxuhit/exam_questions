# 考题生成系统

基于AI的智能考题生成工具，支持从教材文件自动生成高质量考题。

## 快速开始

### 1. 安装依赖
```bash
pip install openpyxl openai python-dotenv xlwt python-docx
```

### 2. 配置API密钥
创建 `.env` 文件：
```
DASHSCOPE_API_KEY=your_api_key_here
```

### 3. 运行脚本
```bash
# 自动处理resources目录中的所有教材文件（推荐）
python generate_questions.py

# 处理指定文件
python generate_questions.py 三级3001-3010.xlsx
```

## 主要特性

- ✅ **智能生成**: 使用通义千问AI生成高质量考题
- ✅ **增量处理**: 自动检测已生成的题目，只处理新增内容
- ✅ **双格式输出**: 同时生成XLS和Word两种格式
- ✅ **质量保证**: 内置题目评估和优化机制

## 目录结构

```
exam_questions/
├── generate_questions.py    # 主程序：AI考题生成
├── docx_to_xls.py          # 工具：Word转XLS格式
├── resources/              # 教材文件目录
└── questions/              # 生成的考题目录（自动创建）
```

## 格式转换工具

```bash
# 将Word格式题目转换为XLS
python docx_to_xls.py A-B-A-001.docx
python docx_to_xls.py questions/  # 批量转换
```
