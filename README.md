# Markdown转Word转换器

将Markdown文件转换为Word文档(.docx)格式的Python工具。

## 安装

### 系统要求

- Python 3.7+
- Pandoc

### 安装步骤

1. 安装Python依赖:
   ```bash
   pip install pypandoc
   ```

2. 安装Pandoc:
   - Windows: 从 https://pandoc.org/installing.html 下载安装
   - macOS: `brew install pandoc`
   - Linux: `sudo apt-get install pandoc`

## 使用

### 命令行模式

单文件转换:
```bash
python convert.py input.md
python convert.py input.md -o output.docx
```

批量转换:
```bash
python convert.py docs/ -o output/
```

### 交互式模式

```bash
python convert.py
```

## 📁 项目文件说明

| 文件 | 说明 |
|------|------|
| `convert.py` | **命令行工具** - 支持单文件和批量转换 |
| `markdown_converter.py` | **核心模块** - 统一的转换器类，包含所有功能 |
| `install_dependencies.py` | **安装脚本** - 自动安装所需依赖 |
| `requirements.txt` | **依赖列表** - Python包依赖 |

## 🛠️ 高级使用

### 命令行选项

```bash
# 自动生成输出文件名
python convert.py document.md  # 生成 document.docx

# 添加目录
python convert.py document.md --toc

# 给章节编号
python convert.py document.md --number-sections

# 批量转换目录
python convert.py --batch ./markdown_files ./word_files

# 批量转换指定模式的文件
python convert.py --batch ./docs ./output --pattern "*.markdown"
```

### Python API

```python
from markdown_converter import MarkdownConverter

converter = MarkdownConverter()
converter.convert_file("input.md", "output.docx")
```

## ✨ 功能特点

### ✅ 支持的Markdown语法

- **文本格式**: 粗体、斜体、删除线、行内代码
- **标题**: 1-6级标题，自动生成目录
- **列表**: 有序列表、无序列表、嵌套列表
- **表格**: 支持复杂表格格式
- **代码块**: 支持语法高亮
- **链接**: 内部和外部链接
- **引用块**: 多级引用
- **图片**: 自动处理图片路径

## 📝 完整示例

以下是一个完整的Markdown示例，展示了各种语法元素的使用：

```markdown
## 1. 介绍

这是一个**markdown转word**的示例文档，展示了各种markdown元素的转换效果。

### 1.1 项目特点

- ✅ 支持**粗体**和*斜体*文本
- ✅ 支持`行内代码`和代码块
- ✅ 支持表格和列表
- ✅ 支持图片和链接
- ✅ 支持数学公式 (如果启用)

## 2. 文本格式示例

### 2.1 基本格式

这是普通文本。这是**粗体文本**。这是*斜体文本*。这是***粗斜体文本***。

这是`行内代码`示例。这是~~删除线~~文本。

### 2.2 引用块

> 这是一个引用块。
> 
> 引用块可以包含多行内容，
> 并且可以嵌套其他markdown元素。

## 3. 列表示例

### 3.1 无序列表

- 第一项
- 第二项
  - 子项目 2.1
  - 子项目 2.2
    - 子子项目 2.2.1
- 第三项

### 3.2 有序列表

1. 第一步：安装依赖
2. 第二步：配置环境
3. 第三步：运行程序
   1. 检查输入文件
   2. 执行转换
   3. 验证输出

## 4. 表格示例

| 功能 | 支持程度 | 说明 |
|------|----------|------|
| 标题 | ✅ 完全支持 | 支持1-6级标题 |
| 列表 | ✅ 完全支持 | 支持有序和无序列表 |
| 表格 | ✅ 完全支持 | 支持复杂表格格式 |
| 图片 | ✅ 完全支持 | 自动处理图片路径 |
| 链接 | ✅ 完全支持 | 支持内部和外部链接 |
| 代码 | ✅ 完全支持 | 支持语法高亮 |

## 5. 代码示例

### 5.1 Python代码

```python
def markdown_to_word(input_file, output_file):
    """转换markdown到word"""
    converter = MarkdownConverter()
    return converter.convert_file(input_file, output_file)

# 使用示例
result = markdown_to_word("input.md", "output.docx")
print(f"转换结果: {result}")
```

### 5.2 JavaScript代码

```javascript
function convertMarkdown(inputFile, outputFile) {
    // 这里是JavaScript示例
    console.log(`Converting ${inputFile} to ${outputFile}`);
    return true;
}
```

## 6. 链接示例

- [Pandoc官网](https://pandoc.org/)
- [Python官网](https://www.python.org/)
- [Markdown语法指南](https://www.markdownguide.org/)

## 7. 总结

这个示例文档展示了markdown的主要语法元素。通过Pandoc转换器，这些元素都能很好地转换为Word文档格式。

---

**注意**: 某些高级功能可能需要特定的pandoc版本或额外配置。
```
### 🔧 高级功能

- **自动路径检测**: 自动查找pandoc安装路径
- **批量转换**: 支持目录批量转换
- **目录生成**: 自动生成Word文档目录
- **章节编号**: 自动给章节添加编号
- **错误处理**: 完善的错误提示和解决建议
- **模板支持**: 支持自定义Word模板

## 🎯 使用示例

### 示例1: 基本转换

```bash
# 创建测试文件并转换
python markdown_converter.py
```

### 示例2: 带目录的文档

```bash
# 转换并添加目录和章节编号
python convert.py document.md final.docx --toc --number-sections
```

### 示例3: 批量转换

```bash
# 转换整个目录的markdown文件
python convert.py --batch ./markdown_docs ./word_docs
```

## 许可证

本项目采用MIT许可证。