#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown转Word转换器 - Pandoc方案

使用Pandoc + pypandoc实现markdown到word的转换
这是功能最强大、效果最好的转换方案

依赖:
- pandoc (系统依赖，需要单独安装)
- pypandoc (Python包)

作者: AI Assistant
日期: 2025
"""

import os
import subprocess
from pathlib import Path
from typing import Optional, List, Dict, Any


class MarkdownConverter:
    """
    使用Pandoc进行markdown到word的转换器
    
    特点:
    - 自动处理pandoc路径问题
    - 支持复杂的markdown语法
    - 完美的格式保持
    - 支持数学公式、表格、图片
    - 可自定义模板和样式
    """
    
    def __init__(self, pandoc_path: Optional[str] = None):
        """
        初始化转换器
        
        参数:
            pandoc_path: 可选的pandoc可执行文件路径
        """
        self.pandoc_path = pandoc_path
        self._setup_environment()
    
    def _setup_environment(self):
        """设置pandoc环境"""
        try:
            import pypandoc
            
            # 如果指定了pandoc路径，设置给pypandoc
            if self.pandoc_path and os.path.exists(self.pandoc_path):
                os.environ['PYPANDOC_PANDOC'] = self.pandoc_path
                print(f"✅ 使用指定的pandoc路径: {self.pandoc_path}")
            else:
                # 尝试查找系统中的pandoc
                found_path = self._find_pandoc()
                if found_path:
                    self.pandoc_path = found_path
                    os.environ['PYPANDOC_PANDOC'] = found_path
                    print(f"✅ 找到pandoc安装: {found_path}")
            
            # 注意：pypandoc.download_pandoc() 会下载pandoc二进制文件
            # 这里不自动下载，避免每次运行都生成文件
            # 如果需要内置pandoc，请手动调用 pypandoc.download_pandoc()
            print("ℹ️ 使用系统pandoc或手动下载的pypandoc内置版本")
            
            # 测试pypandoc是否工作
            test_result = pypandoc.convert_text("# 测试", 'html', format='md')
            if test_result:
                print("✅ pypandoc工作正常")
                return True
                
        except ImportError:
            raise ImportError("❌ pypandoc未安装，请运行: pip install pypandoc")
        except Exception as e:
            print(f"⚠️ 环境设置警告: {e}")
            return False
    
    def _find_pandoc(self):
        """查找系统中的pandoc安装路径"""
        possible_paths = [
            # 常见的Windows安装路径
            r"C:\Program Files\Pandoc\pandoc.exe",
            r"C:\Program Files (x86)\Pandoc\pandoc.exe",
            r"C:\Users\{}\AppData\Local\Pandoc\pandoc.exe".format(os.getenv('USERNAME', '')),
            # Chocolatey安装路径
            r"C:\ProgramData\chocolatey\bin\pandoc.exe",
            # Scoop安装路径
            r"C:\Users\{}\scoop\apps\pandoc\current\pandoc.exe".format(os.getenv('USERNAME', '')),
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        # 尝试通过命令行检查
        try:
            result = subprocess.run(['pandoc', '--version'], capture_output=True, text=True, check=True)
            if result.returncode == 0:
                return 'pandoc'  # 系统PATH中可用
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass
        
        return None
    
    def convert_file(self, 
                    input_file: str, 
                    output_file: str,
                    template: Optional[str] = None,
                    toc: bool = False,
                    number_sections: bool = False,
                    extra_args: Optional[List[str]] = None) -> bool:
        """
        转换markdown文件到docx格式
        
        参数:
            input_file: 输入的markdown文件路径
            output_file: 输出的docx文件路径
            template: 可选的Word模板文件路径
            toc: 是否生成目录
            number_sections: 是否给章节编号
            extra_args: 额外的pandoc参数
        
        返回:
            bool: 转换是否成功
        """
        try:
            import pypandoc
            
            # 检查输入文件是否存在
            if not os.path.exists(input_file):
                print(f"❌ 输入文件不存在: {input_file}")
                return False
            
            # 构建转换参数
            args = extra_args or []
            
            # 添加模板参数
            if template and os.path.exists(template):
                args.extend(['--reference-doc', template])
                print(f"📄 使用模板: {template}")
            
            # 添加目录
            if toc:
                args.append('--toc')
                print("📑 生成目录")
            
            # 添加章节编号
            if number_sections:
                args.append('--number-sections')
                print("🔢 章节编号")
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            print(f"🔄 开始转换: {os.path.basename(input_file)} -> {os.path.basename(output_file)}")
            
            # 执行转换
            pypandoc.convert_file(
                input_file, 
                'docx', 
                outputfile=output_file,
                extra_args=args
            )
            
            # 检查输出文件是否生成
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"✅ 转换成功! 输出文件: {output_file} ({file_size/1024:.1f}KB)")
                return True
            else:
                print("❌ 转换失败: 输出文件未生成")
                return False
                
        except Exception as e:
            print(f"❌ 转换过程中出错: {e}")
            print("\n💡 可能的解决方案:")
            print("1. 检查pypandoc是否正确安装: pip install pypandoc")
            print("2. 确保pandoc已安装: https://pandoc.org/installing.html")
            print("3. 检查文件路径和权限")
            return False
    
    def convert_text(self, 
                    markdown_text: str, 
                    output_file: str,
                    template: Optional[str] = None,
                    toc: bool = False,
                    number_sections: bool = False) -> bool:
        """
        转换markdown字符串到docx格式
        
        参数:
            markdown_text: markdown文本内容
            output_file: 输出的docx文件路径
            template: 可选的Word模板文件路径
            toc: 是否生成目录
            number_sections: 是否给章节编号
        
        返回:
            bool: 转换是否成功
        """
        try:
            import pypandoc
            
            # 构建转换参数
            args = []
            
            if template and os.path.exists(template):
                args.extend(['--reference-doc', template])
            
            if toc:
                args.append('--toc')
            
            if number_sections:
                args.append('--number-sections')
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            print(f"🔄 转换文本到: {os.path.basename(output_file)}")
            
            # 执行转换
            pypandoc.convert_text(
                markdown_text, 
                'docx',
                format='md',
                outputfile=output_file,
                extra_args=args
            )
            
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"✅ 转换成功! 输出文件: {output_file} ({file_size/1024:.1f}KB)")
                return True
            else:
                print("❌ 转换失败: 输出文件未生成")
                return False
                
        except Exception as e:
            print(f"❌ 转换过程中出错: {e}")
            return False
    
    def batch_convert(self, 
                     input_dir: str, 
                     output_dir: str,
                     template: Optional[str] = None,
                     pattern: str = "*.md",
                     toc: bool = False,
                     number_sections: bool = False) -> Dict[str, bool]:
        """
        批量转换目录中的markdown文件
        
        参数:
            input_dir: 输入目录路径
            output_dir: 输出目录路径
            template: 可选的Word模板文件路径
            pattern: 文件匹配模式，默认"*.md"
            toc: 是否生成目录
            number_sections: 是否给章节编号
        
        返回:
            Dict[str, bool]: 文件名到转换结果的映射
        """
        input_path = Path(input_dir)
        output_path = Path(output_dir)
        
        if not input_path.exists():
            print(f"❌ 输入目录不存在: {input_dir}")
            return {}
        
        # 创建输出目录
        output_path.mkdir(parents=True, exist_ok=True)
        
        # 查找所有markdown文件
        md_files = list(input_path.glob(pattern))
        
        if not md_files:
            print(f"❌ 在 {input_dir} 中未找到匹配 {pattern} 的文件")
            return {}
        
        print(f"📁 找到 {len(md_files)} 个文件待转换")
        
        results = {}
        success_count = 0
        
        for md_file in md_files:
            output_file = output_path / f"{md_file.stem}.docx"
            
            print(f"\n📄 处理: {md_file.name}")
            
            success = self.convert_file(
                str(md_file), 
                str(output_file),
                template=template,
                toc=toc,
                number_sections=number_sections
            )
            
            results[md_file.name] = success
            if success:
                success_count += 1
        
        print(f"\n📊 批量转换完成: {success_count}/{len(md_files)} 个文件转换成功")
        return results
    
    def create_sample_markdown(self, filename: str = "sample.md") -> str:
        """
        创建一个示例markdown文件用于测试
        
        参数:
            filename: 输出文件名
        
        返回:
            str: 创建的文件路径
        """
        sample_content = """
# Markdown转Word示例

这是一个**markdown转word**的示例文档。

## 功能特点

- ✅ 支持**粗体**和*斜体*文本
- ✅ 支持`行内代码`和代码块
- ✅ 支持表格和列表

## 代码示例

```python
def hello_world():
    print("Hello, World!")
```

## 表格示例

| 功能 | 状态 |
|------|------|
| 转换 | ✅ |
| 格式 | ✅ |

> 这是一个引用块示例。

更多详细示例请参考README.md文件。
"""
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(sample_content)
        
        print(f"📝 创建示例文件: {filename}")
        return filename


def main():
    """主函数 - 演示转换器的使用"""
    print("🚀 Markdown转Word转换器 (Pandoc方案)")
    print("=" * 50)
    
    try:
        # 创建转换器实例
        converter = MarkdownConverter()
        
        # 创建示例文件
        sample_file = converter.create_sample_markdown()
        
        # 基本转换
        print("\n📄 基本转换演示:")
        success = converter.convert_file(sample_file, "output_basic.docx")
        
        if success:
            print("\n📑 带目录的转换演示:")
            converter.convert_file(
                sample_file, 
                "output_with_toc.docx",
                toc=True,
                number_sections=True
            )
            
            print("\n📝 字符串转换演示:")
            test_markdown = """
# 字符串转换测试

这是通过**字符串**直接转换的示例。

## 功能特点

- 支持实时转换
- 无需临时文件
- 内存效率高

```python
print(\"Hello, World!\")
```
"""
            converter.convert_text(test_markdown, "output_string.docx")
            
            print("\n✅ 所有演示完成!")
            print("\n生成的文件:")
            for file in ["output_basic.docx", "output_with_toc.docx", "output_string.docx"]:
                if os.path.exists(file):
                    size = os.path.getsize(file)
                    print(f"  📄 {file} ({size/1024:.1f}KB)")
        
    except Exception as e:
        print(f"❌ 程序运行出错: {e}")
        print("\n💡 解决建议:")
        print("1. 确保已安装pandoc: https://pandoc.org/installing.html")
        print("2. 确保已安装pypandoc: pip install pypandoc")
        print("3. 检查文件路径和权限")


if __name__ == '__main__':
    main()