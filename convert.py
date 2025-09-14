#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown转Word转换工具

简单的命令行工具，用于将markdown文件转换为Word文档
"""

import os
import argparse
from pathlib import Path
from markdown_converter import MarkdownConverter


def convert_single_file(input_file, output_file=None, add_toc=False, number_sections=False):
    """转换单个文件"""
    if not output_file:
        input_path = Path(input_file)
        output_file = input_path.with_suffix('.docx')
    
    if not os.path.exists(input_file):
        print(f"❌ 输入文件不存在: {input_file}")
        return False
    
    converter = MarkdownConverter()
    return converter.convert_file(
        str(input_file), 
        str(output_file),
        toc=add_toc,
        number_sections=number_sections
    )


def batch_convert(input_dir, output_dir=None, pattern="*.md"):
    """批量转换目录中的文件"""
    input_path = Path(input_dir)
    
    if not input_path.exists():
        print(f"❌ 输入目录不存在: {input_dir}")
        return False
    
    if not output_dir:
        output_dir = input_path / "converted"
    
    converter = MarkdownConverter()
    return converter.batch_convert(str(input_dir), str(output_dir), pattern=pattern)


def interactive_mode():
    """交互式模式"""
    print("🚀 Markdown转Word转换工具")
    
    input_file = input("请输入markdown文件路径: ").strip().strip('"')
    if not input_file:
        return
    
    output_file = input("请输入输出文件路径 (回车使用默认): ").strip().strip('"')
    if not output_file:
        output_file = None
    
    convert_single_file(input_file, output_file)


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='Markdown转Word转换工具')
    parser.add_argument('input', nargs='?', help='输入markdown文件或目录路径')
    parser.add_argument('output', nargs='?', help='输出docx文件或目录路径')
    parser.add_argument('--toc', action='store_true', help='添加目录')
    parser.add_argument('--number-sections', action='store_true', help='给章节编号')
    parser.add_argument('--batch', action='store_true', help='批量转换模式')
    parser.add_argument('--pattern', default='*.md', help='文件匹配模式')
    
    args = parser.parse_args()
    
    if not args.input:
        interactive_mode()
        return
    
    if args.batch:
        batch_convert(args.input, args.output, args.pattern)
    else:
        convert_single_file(
            args.input, 
            args.output,
            add_toc=args.toc,
            number_sections=args.number_sections
        )


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n操作已取消")
    except Exception as e:
        print(f"错误: {e}")