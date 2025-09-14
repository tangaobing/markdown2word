#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown转Word转换工具

简单易用的命令行工具，用于将markdown文件转换为Word文档

使用方法:
    python convert.py input.md output.docx
    python convert.py input.md                    # 自动生成output.docx
    python convert.py                             # 交互式模式

作者: AI Assistant
日期: 2025
"""

import os
import sys
import argparse
from pathlib import Path
from markdown_converter import MarkdownConverter


def find_pandoc_paths():
    """查找可能的pandoc安装路径"""
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
    
    found_paths = []
    for path in possible_paths:
        if os.path.exists(path):
            found_paths.append(path)
    
    return found_paths


def get_converter():
    """设置转换器"""
    # 查找pandoc路径
    pandoc_paths = find_pandoc_paths()
    selected_path = pandoc_paths[0] if pandoc_paths else None
    
    # 创建转换器
    return MarkdownConverter(pandoc_path=selected_path)


def convert_single_file(input_file, output_file=None, add_toc=False, number_sections=False):
    """转换单个文件"""
    # 如果没有指定输出文件，自动生成
    if not output_file:
        input_path = Path(input_file)
        output_file = input_path.with_suffix('.docx')
    
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        print(f"❌ 输入文件不存在: {input_file}")
        return False
    
    # 创建转换器并执行转换
    converter = get_converter()
    success = converter.convert_file(
        str(input_file), 
        str(output_file),
        add_toc=add_toc,
        number_sections=number_sections
    )
    
    if success:
        print(f"\n🎉 转换完成!")
        print(f"📄 输出文件: {output_file}")
        
        # 显示文件信息
        if os.path.exists(output_file):
            size = os.path.getsize(output_file)
            print(f"📊 文件大小: {size/1024:.1f}KB")
    
    return success


def batch_convert(input_dir, output_dir=None, pattern="*.md"):
    """批量转换目录中的文件"""
    input_path = Path(input_dir)
    
    if not input_path.exists():
        print(f"❌ 输入目录不存在: {input_dir}")
        return False
    
    # 设置输出目录
    if not output_dir:
        output_dir = input_path / "converted"
    
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # 查找markdown文件
    md_files = list(input_path.glob(pattern))
    
    if not md_files:
        print(f"❌ 在 {input_dir} 中未找到匹配 {pattern} 的文件")
        return False
    
    print(f"📁 找到 {len(md_files)} 个文件待转换")
    
    # 创建转换器
    converter = setup_converter()
    
    success_count = 0
    for md_file in md_files:
        output_file = output_path / f"{md_file.stem}.docx"
        
        print(f"\n📄 转换: {md_file.name}")
        
        if converter.convert_file(str(md_file), str(output_file)):
            success_count += 1
    
    print(f"\n📊 批量转换完成: {success_count}/{len(md_files)} 个文件转换成功")
    print(f"📁 输出目录: {output_dir}")
    
    return success_count > 0


def interactive_mode():
    """交互式模式"""
    print("🚀 Markdown转Word转换工具 - 交互模式")
    print("=" * 50)
    
    while True:
        print("\n请选择操作:")
        print("1. 转换单个文件")
        print("2. 批量转换目录")
        print("3. 退出")
        
        choice = input("\n请输入选择 (1-3): ").strip()
        
        if choice == '1':
            # 单文件转换
            input_file = input("请输入markdown文件路径: ").strip().strip('"')
            if not input_file:
                continue
            
            output_file = input("请输入输出文件路径 (回车使用默认): ").strip().strip('"')
            if not output_file:
                output_file = None
            
            # 询问是否添加目录
            add_toc = input("是否添加目录? (y/N): ").strip().lower() == 'y'
            number_sections = input("是否给章节编号? (y/N): ").strip().lower() == 'y'
            
            convert_single_file(input_file, output_file, add_toc, number_sections)
            
        elif choice == '2':
            # 批量转换
            input_dir = input("请输入包含markdown文件的目录路径: ").strip().strip('"')
            if not input_dir:
                continue
            
            output_dir = input("请输入输出目录路径 (回车使用默认): ").strip().strip('"')
            if not output_dir:
                output_dir = None
            
            pattern = input("请输入文件匹配模式 (回车使用*.md): ").strip()
            if not pattern:
                pattern = "*.md"
            
            batch_convert(input_dir, output_dir, pattern)
            
        elif choice == '3':
            print("👋 再见!")
            break
        else:
            print("❌ 无效选择，请重新输入")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='Markdown转Word转换工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python convert.py document.md                    # 转换为document.docx
  python convert.py document.md output.docx        # 指定输出文件
  python convert.py document.md --toc              # 添加目录
  python convert.py --batch ./docs ./output       # 批量转换
  python convert.py                               # 交互模式
        """
    )
    
    parser.add_argument('input', nargs='?', help='输入markdown文件或目录路径')
    parser.add_argument('output', nargs='?', help='输出docx文件或目录路径')
    parser.add_argument('--toc', action='store_true', help='添加目录')
    parser.add_argument('--number-sections', action='store_true', help='给章节编号')
    parser.add_argument('--batch', action='store_true', help='批量转换模式')
    parser.add_argument('--pattern', default='*.md', help='批量转换时的文件匹配模式')
    
    args = parser.parse_args()
    
    # 如果没有提供参数，进入交互模式
    if not args.input:
        interactive_mode()
        return
    
    # 批量转换模式
    if args.batch:
        batch_convert(args.input, args.output, args.pattern)
    else:
        # 单文件转换模式
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
        print("\n\n👋 用户取消操作")
    except Exception as e:
        print(f"\n❌ 程序出错: {e}")
        print("\n💡 如果问题持续存在，请检查:")
        print("1. pypandoc是否正确安装")
        print("2. 输入文件路径是否正确")
        print("3. 是否有文件写入权限")