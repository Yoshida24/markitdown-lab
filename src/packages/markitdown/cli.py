#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import argparse
from markitdown import MarkItDown

def convert_pptx_to_markdown(input_path, output_path=None):
    """
    PowerPointファイルをMarkdownに変換する
    
    Args:
        input_path (str): 入力PowerPointファイルのパス
        output_path (str, optional): 出力Markdownファイルのパス
    
    Returns:
        str: 変換されたMarkdownテキスト
    """
    if not os.path.exists(input_path):
        print(f"エラー: ファイル '{input_path}' が見つかりません。")
        return None
    
    print(f"'{input_path}' を変換中...")
    
    try:
        # MarkItDownを使用してPowerPointファイルをMarkdownに変換
        md = MarkItDown()
        result = md.convert(input_path)
        
        # 出力パスが指定されていない場合、入力ファイルと同じディレクトリに保存
        if output_path is None:
            file_name = os.path.splitext(os.path.basename(input_path))[0]
            output_dir = os.path.dirname(input_path)
            output_path = os.path.join(output_dir, f"{file_name}.md")
        
        # 結果をファイルに保存
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(result.text_content)
        
        print(f"変換完了: '{output_path}' に保存されました。")
        return result.text_content
    
    except Exception as e:
        print(f"エラー: 変換中にエラーが発生しました: {str(e)}")
        return None

def main():
    # コマンドライン引数の解析
    parser = argparse.ArgumentParser(description="PowerPointファイルをMarkdownに変換するツール")
    parser.add_argument("input", help="入力PowerPointファイルのパス")
    parser.add_argument("-o", "--output", help="出力Markdownファイルのパス（省略時は入力ファイルと同じディレクトリに保存）")
    
    args = parser.parse_args()
    
    # PowerPointファイルをMarkdownに変換
    convert_pptx_to_markdown(args.input, args.output)

if __name__ == "__main__":
    main() 