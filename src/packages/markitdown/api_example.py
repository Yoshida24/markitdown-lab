#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
MarkItDown Python API使用例

このスクリプトは、markitdownライブラリをPythonプログラム内で使用して
PowerPointファイルをMarkdownに変換する方法を示します。
"""

from markitdown import MarkItDown
import os

def example_basic_usage(pptx_file_path):
    """基本的な使用方法"""
    print("\n=== 基本的な使用方法 ===")
    
    # MarkItDownインスタンスを作成
    md = MarkItDown()
    
    # ファイルを変換
    result = md.convert(pptx_file_path)
    
    # 結果を表示
    print(f"変換されたテキスト（最初の500文字）:\n{result.text_content[:500]}...\n")
    
    return result.text_content

def example_save_to_file(pptx_file_path, output_path):
    """結果をファイルに保存する例"""
    print("\n=== ファイルへの保存例 ===")
    
    # MarkItDownインスタンスを作成
    md = MarkItDown()
    
    # ファイルを変換
    result = md.convert(pptx_file_path)
    
    # 結果をファイルに保存
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(result.text_content)
    
    print(f"Markdownが '{output_path}' に保存されました。")
    
    return output_path

def example_with_plugins(pptx_file_path):
    """プラグイン関連の機能（非サポート）"""
    print("\n=== プラグイン機能は現在サポートされていません ===")
    print("注: 現在のバージョンではプラグイン機能は使用できません。")
    
    # 基本的な変換を実行
    md = MarkItDown()
    result = md.convert(pptx_file_path)
    
    # 結果を表示
    print(f"変換されたテキスト（最初の500文字）:\n{result.text_content[:500]}...\n")
    
    return result.text_content

def main():
    # サンプルのPowerPointファイルパス
    # 実際のファイルパスに変更してください
    pptx_file_path = "path/to/your/presentation.pptx"
    
    # ファイルが存在しない場合は終了
    if not os.path.exists(pptx_file_path):
        print(f"エラー: ファイル '{pptx_file_path}' が見つかりません。")
        print("使用例を実行するには、変数 pptx_file_path にあなたのPowerPointファイルパスを設定してください。")
        return
    
    # 基本的な使用方法
    example_basic_usage(pptx_file_path)
    
    # ファイルに保存する例
    output_path = os.path.splitext(pptx_file_path)[0] + "_converted.md"
    example_save_to_file(pptx_file_path, output_path)
    
    # プラグイン関連機能（現在は非サポート）
    try:
        example_with_plugins(pptx_file_path)
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
    
    print("\n全ての例が完了しました。")

if __name__ == "__main__":
    main() 