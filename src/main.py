import os
import sys
from packages.cowsay_demo.greeting import greet_to


def main():
    """
    メインエントリーポイント
    
    このアプリには以下の機能があります：
    - cowsay: cowsayデモ
    - markitdown: PowerPointからMarkdownへの変換ツール
    
    使用方法：
    - Tkinter GUI: make app
    - Streamlit GUI: make streamlit
    - CLI: make cli ARGS="file.pptx -o output.md"
    """
    # Cowsayデモを実行
    print("\nCowsayデモ:\n")
    greet_to(your_name="User")
    
    # Markitdownの情報を表示
    print("\nMarkitdownアプリケーション情報:\n")
    print("PowerPointファイルをMarkdownに変換するツールです。")
    print("以下のコマンドで利用できます：")
    print("- Tkinter GUI: make app")
    print("- Streamlit GUI: make streamlit")
    print("- CLI: make cli ARGS=\"file.pptx -o output.md\"")


if __name__ == "__main__":
    main()
