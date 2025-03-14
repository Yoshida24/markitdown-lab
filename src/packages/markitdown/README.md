# MarkItDown PowerPoint文字起こしツール

Microsoft の [Microsoft の MarkItDown ライブラリ](https://github.com/microsoft/markitdown) ライブラリを使用して、ローカルの PowerPoint (PPTX) ファイルを Markdown 形式に変換するツールです。

# 機能

- PowerPoint (PPTX) ファイルを Markdown に変換
- コマンドライン操作とStreamlit GUIの両方をサポート
- 出力ファイルパスの指定オプション
- 変換結果のプレビューとダウンロード（StreamlitのGUI版）

# セットアップ方法

### 前提条件

- Python 3.8以上
- pip (Pythonパッケージマネージャー)

### 仮想環境のセットアップ

このツールは Python の仮想環境を使用してグローバル環境を汚さないようにしています。

1. リポジトリをクローンまたはダウンロードします

```bash
git clone https://github.com/yourusername/markitdown-tool.git
cd markitdown-tool
```

2. 仮想環境を作成してアクティベートします

```bash
# 仮想環境を作成
python3 -m venv .venv

# 仮想環境をアクティベート
# macOS/Linux:
source .venv/bin/activate

# Windows:
# .venv\Scripts\activate
```

3. 必要なパッケージをインストールします

```bash
pip install -r requirements.txt
```

# 使用方法

## Streamlit GUI版

1. 仮想環境がアクティベートされていることを確認します

```bash
# macOS/Linux:
source .venv/bin/activate
# Windows:
# .venv\Scripts\activate
```

2. Streamlitアプリを起動します

```bash
streamlit run streamlit_app.py
```

3. ブラウザが自動的に開き、アプリのインターフェースが表示されます
   - PowerPointファイルをアップロードします
   - 「Markdownに変換」ボタンをクリックして変換します
   - 変換結果を確認し、必要に応じてダウンロードします


## コマンドライン版

1. 仮想環境がアクティベートされていることを確認します

```bash
# macOS/Linux:
source .venv/bin/activate
# Windows:
.venv\Scripts\activate
```

2. コマンドラインからスクリプトを実行します

```bash
# 基本的な使い方
python cli.py path/to/your/presentation.pptx

# 出力ファイルを指定する場合
python cli.py path/to/your/presentation.pptx -o path/to/output/markdown.md
```

### コマンドラインオプション（cli.py）

- 第1引数（必須）: 変換する PowerPoint ファイルのパス
- `-o, --output`: 出力する Markdown ファイルのパス（省略時は入力ファイルと同じディレクトリに保存されます）

## 注意事項

- 大きなファイルの変換には時間がかかる場合があります
- 一部の複雑なフォーマットやレイアウトは正確に変換されない場合があります
- 画像や特殊な要素を含むスライドは、テキスト部分のみ抽出される場合があります

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。
