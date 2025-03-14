![badge](https://img.shields.io/badge/Python-white?logo=python) 
![badge](https://img.shields.io/badge/preset-red) 
[![badge](https://img.shields.io/badge/Package_manager-uv-8A2BE2)](https://docs.astral.sh/uv/) ![badge](https://img.shields.io/badge/Linter-Ruff-yellow)

# MarkItDown Lab

PowerPointファイルをMarkdownに変換するツール。ryeパッケージマネージャを使用して管理されています。

## 機能

- PowerPointプレゼンテーションをMarkdownドキュメントに変換
- Streamlitを使用したウェブインターフェース
- コマンドラインインターフェース

## 使用方法

- Python: 3.12
- uv: 0.6.2
- OS and Device: M1 Macbook Air Sequoia 15.3.1

## 前提条件

`uv`をインストールします：

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
echo 'eval "$(uv generate-shell-completion zsh)"' >> ~/.zshrc
```

## 始め方
まず、VSCodeの推奨拡張機能をインストールしてください。これには、リンター、フォーマッターなどが含まれています。推奨設定は`.vscode/extensions.json`に記載されています。

依存関係をインストールします：

```bash
uv sync
```

[Optional] 環境変数を使う場合  
`.env`ファイルの環境変数を使用するには、以下のスクリプトを実行して`.env`を作成します：

```bash
if [ ! -f .env ]; then
    cp .env.tmpl .env
    echo 'Info: .env file has successfully created. Please rewrite .env file'
else
    echo 'Info: Skip to create .env file. Because it is already exists.'
fi
```

これでスクリプトを実行できます：

### Streamlit GUIアプリの起動

```bash
# .envから環境変数をシェルにロード
# set -a && source ./.env && set +a
make streamlit
```

### Tkinter GUIアプリの起動

```bash
set -a && source ./.env && set +a
make app
```

### コマンドラインインターフェースの使用

```bash
set -a && source ./.env && set +a
make cli ARGS="path/to/file.pptx -o output.md"
```

## Jupyter Notebookの使用

VSCodeでは、`src/notebook.ipyenv`を`.venv`と自動的に使用できます。

Jupyterサーバーを起動するには：

```bash
uv run --with jupyter jupyter lab
# 自動的に http://localhost:8888/lab が開きます
```

## チートシート
依存関係の追加：

```bash
uv add requests
```

開発依存関係の追加：

```bash
uv add --dev ruff
```

Pythonバージョンの固定：

```bash
uv python pin 3.12
```

`uv`の更新：

```bash
uv self update
```