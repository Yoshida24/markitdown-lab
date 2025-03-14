.PHONY: run streamlit cli app fmt test

run:
	uv run python src/main.py

# markitdownのStreamlit GUIアプリの起動
streamlit:
	uv run streamlit run src/packages/markitdown/streamlit_app.py

# markitdownのTkinter GUIアプリの起動
app:
	uv run python src/packages/markitdown/app.py

# markitdownのコマンドラインアプリの実行
cli:
	uv run python src/packages/markitdown/cli.py $(ARGS)

.PHONY: fmt
fmt:
	@uvx ruff format
	@uvx ruff check --fix --extend-select I

.PHONY: test
test:
	@echo "Define your test!"
