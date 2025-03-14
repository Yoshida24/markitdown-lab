import streamlit as st
import os
import tempfile
from markitdown import MarkItDown
import time
import base64

# ページ設定
st.set_page_config(
    page_title="MarkItDown PowerPoint文字起こしツール",
    page_icon="📝",
    layout="wide",
)

# タイトルとヘッダー
st.title("MarkItDown PowerPoint文字起こしツール")
st.markdown("### PowerPointファイルをMarkdownに変換")

# サイドバー
st.sidebar.header("MarkItDown について")
st.sidebar.markdown("""
このアプリは [Microsoft の MarkItDown ライブラリ](https://github.com/microsoft/markitdown) を使用して、
PowerPoint（PPTX）ファイルをMarkdown形式に変換します。

**特徴：**
- 簡単な操作でファイル変換
- スライドの内容を構造化されたMarkdownとして取得
- 変換結果をプレビューおよびダウンロード
""")

st.sidebar.markdown("---")
st.sidebar.markdown("**注意事項：**")
st.sidebar.markdown("""
- 大きなファイルの変換には時間がかかる場合があります
- 複雑なフォーマットやレイアウトは正確に変換されない場合があります
- 画像や特殊な要素を含むスライドは、テキスト部分のみ抽出される場合があります
""")

# メイン機能
def convert_pptx_to_markdown(uploaded_file, use_plugins=False):
    """
    アップロードされたPowerPointファイルをMarkdownに変換する
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
        # アップロードされたファイルを一時ファイルとして保存
        tmp_file.write(uploaded_file.getvalue())
        tmp_file_path = tmp_file.name
    
    try:
        # 進捗バーの表示
        progress_text = "変換中...しばらくお待ちください"
        progress_bar = st.progress(0)
        progress_bar.progress(10, text=progress_text)
        
        # MarkItDownを使用してPowerPointファイルをMarkdownに変換
        # 使用可能なパラメータに応じて初期化
        md = MarkItDown()
        
        # 進捗表示のアップデート
        progress_bar.progress(30, text=progress_text)
        time.sleep(0.5)  # 進捗バーの表示のための遅延
        
        # 変換実行
        result = md.convert(tmp_file_path)
        
        # 進捗表示のアップデート
        progress_bar.progress(90, text="変換完了！")
        time.sleep(0.5)  # 進捗バーの表示のための遅延
        
        # 変換完了
        progress_bar.progress(100, text="変換完了！")
        
        # 一時ファイルを削除
        os.unlink(tmp_file_path)
        
        return result.text_content
    
    except Exception as e:
        # エラーが発生した場合
        if os.path.exists(tmp_file_path):
            os.unlink(tmp_file_path)
        st.error(f"変換中にエラーが発生しました: {str(e)}")
        return None

def get_download_link(markdown_text, filename="converted.md"):
    """
    ダウンロードリンクを生成する
    """
    b64 = base64.b64encode(markdown_text.encode()).decode()
    return f'<a href="data:file/markdown;base64,{b64}" download="{filename}">クリックしてダウンロード</a>'

# ファイルアップロードエリア
uploaded_file = st.file_uploader("PowerPointファイルをアップロード", type=["pptx"])

# オプション設定
with st.expander("詳細設定"):
    use_plugins = st.checkbox("プラグインを使用する（現在のバージョンでは無効）", value=False, disabled=True)
    st.info("注: このオプションは現在のバージョンでは利用できません。基本機能のみ使用できます。")

# 変換ボタン
if uploaded_file is not None:
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.write(f"ファイル名: **{uploaded_file.name}**")
        st.write(f"ファイルサイズ: **{uploaded_file.size / 1024:.1f} KB**")
        
        convert_button = st.button("Markdownに変換", type="primary")
    
    if convert_button:
        markdown_text = convert_pptx_to_markdown(uploaded_file, use_plugins)
        
        if markdown_text:
            st.session_state["markdown_result"] = markdown_text
            st.session_state["filename"] = os.path.splitext(uploaded_file.name)[0]

# 変換結果の表示
if "markdown_result" in st.session_state and st.session_state["markdown_result"]:
    markdown_text = st.session_state["markdown_result"]
    filename = st.session_state["filename"]
    
    st.markdown("---")
    st.subheader("変換結果")
    
    # タブで表示方法を切り替え
    tab1, tab2 = st.tabs(["プレビュー", "マークダウンコード"])
    
    with tab1:
        st.markdown(markdown_text)
    
    with tab2:
        st.code(markdown_text, language="markdown")
    
    # ダウンロードボタン
    st.markdown("---")
    st.markdown(f"### マークダウンファイルをダウンロード")
    download_filename = f"{filename}.md"
    st.markdown(get_download_link(markdown_text, download_filename), unsafe_allow_html=True)
    st.caption(f"ファイル名: {download_filename}")

# フッター
st.markdown("---")
st.caption("Powered by [Microsoft MarkItDown](https://github.com/microsoft/markitdown)") 