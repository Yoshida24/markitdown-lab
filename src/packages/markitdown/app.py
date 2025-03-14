#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from markitdown import MarkItDown

class MarkItDownApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MarkItDown PowerPoint文字起こしアプリ")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        self.setup_ui()
    
    def setup_ui(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # タイトルラベル
        title_label = ttk.Label(
            main_frame, 
            text="PowerPointファイルをMarkdownに変換", 
            font=("Helvetica", 16)
        )
        title_label.pack(pady=(0, 20))
        
        # ファイル選択フレーム
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=10)
        
        self.file_path_var = tk.StringVar()
        
        file_label = ttk.Label(file_frame, text="PowerPointファイル:")
        file_label.pack(side=tk.LEFT, padx=(0, 5))
        
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_button = ttk.Button(file_frame, text="参照...", command=self.browse_file)
        browse_button.pack(side=tk.LEFT)
        
        # 出力先フレーム
        output_frame = ttk.Frame(main_frame)
        output_frame.pack(fill=tk.X, pady=10)
        
        self.output_path_var = tk.StringVar()
        
        output_label = ttk.Label(output_frame, text="出力先フォルダ:")
        output_label.pack(side=tk.LEFT, padx=(0, 5))
        
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path_var, width=50)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        output_button = ttk.Button(output_frame, text="参照...", command=self.browse_output_folder)
        output_button.pack(side=tk.LEFT)
        
        # 進捗状況
        self.progress_var = tk.DoubleVar()
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=10)
        
        progress_label = ttk.Label(progress_frame, text="進捗状況:")
        progress_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100, 
            length=400
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 変換ボタン
        convert_button = ttk.Button(
            main_frame, 
            text="変換開始", 
            command=self.convert_file,
            style="Accent.TButton"
        )
        convert_button.pack(pady=20)
        
        # ステータスラベル
        self.status_var = tk.StringVar()
        self.status_var.set("PowerPointファイルを選択してください。")
        
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.pack(pady=5)
    
    def browse_file(self):
        filetypes = [
            ("PowerPointファイル", "*.pptx"),
            ("すべてのファイル", "*.*")
        ]
        
        file_path = filedialog.askopenfilename(
            title="PowerPointファイルを選択",
            filetypes=filetypes
        )
        
        if file_path:
            self.file_path_var.set(file_path)
            # デフォルトの出力先を同じディレクトリに設定
            output_dir = os.path.dirname(file_path)
            self.output_path_var.set(output_dir)
    
    def browse_output_folder(self):
        output_dir = filedialog.askdirectory(
            title="出力先フォルダを選択"
        )
        
        if output_dir:
            self.output_path_var.set(output_dir)
    
    def convert_file(self):
        input_path = self.file_path_var.get()
        output_dir = self.output_path_var.get()
        
        if not input_path:
            messagebox.showerror("エラー", "PowerPointファイルが選択されていません。")
            return
        
        if not output_dir:
            messagebox.showerror("エラー", "出力先フォルダが選択されていません。")
            return
        
        if not os.path.exists(input_path):
            messagebox.showerror("エラー", "選択されたファイルが存在しません。")
            return
        
        if not os.path.exists(output_dir):
            messagebox.showerror("エラー", "選択された出力先フォルダが存在しません。")
            return
        
        # ファイル名を取得（拡張子なし）
        file_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(output_dir, f"{file_name}.md")
        
        try:
            self.status_var.set("変換中...")
            self.progress_var.set(10)
            self.root.update()
            
            # MarkItDownを使用してPowerPointファイルをMarkdownに変換
            md = MarkItDown(enable_plugins=False)
            result = md.convert(input_path)
            
            self.progress_var.set(80)
            self.root.update()
            
            # 結果をファイルに保存
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(result.text_content)
            
            self.progress_var.set(100)
            self.status_var.set(f"変換完了: {output_path}")
            
            messagebox.showinfo("完了", f"Markdownファイルが保存されました:\n{output_path}")
            
        except Exception as e:
            self.status_var.set(f"エラー: {str(e)}")
            messagebox.showerror("エラー", f"変換中にエラーが発生しました:\n{str(e)}")

def main():
    root = tk.Tk()
    style = ttk.Style()
    style.configure("Accent.TButton", font=("Helvetica", 12))
    app = MarkItDownApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 