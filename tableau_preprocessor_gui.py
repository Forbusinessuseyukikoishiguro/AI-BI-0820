"""
Tableau Excel データ前処理ツール GUI版
Author: Claude
Description: ExcelデータをTableau分析用に効率的に前処理するためのGUIツール
"""

import pandas as pd
import numpy as np
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
from typing import List, Dict, Optional, Union, Any
import threading
import warnings
warnings.filterwarnings('ignore')


class TableauPreprocessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Tableau Excel前処理ツール")
        self.root.geometry("1400x800")
        self.root.configure(bg='#f0f0f0')
        
        # データ関連変数
        self.data = None
        self.original_data = None
        self.data_info = {}
        self.processing_log = []
        self.file_path = None
        
        # スタイル設定
        self.setup_styles()
        
        # GUI構築
        self.create_widgets()
        
        # 初期状態設定
        self.update_button_states()
    
    def setup_styles(self):
        """スタイルを設定"""
        self.style = ttk.Style()
        
        # テーマ設定
        available_themes = self.style.theme_names()
        if 'clam' in available_themes:
            self.style.theme_use('clam')
        
        # カスタムスタイル
        self.style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        self.style.configure('Heading.TLabel', font=('Arial', 12, 'bold'))
        self.style.configure('Action.TButton', font=('Arial', 10, 'bold'))
        self.style.configure('Success.TButton', 
                           background='#28a745', 
                           foreground='white',
                           font=('Arial', 10, 'bold'))
        self.style.configure('Danger.TButton', 
                           background='#dc3545', 
                           foreground='white',
                           font=('Arial', 10, 'bold'))
    
    def create_widgets(self):
        """GUIウィジェットを作成"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # グリッド設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # ヘッダー
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        title_label = ttk.Label(header_frame, text="📊 Tableau Excel前処理ツール", style='Title.TLabel')
        title_label.pack()
        
        subtitle_label = ttk.Label(header_frame, text="ExcelデータをTableau分析用に効率的にクリーニング・変換")
        subtitle_label.pack()
        
        # 左側パネル（コントロール）
        self.create_control_panel(main_frame)
        
        # 右側パネル（データ表示）
        self.create_data_panel(main_frame)
    
    def create_control_panel(self, parent):
        """左側のコントロールパネルを作成"""
        control_frame = ttk.Frame(parent, width=350)
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        control_frame.grid_propagate(False)
        
        # ファイル選択セクション
        file_section = ttk.LabelFrame(control_frame, text="📁 ファイル操作", padding="10")
        file_section.pack(fill=tk.X, pady=(0, 10))
        
        self.file_label = ttk.Label(file_section, text="ファイルが選択されていません", foreground='gray')
        self.file_label.pack(fill=tk.X, pady=(0, 5))
        
        file_button_frame = ttk.Frame(file_section)
        file_button_frame.pack(fill=tk.X)
        
        self.load_button = ttk.Button(file_button_frame, text="📂 ファイル選択", 
                                     command=self.load_file, style='Action.TButton')
        self.load_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.reload_button = ttk.Button(file_button_frame, text="🔄 再読み込み", 
                                       command=self.reload_file, state='disabled')
        self.reload_button.pack(side=tk.LEFT)
        
        # データ情報セクション
        info_section = ttk.LabelFrame(control_frame, text="📊 データ概要", padding="10")
        info_section.pack(fill=tk.X, pady=(0, 10))
        
        self.info_text = tk.Text(info_section, height=8, font=('Consolas', 9), 
                                state='disabled', bg='#f8f8f8')
        self.info_text.pack(fill=tk.BOTH)
        
        # クリーニングセクション
        cleaning_section = ttk.LabelFrame(control_frame, text="🧹 データクリーニング", padding="10")
        cleaning_section.pack(fill=tk.X, pady=(0, 10))
        
        self.empty_rows_button = ttk.Button(cleaning_section, text="空行削除", 
                                           command=self.remove_empty_rows, state='disabled')
        self.empty_rows_button.pack(fill=tk.X, pady=2)
        
        self.empty_cols_button = ttk.Button(cleaning_section, text="空列削除", 
                                           command=self.remove_empty_columns, state='disabled')
        self.empty_cols_button.pack(fill=tk.X, pady=2)
        
        self.duplicates_button = ttk.Button(cleaning_section, text="重複行削除", 
                                           command=self.remove_duplicates, state='disabled')
        self.duplicates_button.pack(fill=tk.X, pady=2)
        
        self.missing_button = ttk.Button(cleaning_section, text="欠損値処理", 
                                        command=self.fill_missing_dialog, state='disabled')
        self.missing_button.pack(fill=tk.X, pady=2)
        
        # 変換セクション
        transform_section = ttk.LabelFrame(control_frame, text="🔄 データ変換", padding="10")
        transform_section.pack(fill=tk.X, pady=(0, 10))
        
        self.text_clean_button = ttk.Button(transform_section, text="テキスト整形", 
                                           command=self.clean_text_data, state='disabled')
        self.text_clean_button.pack(fill=tk.X, pady=2)
        
        self.type_convert_button = ttk.Button(transform_section, text="型変換", 
                                             command=self.convert_data_types, state='disabled')
        self.type_convert_button.pack(fill=tk.X, pady=2)
        
        self.rename_button = ttk.Button(transform_section, text="列名変更", 
                                       command=self.rename_columns_dialog, state='disabled')
        self.rename_button.pack(fill=tk.X, pady=2)
        
        self.filter_button = ttk.Button(transform_section, text="データフィルタ", 
                                       command=self.filter_data_dialog, state='disabled')
        self.filter_button.pack(fill=tk.X, pady=2)
        
        # エクスポートセクション
        export_section = ttk.LabelFrame(control_frame, text="💾 エクスポート", padding="10")
        export_section.pack(fill=tk.X, pady=(0, 10))
        
        self.export_excel_button = ttk.Button(export_section, text="📄 Excel保存", 
                                             command=self.export_excel, 
                                             style='Success.TButton', state='disabled')
        self.export_excel_button.pack(fill=tk.X, pady=2)
        
        self.export_csv_button = ttk.Button(export_section, text="📄 CSV保存", 
                                           command=self.export_csv, 
                                           style='Success.TButton', state='disabled')
        self.export_csv_button.pack(fill=tk.X, pady=2)
        
        # その他セクション
        other_section = ttk.LabelFrame(control_frame, text="🔧 その他", padding="10")
        other_section.pack(fill=tk.X, pady=(0, 10))
        
        self.reset_button = ttk.Button(other_section, text="🔄 データリセット", 
                                      command=self.reset_data, 
                                      style='Danger.TButton', state='disabled')
        self.reset_button.pack(fill=tk.X, pady=2)
        
        self.log_button = ttk.Button(other_section, text="📋 処理ログ表示", 
                                    command=self.show_processing_log, state='disabled')
        self.log_button.pack(fill=tk.X, pady=2)
    
    def create_data_panel(self, parent):
        """右側のデータ表示パネルを作成"""
        data_frame = ttk.Frame(parent)
        data_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(1, weight=1)
        
        # データ表示ヘッダー
        data_header = ttk.Label(data_frame, text="データプレビュー", style='Heading.TLabel')
        data_header.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # データテーブル
        self.create_data_table(data_frame)
        
        # ステータスバー
        status_frame = ttk.Frame(data_frame)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        status_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(status_frame, text="ファイルを選択してください")
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        # プログレスバー
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, 
                                          maximum=100, length=200)
        self.progress_bar.grid(row=0, column=1, sticky=tk.E)
    
    def create_data_table(self, parent):
        """データテーブルを作成"""
        table_frame = ttk.Frame(parent)
        table_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # Treeviewウィジェット
        self.tree = ttk.Treeview(table_frame, show='tree headings', height=20)
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # スクロールバー
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.tree.configure(xscrollcommand=h_scrollbar.set)
    
    def load_file(self):
        """ファイルを読み込む"""
        file_path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.file_path = file_path
            self.load_data_async()
    
    def load_data_async(self):
        """非同期でデータを読み込む"""
        def load_worker():
            try:
                self.update_status("ファイルを読み込み中...")
                self.progress_var.set(20)
                
                # ファイル読み込み
                file_extension = os.path.splitext(self.file_path)[1].lower()
                
                if file_extension in ['.xlsx', '.xls']:
                    self.data = pd.read_excel(self.file_path)
                elif file_extension == '.csv':
                    encodings = ['utf-8', 'shift_jis', 'cp932', 'iso-2022-jp']
                    for encoding in encodings:
                        try:
                            self.data = pd.read_csv(self.file_path, encoding=encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                    else:
                        self.data = pd.read_csv(self.file_path, encoding='utf-8', errors='ignore')
                
                self.progress_var.set(70)
                
                # 元データのバックアップ
                self.original_data = self.data.copy()
                
                # GUI更新
                self.root.after(0, self.on_data_loaded)
                
            except Exception as e:
                self.root.after(0, lambda: self.show_error(f"ファイル読み込みエラー: {str(e)}"))
        
        # バックグラウンドで実行
        thread = threading.Thread(target=load_worker)
        thread.daemon = True
        thread.start()
    
    def on_data_loaded(self):
        """データ読み込み完了後の処理"""
        self.progress_var.set(100)
        
        # ファイル名表示
        filename = os.path.basename(self.file_path)
        self.file_label.config(text=f"📁 {filename}", foreground='black')
        
        # データ情報更新
        self.update_data_info()
        self.update_data_table()
        self.update_button_states(True)
        
        self.log_action(f"ファイル読み込み完了: {filename}")
        self.update_status(f"✅ データ読み込み完了 ({self.data.shape[0]}行 × {self.data.shape[1]}列)")
        
        # プログレスバーリセット
        self.root.after(2000, lambda: self.progress_var.set(0))
    
    def reload_file(self):
        """ファイルを再読み込み"""
        if self.file_path:
            self.load_data_async()
    
    def update_data_info(self):
        """データ情報を更新"""
        if self.data is None:
            return
        
        self.data_info = {
            'rows': len(self.data),
            'columns': len(self.data.columns),
            'empty_cells': self.data.isnull().sum().sum(),
            'duplicates': self.data.duplicated().sum(),
            'memory_usage': self.data.memory_usage(deep=True).sum()
        }
        
        # 情報テキスト更新
        info_text = f"""行数: {self.data_info['rows']:,}
列数: {self.data_info['columns']}
空セル数: {self.data_info['empty_cells']:,}
重複行数: {self.data_info['duplicates']:,}
メモリ使用量: {self.data_info['memory_usage']/1024/1024:.2f} MB

列情報:"""
        
        for i, (col, dtype) in enumerate(self.data.dtypes.items()):
            null_count = self.data[col].isnull().sum()
            null_rate = (null_count / len(self.data)) * 100
            info_text += f"\n{i+1:2d}. {col[:20]:<20} | {str(dtype):<10}"
            if null_count > 0:
                info_text += f" | 欠損:{null_count}({null_rate:.1f}%)"
        
        self.info_text.config(state='normal')
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(1.0, info_text)
        self.info_text.config(state='disabled')
    
    def update_data_table(self, max_rows=100):
        """データテーブルを更新"""
        if self.data is None:
            return
        
        # 既存のデータをクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 列設定
        columns = list(self.data.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"
        
        # ヘッダー設定
        for col in columns:
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=120, minwidth=80)
        
        # データ追加（最大100行まで）
        display_data = self.data.head(max_rows)
        for index, row in display_data.iterrows():
            values = []
            for col in columns:
                value = row[col]
                if pd.isna(value):
                    values.append("")
                elif isinstance(value, float):
                    values.append(f"{value:.2f}" if abs(value) < 1000000 else f"{value:.2e}")
                else:
                    values.append(str(value)[:50])  # 長すぎる値を切り詰め
            self.tree.insert("", "end", values=values)
        
        # 省略表示
        if len(self.data) > max_rows:
            self.tree.insert("", "end", values=[f"... 他 {len(self.data) - max_rows} 行"] + [""] * (len(columns) - 1))
    
    def update_button_states(self, enabled=False):
        """ボタンの有効/無効状態を更新"""
        state = 'normal' if enabled else 'disabled'
        buttons = [
            self.reload_button, self.empty_rows_button, self.empty_cols_button,
            self.duplicates_button, self.missing_button, self.text_clean_button,
            self.type_convert_button, self.rename_button, self.filter_button,
            self.export_excel_button, self.export_csv_button, self.reset_button,
            self.log_button
        ]
        
        for button in buttons:
            button.config(state=state)
    
    def update_status(self, message):
        """ステータス表示を更新"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def log_action(self, action):
        """処理ログを記録"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {action}"
        self.processing_log.append(log_entry)
    
    def show_error(self, message):
        """エラーダイアログを表示"""
        messagebox.showerror("エラー", message)
        self.progress_var.set(0)
    
    def show_success(self, message):
        """成功メッセージを表示"""
        messagebox.showinfo("完了", message)
    
    # データ処理メソッド
    def remove_empty_rows(self):
        """空行を削除"""
        if self.data is None:
            return
        
        try:
            initial_rows = len(self.data)
            self.data = self.data.dropna(how='all')
            removed_rows = initial_rows - len(self.data)
            
            self.update_data_info()
            self.update_data_table()
            self.log_action(f"空行削除: {removed_rows}行削除")
            self.update_status(f"✅ {removed_rows}行の空行を削除しました")
            
        except Exception as e:
            self.show_error(f"空行削除エラー: {str(e)}")
    
    def remove_empty_columns(self):
        """空列を削除"""
        if self.data is None:
            return
        
        try:
            initial_cols = len(self.data.columns)
            self.data = self.data.dropna(axis=1, how='all')
            removed_cols = initial_cols - len(self.data.columns)
            
            self.update_data_info()
            self.update_data_table()
            self.log_action(f"空列削除: {removed_cols}列削除")
            self.update_status(f"✅ {removed_cols}列の空列を削除しました")
            
        except Exception as e:
            self.show_error(f"空列削除エラー: {str(e)}")
    
    def remove_duplicates(self):
        """重複行を削除"""
        if self.data is None:
            return
        
        try:
            initial_rows = len(self.data)
            self.data = self.data.drop_duplicates()
            removed_rows = initial_rows - len(self.data)
            
            self.update_data_info()
            self.update_data_table()
            self.log_action(f"重複行削除: {removed_rows}行削除")
            self.update_status(f"✅ {removed_rows}行の重複を削除しました")
            
        except Exception as e:
            self.show_error(f"重複行削除エラー: {str(e)}")
    
    def clean_text_data(self):
        """テキストデータをクリーニング"""
        if self.data is None:
            return
        
        try:
            text_columns = self.data.select_dtypes(include=['object']).columns
            processed_cols = 0
            
            for col in text_columns:
                # 前後の空白を削除
                self.data[col] = self.data[col].astype(str).str.strip()
                # 連続する空白を1つに
                self.data[col] = self.data[col].str.replace(r'\s+', ' ', regex=True)
                processed_cols += 1
            
            self.update_data_info()
            self.update_data_table()
            self.log_action(f"テキストクリーニング: {processed_cols}列処理")
            self.update_status(f"✅ {processed_cols}列のテキストを整形しました")
            
        except Exception as e:
            self.show_error(f"テキストクリーニングエラー: {str(e)}")
    
    def convert_data_types(self):
        """データ型を変換"""
        if self.data is None:
            return
        
        try:
            converted_cols = 0
            
            for col in self.data.columns:
                if self.data[col].dtype == 'object':
                    # 数値変換を試行
                    numeric_series = pd.to_numeric(self.data[col], errors='coerce')
                    if numeric_series.notna().sum() / len(self.data[col]) > 0.8:
                        self.data[col] = numeric_series
                        converted_cols += 1
                    else:
                        # 日付変換を試行
                        try:
                            date_series = pd.to_datetime(self.data[col], errors='coerce')
                            if date_series.notna().sum() / len(self.data[col]) > 0.5:
                                self.data[col] = date_series
                                converted_cols += 1
                        except:
                            pass
            
            self.update_data_info()
            self.update_data_table()
            self.log_action(f"データ型変換: {converted_cols}列変換")
            self.update_status(f"✅ {converted_cols}列のデータ型を変換しました")
            
        except Exception as e:
            self.show_error(f"データ型変換エラー: {str(e)}")
    
    def fill_missing_dialog(self):
        """欠損値処理ダイアログを表示"""
        if self.data is None:
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("欠損値処理")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 処理方法選択
        ttk.Label(dialog, text="処理方法を選択:", font=('Arial', 10, 'bold')).pack(pady=10)
        
        method_var = tk.StringVar(value='remove')
        methods = [
            ('remove', '空行を削除'),
            ('forward', '前の値で埋める'),
            ('mean', '平均値で埋める（数値列のみ）'),
            ('median', '中央値で埋める（数値列のみ）'),
            ('zero', '0で埋める'),
            ('custom', 'カスタム値で埋める')
        ]
        
        for value, text in methods:
            ttk.Radiobutton(dialog, text=text, variable=method_var, value=value).pack(anchor=tk.W, padx=20)
        
        # カスタム値入力
        custom_frame = ttk.Frame(dialog)
        custom_frame.pack(pady=10)
        ttk.Label(custom_frame, text="カスタム値:").pack(side=tk.LEFT)
        custom_var = tk.StringVar()
        ttk.Entry(custom_frame, textvariable=custom_var).pack(side=tk.LEFT, padx=5)
        
        # ボタン
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def apply_fill():
            try:
                method = method_var.get()
                initial_nulls = self.data.isnull().sum().sum()
                
                if method == 'remove':
                    self.data = self.data.dropna()
                elif method == 'forward':
                    self.data = self.data.fillna(method='ffill')
                elif method == 'mean':
                    numeric_cols = self.data.select_dtypes(include=[np.number]).columns
                    self.data[numeric_cols] = self.data[numeric_cols].fillna(self.data[numeric_cols].mean())
                elif method == 'median':
                    numeric_cols = self.data.select_dtypes(include=[np.number]).columns
                    self.data[numeric_cols] = self.data[numeric_cols].fillna(self.data[numeric_cols].median())
                elif method == 'zero':
                    self.data = self.data.fillna(0)
                elif method == 'custom':
                    custom_value = custom_var.get()
                    self.data = self.data.fillna(custom_value)
                
                final_nulls = self.data.isnull().sum().sum()
                processed_count = initial_nulls - final_nulls
                
                self.update_data_info()
                self.update_data_table()
                self.log_action(f"欠損値処理: {processed_count}個処理 (方法: {method})")
                self.update_status(f"✅ {processed_count}個の欠損値を処理しました")
                
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("エラー", f"欠損値処理エラー: {str(e)}")
        
        ttk.Button(button_frame, text="適用", command=apply_fill).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def rename_columns_dialog(self):
        """列名変更ダイアログを表示"""
        if self.data is None:
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("列名変更")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 列リスト
        ttk.Label(dialog, text="変更する列を選択:", font=('Arial', 10, 'bold')).pack(pady=10)
        
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        
        listbox = tk.Listbox(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)
        
        for col in self.data.columns:
            listbox.insert(tk.END, col)
        
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 新しい名前入力
        input_frame = ttk.Frame(dialog)
        input_frame.pack(pady=10)
        ttk.Label(input_frame, text="新しい列名:").pack(side=tk.LEFT)
        new_name_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=new_name_var).pack(side=tk.LEFT, padx=5)
        
        # ボタン
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def apply_rename():
            try:
                selection = listbox.curselection()
                if not selection:
                    messagebox.showwarning("警告", "列を選択してください")
                    return
                
                old_name = listbox.get(selection[0])
                new_name = new_name_var.get().strip()
                
                if not new_name:
                    messagebox.showwarning("警告", "新しい列名を入力してください")
                    return
                
                self.data = self.data.rename(columns={old_name: new_name})
                
                self.update_data_info()
                self.update_data_table()
                self.log_action(f"列名変更: {old_name} -> {new_name}")
                self.update_status(f"✅ 列名を変更しました: {old_name} -> {new_name}")
                
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("エラー", f"列名変更エラー: {str(e)}")
        
        ttk.Button(button_frame, text="変更", command=apply_rename).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def filter_data_dialog(self):
        """データフィルタダイアログを表示"""
        if self.data is None:
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("データフィルタ")
        dialog.geometry("450x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 列選択
        ttk.Label(dialog, text="フィルタする列:", font=('Arial', 10, 'bold')).pack(pady=5)
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(dialog, textvariable=column_var, values=list(self.data.columns))
        column_combo.pack(pady=5)
        
        # 条件選択
        ttk.Label(dialog, text="条件:", font=('Arial', 10, 'bold')).pack(pady=5)
        condition_var = tk.StringVar(value='==')
        conditions = ['==', '!=', '>', '<', '>=', '<=', 'contains', 'startswith', 'endswith']
        condition_combo = ttk.Combobox(dialog, textvariable=condition_var, values=conditions)
        condition_combo.pack(pady=5)
        
        # 値入力
        ttk.Label(dialog, text="値:", font=('Arial', 10, 'bold')).pack(pady=5)
        value_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=value_var).pack(pady=5)
        
        # ボタン
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def apply_filter():
            try:
                column = column_var.get()
                condition = condition_var.get()
                value = value_var.get()
                
                if not column or column not in self.data.columns:
                    messagebox.showwarning("警告", "有効な列を選択してください")
                    return
                
                if not value:
                    messagebox.showwarning("警告", "フィルタ値を入力してください")
                    return
                
                initial_rows = len(self.data)
                
                if condition == '==':
                    self.data = self.data[self.data[column] == value]
                elif condition == '!=':
                    self.data = self.data[self.data[column] != value]
                elif condition == '>':
                    self.data = self.data[pd.to_numeric(self.data[column], errors='coerce') > float(value)]
                elif condition == '<':
                    self.data = self.data[pd.to_numeric(self.data[column], errors='coerce') < float(value)]
                elif condition == '>=':
                    self.data = self.data[pd.to_numeric(self.data[column], errors='coerce') >= float(value)]
                elif condition == '<=':
                    self.data = self.data[pd.to_numeric(self.data[column], errors='coerce') <= float(value)]
                elif condition == 'contains':
                    self.data = self.data[self.data[column].astype(str).str.contains(value, na=False)]
                elif condition == 'startswith':
                    self.data = self.data[self.data[column].astype(str).str.startswith(value, na=False)]
                elif condition == 'endswith':
                    self.data = self.data[self.data[column].astype(str).str.endswith(value, na=False)]
                
                filtered_rows = initial_rows - len(self.data)
                
                self.update_data_info()
                self.update_data_table()
                self.log_action(f"データフィルタ: {column} {condition} {value} ({filtered_rows}行除外)")
                self.update_status(f"✅ {filtered_rows}行をフィルタしました")
                
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("エラー", f"フィルタエラー: {str(e)}")
        
        ttk.Button(button_frame, text="適用", command=apply_filter).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def export_excel(self):
        """Excel形式で保存"""
        if self.data is None:
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Excelファイルとして保存",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    self.data.to_excel(writer, sheet_name='Data', index=False)
                    
                    # 処理ログも保存
                    if self.processing_log:
                        log_df = pd.DataFrame(self.processing_log, columns=['処理ログ'])
                        log_df.to_excel(writer, sheet_name='Log', index=False)
                
                self.log_action(f"Excel保存: {os.path.basename(file_path)}")
                self.update_status(f"✅ Excelファイルを保存しました: {os.path.basename(file_path)}")
                self.show_success(f"Excelファイルを保存しました\n{file_path}")
                
            except Exception as e:
                self.show_error(f"Excel保存エラー: {str(e)}")
    
    def export_csv(self):
        """CSV形式で保存"""
        if self.data is None:
            return
        
        file_path = filedialog.asksaveasfilename(
            title="CSVファイルとして保存",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.data.to_csv(file_path, index=False, encoding='utf-8-sig')
                
                self.log_action(f"CSV保存: {os.path.basename(file_path)}")
                self.update_status(f"✅ CSVファイルを保存しました: {os.path.basename(file_path)}")
                self.show_success(f"CSVファイルを保存しました\n{file_path}")
                
            except Exception as e:
                self.show_error(f"CSV保存エラー: {str(e)}")
    
    def reset_data(self):
        """データをリセット"""
        if self.original_data is None:
            return
        
        if messagebox.askyesno("確認", "データをリセットしますか？\nすべての変更が失われます。"):
            self.data = self.original_data.copy()
            self.update_data_info()
            self.update_data_table()
            self.processing_log = []
            self.log_action("データリセット完了")
            self.update_status("✅ データをリセットしました")
    
    def show_processing_log(self):
        """処理ログを表示"""
        if not self.processing_log:
            messagebox.showinfo("処理ログ", "処理ログがありません")
            return
        
        log_window = tk.Toplevel(self.root)
        log_window.title("処理ログ")
        log_window.geometry("600x400")
        
        text_widget = scrolledtext.ScrolledText(log_window, wrap=tk.WORD, font=('Consolas', 9))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        for log in self.processing_log:
            text_widget.insert(tk.END, log + '\n')
        
        text_widget.config(state='disabled')


def main():
    """メイン関数"""
    root = tk.Tk()
    app = TableauPreprocessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
