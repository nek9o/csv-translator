import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import deepl
import logging
import os
import time
import csv
import json
import threading
from dotenv import load_dotenv
from pathlib import Path
from typing import List, Optional

# Load environment variables
load_dotenv()

# Setup logging with custom handler to capture logs for GUI
class GUILogHandler(logging.Handler):
    def __init__(self, text_widget):
        logging.Handler.__init__(self)
        self.text_widget = text_widget
        
    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
            self.text_widget.configure(state='disabled')
        # Schedule append in the GUI thread
        self.text_widget.after(0, append)

class DeepLTranslator:
    def __init__(self, log_widget=None):
        self.auth_key = self.get_api_key()
        self.translator = deepl.Translator(self.auth_key)
        self.supported_languages = self.load_supported_languages()
        self.log_widget = log_widget
        self.progress_callback = None
        self.stop_translation = False

    @staticmethod
    def get_api_key() -> str:
        api_key = os.environ.get("DEEPL_AUTH_KEY")
        if not api_key:
            raise ValueError("DeepL APIキーが設定されていません。.envファイルを確認してください。")
        return api_key

    @staticmethod
    def load_supported_languages() -> List[dict]:
        try:
            with open("languages.json", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"言語ファイルの読み込みエラー: {str(e)}")
            return []

    def get_output_path(self, input_path: str) -> str:
        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        return str(Path(input_path).parent / f"output_{timestamp}.csv")

    def translate_csv_column(self, input_file: str, column_name: Optional[str] = None, 
                             column_index: Optional[int] = None, target_lang: str = "JA", 
                             has_header: bool = True, encoding: str = 'utf-8', log_interval: int = 10):
        output_file = self.get_output_path(input_file)
        self.stop_translation = False

        try:
            df = pd.read_csv(input_file, encoding=encoding, header=0 if has_header else None)
            if not has_header:
                df.columns = [f"Column_{i}" for i in range(len(df.columns))]

            target_column = self.determine_column(df, column_name, column_index)
            logging.info(f"'{target_column}' 列を{target_lang}に翻訳します...")

            translated_texts = []
            total = len(df)

            for i, text in enumerate(df[target_column], 1):
                if self.stop_translation:
                    logging.info("翻訳が中断されました。")
                    return False

                translated_texts.append(self.translate_text(text, target_lang))

                if i % log_interval == 0 or i == total:
                    progress = i / total
                    logging.info(f"進捗: {i}/{total} ({progress*100:.1f}%)")
                    if self.progress_callback:
                        self.progress_callback(progress)

            df[target_column] = translated_texts
            df.to_csv(output_file, index=False, encoding='utf-8', quoting=csv.QUOTE_ALL)

            logging.info(f"翻訳が完了し、結果を '{output_file}' に保存しました。")
            return output_file
        except Exception as e:
            logging.error(f"エラーが発生しました: {str(e)}")
            raise

    def determine_column(self, df: pd.DataFrame, column_name: Optional[str], column_index: Optional[int]) -> str:
        if column_name and column_name in df.columns:
            return column_name
        elif column_index is not None and 0 <= column_index < len(df.columns):
            return df.columns[column_index]
        else:
            raise ValueError("指定された列が見つかりません。")

    def translate_text(self, text: str, target_lang: str) -> str:
        if pd.isna(text) or not str(text).strip():
            return ""
        try:
            return self.translator.translate_text(str(text), target_lang=target_lang).text
        except deepl.DeepLException as e:
            logging.warning(f"DeepLエラー: {str(e)}")
            return text

    def set_progress_callback(self, callback):
        self.progress_callback = callback

    def stop(self):
        self.stop_translation = True


class DeepLTranslatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DeepL CSV翻訳ツール")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        self.create_widgets()
        self.translator = None
        self.translation_thread = None
        
        # ロギング設定
        self.setup_logging()
        
    def setup_logging(self):
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        
        # 既存のハンドラをクリア
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
            
        # コンソールハンドラ
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        console_handler.setFormatter(console_formatter)
        root_logger.addHandler(console_handler)
        
        # GUIハンドラ
        gui_handler = GUILogHandler(self.log_text)
        gui_handler.setLevel(logging.INFO)
        gui_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        gui_handler.setFormatter(gui_formatter)
        root_logger.addHandler(gui_handler)
    
    def create_widgets(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ファイル選択セクション
        file_frame = ttk.LabelFrame(main_frame, text="CSVファイル選択", padding="5")
        file_frame.pack(fill=tk.X, pady=5)
        
        self.file_path = tk.StringVar()
        ttk.Label(file_frame, text="ファイルパス:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path, width=50).grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(file_frame, text="参照...", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)
        
        # ヘッダーのチェックボックス
        self.has_header = tk.BooleanVar(value=True)
        ttk.Checkbutton(file_frame, text="ヘッダー行あり", variable=self.has_header).grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        
        # 列選択セクション
        column_frame = ttk.LabelFrame(main_frame, text="翻訳する列の選択", padding="5")
        column_frame.pack(fill=tk.X, pady=5)
        
        self.column_method = tk.StringVar(value="name")
        ttk.Radiobutton(column_frame, text="列名で指定", variable=self.column_method, value="name").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Radiobutton(column_frame, text="インデックスで指定", variable=self.column_method, value="index").grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        self.column_value = tk.StringVar()
        ttk.Label(column_frame, text="列名/インデックス:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(column_frame, textvariable=self.column_value, width=20).grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
        
        # 言語選択セクション
        lang_frame = ttk.LabelFrame(main_frame, text="翻訳設定", padding="5")
        lang_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(lang_frame, text="翻訳先言語:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.target_lang = ttk.Combobox(lang_frame, width=30)
        self.target_lang.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        
        # ログ間隔設定
        ttk.Label(lang_frame, text="ログ間隔 (行数):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.log_interval = tk.StringVar(value="10")
        ttk.Entry(lang_frame, textvariable=self.log_interval, width=10).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 実行ボタンセクション
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.start_button = ttk.Button(button_frame, text="翻訳開始", command=self.start_translation)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="中止", command=self.stop_translation, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # プログレスバー
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(progress_frame, text="進捗状況:").pack(side=tk.LEFT, padx=5)
        self.progress_bar = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress_bar.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # ログ表示エリア
        log_frame = ttk.LabelFrame(main_frame, text="ログ", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state='disabled')
        
        # 言語リストの読み込み
        self.load_languages()
        
        # カラム調整
        main_frame.columnconfigure(0, weight=1)
        file_frame.columnconfigure(1, weight=1)
        column_frame.columnconfigure(1, weight=1)
        lang_frame.columnconfigure(1, weight=1)
    
    def load_languages(self):
        try:
            translator = DeepLTranslator()
            languages = translator.supported_languages
            
            if languages:
                language_options = [f"{lang['code']} - {lang['name']}" for lang in languages]
                self.target_lang['values'] = language_options
                self.target_lang.current(languages.index(next(lang for lang in languages if lang['code'] == 'JA')))
            else:
                logging.error("言語リストを読み込めませんでした。")
        except Exception as e:
            logging.error(f"言語リスト読み込みエラー: {str(e)}")
    
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)
    
    def update_progress(self, value):
        self.progress_bar['value'] = value * 100
    
    def enable_controls(self, enabled=True):
        state = tk.NORMAL if enabled else tk.DISABLED
        disabled_state = tk.DISABLED if enabled else tk.NORMAL
        
        self.start_button['state'] = state
        self.stop_button['state'] = disabled_state
    
    def start_translation(self):
        if not self.file_path.get():
            messagebox.showerror("エラー", "CSVファイルを選択してください。")
            return
        
        if not self.column_value.get():
            messagebox.showerror("エラー", "翻訳する列を指定してください。")
            return
        
        try:
            # ログインターバルの検証
            log_interval = int(self.log_interval.get())
            if log_interval <= 0:
                messagebox.showerror("エラー", "ログ間隔は正の整数を指定してください。")
                return
        except ValueError:
            messagebox.showerror("エラー", "ログ間隔は整数を入力してください。")
            return
        
        # UI要素の状態変更
        self.enable_controls(False)
        self.progress_bar['value'] = 0
        
        # 翻訳オブジェクトの初期化
        self.translator = DeepLTranslator(self.log_text)
        self.translator.set_progress_callback(self.update_progress)
        
        # 翻訳パラメータの準備
        target_lang_code = self.target_lang.get().split(' - ')[0].strip()
        column_name = None
        column_index = None
        
        if self.column_method.get() == "name":
            column_name = self.column_value.get()
        else:
            try:
                column_index = int(self.column_value.get())
            except ValueError:
                messagebox.showerror("エラー", "列インデックスは整数を入力してください。")
                self.enable_controls(True)
                return
        
        # 翻訳処理を別スレッドで実行
        self.translation_thread = threading.Thread(
            target=self.run_translation,
            args=(
                self.file_path.get(),
                column_name,
                column_index,
                target_lang_code,
                self.has_header.get(),
                'utf-8',
                log_interval
            )
        )
        self.translation_thread.daemon = True
        self.translation_thread.start()
    
    def run_translation(self, input_file, column_name, column_index, target_lang, has_header, encoding, log_interval):
        try:
            output_file = self.translator.translate_csv_column(
                input_file, column_name, column_index, target_lang, 
                has_header, encoding, log_interval
            )
            
            if output_file:
                self.root.after(0, lambda: messagebox.showinfo("完了", f"翻訳が完了しました。\n出力ファイル: {output_file}"))
        except Exception as e:
            self.root.after(0, lambda: self.show_error(str(e)))
        finally:
            self.root.after(0, lambda: self.enable_controls(True))
    
    def stop_translation(self):
        if self.translator:
            self.translator.stop()
            logging.info("翻訳を中止しています...")
    
    def show_error(self, error_message):
        error_window = tk.Toplevel(self.root)
        error_window.title("エラー")
        error_window.geometry("500x300")
        
        error_text = scrolledtext.ScrolledText(error_window, wrap=tk.WORD)
        error_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        error_text.insert(tk.END, f"エラーが発生しました:\n\n{error_message}")
        error_text.configure(state='disabled')
        
        ttk.Button(error_window, text="閉じる", command=error_window.destroy).pack(pady=10)


if __name__ == "__main__":
    root = tk.Tk()
    app = DeepLTranslatorGUI(root)
    root.mainloop()
