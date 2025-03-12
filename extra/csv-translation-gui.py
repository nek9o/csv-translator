import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import deepl
import os
from dotenv import load_dotenv
from pathlib import Path
import time
import csv
import threading

# Load environment variables
load_dotenv()

def get_api_key() -> str:
    """Retrieve DeepL API key from environment variables."""
    api_key = os.environ.get("DEEPL_AUTH_KEY")
    if not api_key:
        raise ValueError("DeepL APIキーが設定されていません。.envファイルを確認してください。")
    return api_key

def get_output_path(input_path: str) -> str:
    """Generate output CSV file path with timestamp to avoid overwriting."""
    input_path = Path(input_path)
    timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
    return str(input_path.parent / f"output_{timestamp}.csv")

def translate_csv(input_file, column_name, column_index, target_lang, has_header):
    """Translate a CSV column using DeepL API."""
    try:
        output_file = get_output_path(input_file)
        auth_key = get_api_key()
        translator = deepl.Translator(auth_key)

        header_option = 0 if has_header else None
        df = pd.read_csv(input_file, encoding='utf-8', header=header_option)
        if not has_header:
            df.columns = [f"Column_{i}" for i in range(len(df.columns))]

        if not column_name and column_index is None:
            messagebox.showerror("エラー", "列名またはインデックスを指定してください。")
            return

        target_column = column_name if column_name else df.columns[column_index]
        translated_texts = []

        progress_var.set(0)
        progress_bar.update()

        def process_translation():
            total_rows = len(df[target_column])
            for i, text in enumerate(df[target_column]):
                if pd.isna(text) or str(text).strip() == "":
                    translated_texts.append("")
                else:
                    try:
                        result = translator.translate_text(str(text), target_lang=target_lang)
                        translated_texts.append(result.text)
                    except deepl.DeepLException:
                        translated_texts.append(str(text))
                
                progress_var.set((i + 1) / total_rows * 100)
                root.after(100, progress_bar.update())  # Allow UI updates
            
            df[target_column] = translated_texts
            df.to_csv(output_file, index=False, encoding='utf-8', quoting=csv.QUOTE_ALL)
            messagebox.showinfo("完了", f"翻訳が完了しました。結果は {output_file} に保存されました。")
        
        threading.Thread(target=process_translation, daemon=True).start()
    except Exception as e:
        messagebox.showerror("エラー", str(e))

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)

def start_translation():
    input_file = entry_file.get()
    column_name = entry_column.get() if var_method.get() == 1 else None
    column_index = int(entry_index.get()) if var_method.get() == 2 and entry_index.get().isdigit() else None
    target_lang = lang_var.get()
    has_header = var_header.get()
    
    if not input_file:
        messagebox.showerror("エラー", "CSVファイルを選択してください。")
        return
    
    threading.Thread(target=translate_csv, args=(input_file, column_name, column_index, target_lang, has_header), daemon=True).start()

# DeepL言語コードと表示名の辞書
LANGUAGE_DICT = {
    "BG": "ブルガリア語",
    "CS": "チェコ語",
    "DA": "デンマーク語",
    "DE": "ドイツ語",
    "EL": "ギリシャ語",
    "EN-GB": "英語（イギリス）",
    "EN-US": "英語（アメリカ）",
    "EN": "英語",
    "ES": "スペイン語",
    "ET": "エストニア語",
    "FI": "フィンランド語",
    "FR": "フランス語",
    "HU": "ハンガリー語",
    "ID": "インドネシア語",
    "IT": "イタリア語",
    "JA": "日本語",
    "KO": "韓国語",
    "LT": "リトアニア語",
    "LV": "ラトビア語",
    "NB": "ノルウェー語（ブークモール）",
    "NL": "オランダ語",
    "PL": "ポーランド語",
    "PT-BR": "ポルトガル語（ブラジル）",
    "PT-PT": "ポルトガル語（ポルトガル）",
    "PT": "ポルトガル語",
    "RO": "ルーマニア語",
    "RU": "ロシア語",
    "SK": "スロバキア語",
    "SL": "スロベニア語",
    "SV": "スウェーデン語",
    "TR": "トルコ語",
    "UK": "ウクライナ語",
    "ZH": "中国語（簡体字）"
}

# GUI setup
root = tk.Tk()
root.title("CSV翻訳ツール")

tk.Label(root, text="CSVファイル").grid(row=0, column=0, padx=5, pady=5)
entry_file = tk.Entry(root, width=50)
entry_file.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="参照", command=select_file).grid(row=0, column=2, padx=5, pady=5)

var_header = tk.BooleanVar(value=True)
tk.Checkbutton(root, text="ヘッダーあり", variable=var_header).grid(row=1, column=1)

var_method = tk.IntVar(value=1)
tk.Radiobutton(root, text="列名指定", variable=var_method, value=1).grid(row=2, column=0)
tk.Radiobutton(root, text="インデックス指定", variable=var_method, value=2).grid(row=3, column=0)

entry_column = tk.Entry(root)
entry_column.grid(row=2, column=1, padx=5, pady=5)
entry_index = tk.Entry(root)
entry_index.grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="翻訳先言語").grid(row=4, column=0, padx=5, pady=5)
lang_var = tk.StringVar(value="JA")

# 言語コードと表示名を結合してコンボボックスの項目を作成
language_display = [f"{code} - {name}" for code, name in LANGUAGE_DICT.items()]
language_dropdown = ttk.Combobox(root, textvariable=lang_var, values=language_display, state="readonly", width=30)
language_dropdown.grid(row=4, column=1, padx=5, pady=5)

# コンボボックスで選択時に言語コードのみを取得する関数
def on_language_select(event):
    selected = language_dropdown.get()
    code = selected.split(" - ")[0]
    lang_var.set(code)

language_dropdown.bind("<<ComboboxSelected>>", on_language_select)

# 日本語をデフォルト選択
for i, item in enumerate(language_display):
    if item.startswith("JA"):
        language_dropdown.current(i)
        break

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.grid(row=6, column=1, padx=5, pady=5, sticky='ew')

tk.Button(root, text="翻訳開始", command=start_translation).grid(row=5, column=1, padx=5, pady=10)

root.mainloop()
