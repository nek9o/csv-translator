import pandas as pd
import deepl
from dotenv import load_dotenv
import os
from pathlib import Path

# .envファイルを読み込む
load_dotenv()

def get_api_key() -> str:
    api_key = os.getenv("DEEPL_AUTH_KEY")
    if not api_key:
        raise ValueError("DeepL APIキーが設定されていません。.envファイルを確認してください。")
    return api_key

def get_output_path(input_path: str) -> str:
    input_path = Path(input_path)
    output_dir = input_path.parent
    output_name = 'output.csv'
    return str(output_dir / output_name)

def translate_csv_column(input_file: str, column_name: str = None, column_index: int = None, target_lang: str = "JA", has_header: bool = True) -> None:
    try:
        output_file = get_output_path(input_file)
        auth_key = get_api_key()
        translator = deepl.Translator(auth_key)
        
        # CSVファイルの読み込み（ヘッダーの有無に対応）
        header_option = 0 if has_header else None
        df = pd.read_csv(input_file, encoding='utf-8', header=header_option)
        
        # ヘッダーなしの場合、デフォルトの列名を設定
        if not has_header:
            df.columns = [f"Column_{i}" for i in range(len(df.columns))]
        
        # 翻訳対象の列を決定
        if column_name and column_name in df.columns:
            target_column = column_name
        elif column_index is not None and 0 <= column_index < len(df.columns):
            target_column = df.columns[column_index]
        else:
            raise ValueError("指定された列が見つかりません。列名またはインデックスを確認してください。")
        
        print(f"'{target_column}' 列を翻訳します...")
        translated_texts = []
        total = len(df)
        
        for i, text in enumerate(df[target_column], 1):
            if pd.isna(text) or str(text).strip() == "":
                translated_texts.append("")
                continue
            
            result = translator.translate_text(str(text), target_lang=target_lang)
            translated_texts.append(result.text)
            
            if i % 10 == 0 or i == total:
                print(f"進捗: {i}/{total} ({(i/total*100):.1f}%)")
        
        df[target_column] = translated_texts
        df.to_csv(output_file, index=False, encoding='utf-8')
        print(f"翻訳が完了し、結果を '{output_file}' に保存しました。")
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        raise

if __name__ == "__main__":
    input_path = input("翻訳したいCSVファイルのパスを入力してください: ").strip()
    if os.path.exists(input_path) and input_path.endswith('.csv'):
        has_header = input("CSVファイルにヘッダー行がありますか？ (y/n): ").strip().lower() == "y"
        method = input("列の指定方法を選択してください（1: 列名, 2: インデックス）: ").strip()
        if method == "1":
            column_name = input("翻訳する列名を入力してください: ").strip()
            translate_csv_column(input_path, column_name=column_name, has_header=has_header)
        elif method == "2":
            column_index = int(input("翻訳する列のインデックスを入力してください（0始まり）: ").strip())
            translate_csv_column(input_path, column_index=column_index, has_header=has_header)
        else:
            print("エラー: 無効な選択です。")
    else:
        print("エラー: 有効なCSVファイルを指定してください。")
