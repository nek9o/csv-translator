import pandas as pd
import deepl
import logging
import os
import time
import csv
import json
from dotenv import load_dotenv
from pathlib import Path
from typing import List, Optional

# Load environment variables
load_dotenv()

# Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logging.getLogger("deepl").setLevel(logging.WARNING)


class DeepLTranslator:
    def __init__(self):
        self.auth_key = self.get_api_key()
        self.translator = deepl.Translator(self.auth_key)
        self.supported_languages = self.load_supported_languages()

    @staticmethod
    def get_api_key() -> str:
        api_key = os.environ.get("DEEPL_AUTH_KEY")
        if not api_key:
            raise ValueError("DeepL APIキーが設定されていません。.envファイルを確認してください。")
        return api_key

    @staticmethod
    def load_supported_languages() -> List[dict]:
        with open("languages.json", encoding="utf-8") as f:
            return json.load(f)

    def get_output_path(self, input_path: str) -> str:
        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        return str(Path(input_path).parent / f"output_{timestamp}.csv")

    def show_language_codes(self) -> None:
        codes = [lang['code'] for lang in self.supported_languages]
        print("Supported language codes:", ", ".join(codes))
        
    def show_supported_languages(self) -> None:
        print("利用可能な言語一覧:")
        print("------------------")
        print("コード | 言語名")
        print("------------------")
        for lang in self.supported_languages:
            print(f"{lang['code']:<6} | {lang['name']}")
        print("------------------")

    def translate_csv_column(self, input_file: str, column_name: Optional[str] = None, 
                             column_index: Optional[int] = None, target_lang: str = "JA", 
                             has_header: bool = True, encoding: str = 'utf-8', log_interval: int = 10) -> None:
        output_file = self.get_output_path(input_file)

        try:
            df = pd.read_csv(input_file, encoding=encoding, header=0 if has_header else None)
            if not has_header:
                df.columns = [f"Column_{i}" for i in range(len(df.columns))]

            target_column = self.determine_column(df, column_name, column_index)
            logging.info(f"'{target_column}' 列を{target_lang}に翻訳します...")

            translated_texts = []
            total = len(df)

            for i, text in enumerate(df[target_column], 1):
                translated_texts.append(self.translate_text(text, target_lang))

                if i % log_interval == 0 or i == total:
                    logging.info(f"進捗: {i}/{total} ({(i/total*100):.1f}%)")

            df[target_column] = translated_texts
            df.to_csv(output_file, index=False, encoding='utf-8', quoting=csv.QUOTE_ALL)

            logging.info(f"翻訳が完了し、結果を '{output_file}' に保存しました。")
        except Exception as e:
            logging.error(f"エラーが発生しました: {str(e)}", exc_info=True)

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


if __name__ == "__main__":
    translator = DeepLTranslator()

    print("DeepL翻訳ツール")
    print("1: CSVファイルの列を翻訳する")
    print("2: 利用可能な言語一覧を表示する")
    print("0: 終了")
    
    choice = input("選択してください: ").strip()
    
    if choice == "1":
        # 1. ファイルパスの入力
        input_path = input("翻訳したいCSVファイルのパスを入力してください: ").strip()
        if not os.path.exists(input_path) or not input_path.endswith('.csv'):
            logging.error("エラー: 有効なCSVファイルを指定してください。")
            exit(1)
            
        # CSVファイルを読み込む前にヘッダーの有無を確認
        has_header = input("CSVファイルにヘッダー行がありますか？ (y/n): ").strip().lower() == "y"
        
        # 2. 列の指定方法
        method = input("列の指定方法を選択してください（1: 列名, 2: インデックス）: ").strip()
        
        # 3. 列の指定
        column_name = None
        column_index = None
        
        if method == "1":
            column_name = input("翻訳する列名を入力してください: ").strip()
        elif method == "2":
            try:
                column_index = int(input("翻訳する列のインデックスを入力してください（0始まり）: ").strip())
            except ValueError:
                logging.error("エラー: インデックスは整数で入力してください。")
                exit(1)
        else:
            logging.error("エラー: 無効な選択です。")
            exit(1)
        
        # 4. 言語一覧の表示と言語コードの指定
        translator.show_supported_languages()
        target_lang = input("翻訳先の言語コードを入力してください (例: JA): ").strip().upper()
        
        # 5. ログの間隔の指定
        try:
            log_interval = int(input("進捗を表示する間隔（行数）を入力してください (例: 10): ").strip())
        except ValueError:
            logging.warning("警告: 入力が無効です。デフォルト値(10)を使用します。")
            log_interval = 10

        # 翻訳実行
        translator.translate_csv_column(
            input_path, 
            column_name=column_name, 
            column_index=column_index, 
            target_lang=target_lang, 
            has_header=has_header, 
            log_interval=log_interval
        )
            
    elif choice == "2":
        translator.show_supported_languages()
    elif choice == "0":
        print("プログラムを終了します。")
    else:
        logging.error("エラー: 無効な選択です。")
