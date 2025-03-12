import pandas as pd
import deepl
import logging
import os
import time
import csv
from dotenv import load_dotenv
from pathlib import Path

# Load environment variables
load_dotenv()

# Setup logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

logging.getLogger("deepl").setLevel(logging.WARNING)

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

def get_supported_languages() -> list:
    """DeepLがサポートする言語コードと名前のリストを返す"""
    languages = [
        ("BG", "ブルガリア語"),
        ("CS", "チェコ語"),
        ("DA", "デンマーク語"),
        ("DE", "ドイツ語"),
        ("EL", "ギリシャ語"),
        ("EN-GB", "英語（イギリス）"),
        ("EN-US", "英語（アメリカ）"),
        ("ES", "スペイン語"),
        ("ET", "エストニア語"),
        ("FI", "フィンランド語"),
        ("FR", "フランス語"),
        ("HU", "ハンガリー語"),
        ("ID", "インドネシア語"),
        ("IT", "イタリア語"),
        ("JA", "日本語"),
        ("KO", "韓国語"),
        ("LT", "リトアニア語"),
        ("LV", "ラトビア語"),
        ("NB", "ノルウェー語（ブークモール）"),
        ("NL", "オランダ語"),
        ("PL", "ポーランド語"),
        ("PT-BR", "ポルトガル語（ブラジル）"),
        ("PT-PT", "ポルトガル語（ポルトガル）"),
        ("RO", "ルーマニア語"),
        ("RU", "ロシア語"),
        ("SK", "スロバキア語"),
        ("SL", "スロベニア語"),
        ("SV", "スウェーデン語"),
        ("TR", "トルコ語"),
        ("UK", "ウクライナ語"),
        ("ZH", "中国語")
    ]
    return languages

def select_target_language() -> str:
    """ユーザーに翻訳先言語を選択させる"""
    languages = get_supported_languages()
    
    print("\n利用可能な言語:")
    for i, (code, name) in enumerate(languages, 1):
        print(f"{i}: {name} ({code})")
    
    while True:
        try:
            selection = int(input("\n翻訳先の言語番号を選択してください: ").strip())
            if 1 <= selection <= len(languages):
                return languages[selection-1][0]
            else:
                print(f"1から{len(languages)}までの数字を入力してください。")
        except ValueError:
            print("有効な数字を入力してください。")

def translate_csv_column(input_file: str, column_name: str = None, column_index: int = None, target_lang: str = "JA", has_header: bool = True) -> None:
    """DeepL APIを使用してCSVファイルの特定の列を翻訳"""
    try:
        output_file = get_output_path(input_file)
        auth_key = get_api_key()
        translator = deepl.Translator(auth_key)

        # Load CSV file
        header_option = 0 if has_header else None
        try:
            df = pd.read_csv(input_file, encoding='utf-8', header=header_option)
        except FileNotFoundError:
            logging.error("指定されたCSVファイルが見つかりません。")
            return
        except pd.errors.EmptyDataError:
            logging.error("CSVファイルが空です。")
            return
        
        if not has_header:
            df.columns = [f"Column_{i}" for i in range(len(df.columns))]

        # Determine target column
        if column_name and column_name in df.columns:
            target_column = column_name
        elif column_index is not None and 0 <= column_index < len(df.columns):
            target_column = df.columns[column_index]
        else:
            logging.error("指定された列が見つかりません。")
            return

        # 選択された言語コードに対応する言語名を取得
        language_name = next((name for code, name in get_supported_languages() if code == target_lang), target_lang)
        logging.info(f"'{target_column}' 列を{language_name}に翻訳します...")
        
        translated_texts = []
        total = len(df)

        for i, text in enumerate(df[target_column], 1):
            if pd.isna(text) or str(text).strip() == "":
                translated_texts.append("")  # Keep empty cells
                continue

            try:
                result = translator.translate_text(str(text), target_lang=target_lang)
                translated_texts.append(result.text)
                time.sleep(0.5)  # Prevent API throttling
            except deepl.DeepLException as e:
                logging.warning(f"警告: 行 {i} の翻訳中にエラーが発生しました: {str(e)}")
                translated_texts.append(str(text))

            if i % 10 == 0 or i == total:
                logging.info(f"進捗: {i}/{total} ({(i/total*100):.1f}%)")

        df[target_column] = translated_texts
        df.to_csv(output_file, index=False, encoding='utf-8', quoting=csv.QUOTE_ALL)
        logging.info(f"翻訳が完了し、結果を '{output_file}' に保存しました。")

    except Exception as e:
        logging.error(f"エラーが発生しました: {str(e)}", exc_info=True)

if __name__ == "__main__":
    input_path = input("翻訳したいCSVファイルのパスを入力してください: ").strip()
    
    if not os.path.exists(input_path) or not input_path.endswith('.csv'):
        logging.error("エラー: 有効なCSVファイルを指定してください。")
    else:
        has_header = input("CSVファイルにヘッダー行がありますか？ (y/n): ").strip().lower() == "y"
        method = input("列の指定方法を選択してください（1: 列名, 2: インデックス）: ").strip()
        
        # 翻訳先言語の選択
        target_lang = select_target_language()
        
        if method == "1":
            column_name = input("翻訳する列名を入力してください: ").strip()
            translate_csv_column(input_path, column_name=column_name, target_lang=target_lang, has_header=has_header)
        elif method == "2":
            try:
                column_index = int(input("翻訳する列のインデックスを入力してください（0始まり）: ").strip())
                translate_csv_column(input_path, column_index=column_index, target_lang=target_lang, has_header=has_header)
            except ValueError:
                logging.error("エラー: インデックスは整数で入力してください。")
        else:
            logging.error("エラー: 無効な選択です。")
