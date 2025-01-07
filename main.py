import argparse
import datetime
import glob
import traceback
import os

import pandas as pd
from tqdm import tqdm
from docxcompose.composer import Composer
from docx import Document

# 定数定義
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(ROOT_DIR, "input")
OUTPUT_DIR = os.path.join(ROOT_DIR, "output")
ERROR_DIR = os.path.join(ROOT_DIR, "error")
DATA = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
NOW = datetime.datetime.now().strftime("%Y%m%d%H%M%S")

# ログ書き込み
def write_log(output, log):
    os.makedirs(output, exist_ok=True)
    with open(os.path.join(output, f"log_{NOW}.txt"), mode="a", encoding='utf-8') as f:
        f.write(log + "\n")

# エラーログ書き込み
def write_error_log(error):
    os.makedirs(ERROR_DIR, exist_ok=True)
    with open(os.path.join(ERROR_DIR, f"error_{NOW}.txt"), mode="a", encoding='utf-8') as f:
        f.write(error + "\n")

# 絶対パス取得
def get_absolute_path(relative_path):
    return os.path.abspath(os.path.join(ROOT_DIR, relative_path))

def main(args):
    # エクセルファイルの存在確認
    if not os.path.isfile(args.data):
        print(f"エクセルファイルが存在しません: {args.data}")
        return

    # 出力ディレクトリ作成
    os.makedirs(args.output_dir, exist_ok=True)
    os.makedirs(os.path.join(args.output_dir, "docx"), exist_ok=True)

    # エクセルファイルを読み込み
    df = pd.read_excel(args.data, dtype=str)
    df = df.fillna('')

    # 各ファイルを結合
    for data in tqdm(df.values, desc='結合中...', total=len(df.values)):
        output_path = os.path.join(args.output_dir, "docx", data[0])
        inputs_path = [os.path.abspath(os.path.join(args.input_dir, "docx", path)) for path in data[1:]]

        # ファイルの存在確認
        not_exist = [path for path in inputs_path if not os.path.isfile(path)]
        if not_exist:
            missing_files = ", ".join(not_exist)
            write_log(args.output_dir, f"{data[0]} エラー：指定されたファイルが存在しない ({missing_files})")
            continue

        try:
            # DOCXファイルを結合
            master = Document(inputs_path[0])
            master.add_page_break()
            composer = Composer(master)
            for path in inputs_path[1:]:
                doc = Document(path)
                composer.append(doc)
                master.add_page_break()
            composer.save(output_path)
        except Exception as e:
            error_message = f"エラー: {str(e)}"
            write_log(args.output_dir, f"{data[0]} エラー：結合失敗 ({error_message})")
            write_error_log(traceback.format_exc())
            continue

        write_log(args.output_dir, f"{data[0]}  結合成功")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MS Word ドキュメントの結合")
    parser.add_argument('-i', '--input_dir', default=INPUT_DIR, help="入力ディレクトリのパス")
    parser.add_argument('-o', '--output_dir', default=OUTPUT_DIR, help="出力ディレクトリのパス")

    launch = True

    try:
        parser.add_argument('-d', '--data', default=DATA[0], help="インポートするエクセルファイルを指定")
    except Exception:
        write_error_log(traceback.format_exc())
        print('エラー: inputディレクトリにエクセルファイルが見つかりません。')
        input('Enterを押すと終了します。')
        launch = False

    args = parser.parse_args()

    if launch:
        main(args)