import argparse
import datetime
import glob
import traceback
import os

import pandas as pd
from tqdm import tqdm
from docxcompose.composer import Composer
from docx import Document

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(ROOT_DIR, "input")
OUTPUT_DIR = os.path.join(ROOT_DIR, "output")
ERROR_DIR = os.path.join(ROOT_DIR, "error")
DATA = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
NOW = datetime.datetime.now().strftime("%Y%m%d%H%M%S")


# write log
def write_log(output, log):
    os.makedirs(output, exist_ok=True)
    with open(os.path.join(output, f"log_{NOW}.txt"), mode="a", encoding='utf-8') as f:
        f.write(log + "\n")


# write error log
def write_error_log(error):
    os.makedirs(ERROR_DIR, exist_ok=True)
    with open(os.path.join(ERROR_DIR, f"error_{NOW}.txt"), mode="a", encoding='utf-8') as f:
        f.write(error + "\n")


def main(args):
    # 大問確定用シートの存在確認
    if not os.path.isfile(args.data):
        print("インプットエクセルが存在しません．")
        return

    # アウトプットフォルダを作成
    os.makedirs(args.output_dir, exist_ok=True)
    os.makedirs(os.path.join(args.output_dir, "docx"), exist_ok=True)

    df = pd.read_excel(args.data, dtype=str)
    df = df.fillna('')

    for data in tqdm(df.values, desc='結合中...', total=len(df.values)):
        output_path = os.path.join(args.output_dir, "docx", data[0])
        inputs_path = [os.path.join(args.input_dir, "docx", path) for path in data[1:]]

        exists = True
        not_exist = []
        for i, path in enumerate(inputs_path):
            if not os.path.isfile(path):
                exists = False
                not_exist.append(i)
        if not exists:
            write_log(args.output_dir, f"{data[0]} エラー：指定されたファイルが存在しない（{not_exist}番目）")
            continue

        try:
            master = Document(inputs_path[0])
            master.add_page_break()
            composer = Composer(master)
            # composer.save(output_path)
            for path in inputs_path[1:]:
                doc = Document(path)
                composer.append(doc)
                master.add_page_break()
            composer.save(output_path)
        except Exception:
            write_log(args.output_dir, f"{data[0]} エラー：結合失敗")
            write_error_log(traceback.format_exc())
            continue

        write_log(args.output_dir, f"{data[0]}  結合成功")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="combine MS word document")
    parser.add_argument('-i', '--input_dir', default=INPUT_DIR, help="inputディレクトリのパスを指定できる")
    parser.add_argument('-o', '--output_dir', default=OUTPUT_DIR, help="outputディレクトリのパスを指定できる")

    launch = True

    try:
        parser.add_argument('-d', '--data', default=DATA[0], help="インポートエクセルを指定する")
    except Exception:
        save_error_log(traceback.format_exc())
        print('エラー: inputにエクセルファイルが見つかりません。')
        input('Enterを押したら終了します。')
        launch = False

    args = parser.parse_args()

    if launch:
        main(args)
