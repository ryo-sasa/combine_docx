import pandas as pd

def left_align_excel(file_path, output_path):
    # Excelファイルを読み込む
    df = pd.read_excel(file_path, header=None)
    
    # NaNを除外してすべての値を左に寄せる
    df_left_aligned = pd.DataFrame(df.apply(lambda x: pd.Series(x.dropna().values), axis=1))
    
    # 結果をExcelファイルとして保存
    df_left_aligned.to_excel(output_path, index=False, header=False)

# 使用例
input_file = 'input.xlsx'
output_file = 'output.xlsx'
left_align_excel(input_file, output_file)