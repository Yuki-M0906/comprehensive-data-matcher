# -*- coding: utf-8 -*-
"""
サンプルデータ生成スクリプト
GitHub公開用のダミーデータ（examples/ の Excel）を生成します。
中身は sample_data.py にあります（Web 版の「見本で試す」と同じ）。
"""

from pathlib import Path

import pandas as pd

from sample_data import APPAREL_MASTER, APPAREL_VARIANTS

# 出力先ディレクトリ
OUTPUT_DIR = Path(__file__).parent / "examples"


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)

    df_correct = pd.DataFrame({'商品名': APPAREL_MASTER})
    df_yuragi = pd.DataFrame({'商品名': APPAREL_VARIANTS})

    # Excelファイルとして保存
    correct_file = OUTPUT_DIR / "correct_sample.xlsx"
    yuragi_file = OUTPUT_DIR / "yuragi_sample.xlsx"

    df_correct.to_excel(correct_file, sheet_name='Sheet1', index=False, engine='openpyxl')
    df_yuragi.to_excel(yuragi_file, sheet_name='Sheet1', index=False, engine='openpyxl')

    print(" サンプルデータを生成しました:")
    print(f"  - {correct_file}")
    print(f"  - {yuragi_file}")
    print("\n データ件数:")
    print(f"  - 標準商品名: {len(df_correct)}件")
    print(f"  - ゆらぎデータ: {len(df_yuragi)}件")


if __name__ == "__main__":
    main()
