# -*- coding: utf-8 -*-
"""
Comprehensive Data Matcher

高精度な名寄せツール。レーベンシュタイン距離とJaccard係数を
組み合わせたハイブリッドスコアリング方式を使用。

Author: Yuki0906
Version: 2.2.0
License: MIT
GitHub: https://github.com/Yuki-M0906/comprehensive-data-matcher

import pandas as pd
import unicodedata
from rapidfuzz import fuzz, distance
import logging
from pathlib import Path
from datetime import datetime
import sys
import re # 単語分割用にreをインポート

# --- 設定項目 ---
# スクリプトのあるディレクトリを基準とする
BASE_DIR = Path(__file__).parent.resolve()

# 入力ファイル
YURAGI_FILE = BASE_DIR / "yuragi.xlsx"
CORRECT_FILE = BASE_DIR / "correct.xlsx"

# 出力ファイル
RESULT_FILE = BASE_DIR / "result_combined_high_accuracy.xlsx"
LOG_FILE = BASE_DIR / "matcher_log_high_accuracy.txt"

# Excelシートと列の設定
YURAGI_SHEET_NAME = "Sheet1"
CORRECT_SHEET_NAME = "Sheet1"
RESULT_SHEET_NAME = "マッチング結果"

YURAGI_HEADER_ROW = 0
CORRECT_HEADER_ROW = 0

YURAGI_COL_NAME = '商品名'
CORRECT_COL_NAME = '商品名'

# ▼▼▼【高精度化設定】スコアの重み ▼▼▼
# 2つの重みの合計が1.0になるように調整
# Levenshtein: 誤字脱字など、文字レベルの一致度を重視する場合に高くする
WEIGHT_LEVENSHTEIN = 0.7
# Jaccard: 語順が違ってもOKとする場合に高くする
WEIGHT_JACCARD = 0.3
# ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

# --- ロガー設定 ---
log_formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
if logger.hasHandlers():
    logger.handlers.clear()
try:
    file_handler = logging.FileHandler(LOG_FILE, encoding='utf-8', mode='w')
    file_handler.setFormatter(log_formatter)
    logger.addHandler(file_handler)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(log_formatter)
    logger.addHandler(stream_handler)
except Exception as e:
    print(f"致命的エラー: ログファイルハンドラの設定に失敗しました: {e}", file=sys.stderr)
    sys.exit(1)

# --- 正規化関数  ---
def normalize_cell_nfkc(value):
    if isinstance(value, str):
        return unicodedata.normalize('NFKC', value)
    return value

def normalize_for_comparison(s: str) -> str:
    if not isinstance(s, str):
        s = str(s)
    # 大文字化のみ。スペースはJaccard係数で使うため残す
    return s.upper()

def tokenize(s: str) -> set:
    """文字列を単語の集合に分割する"""
    # 半角/全角スペース、スラッシュなどで分割
    tokens = re.split(r'[ 　/]', s)
    # 空の要素を除外して集合で返す
    return set(filter(None, tokens))

def calculate_hybrid_score(s1: str, s2: str) -> dict:
    """レーベンシュタイン類似度とJaccard係数からハイブリッドスコアを計算する"""
    # 1. レーベンシュタイン類似度 (0-100のスコア、100が完全一致)
    #   rapidfuzz.fuzz.ratioは内部で正規化されたレーベンシュタイン距離を計算
    lev_score = fuzz.ratio(s1, s2) / 100.0 # 0.0-1.0の範囲に正規化

    # 2. Jaccard係数
    tokens1 = tokenize(s1)
    tokens2 = tokenize(s2)
    
    if not tokens1 and not tokens2:
        jac_score = 1.0 # 両方とも空なら完全一致
    elif not tokens1 or not tokens2:
        jac_score = 0.0 # 片方だけ空なら不一致
    else:
        intersection = len(tokens1.intersection(tokens2))
        union = len(tokens1.union(tokens2))
        jac_score = intersection / union

    # 3. ハイブリッドスコアの計算
    hybrid_score = (lev_score * WEIGHT_LEVENSHTEIN) + (jac_score * WEIGHT_JACCARD)

    return {
        "hybrid": hybrid_score,
        "levenshtein": lev_score,
        "jaccard": jac_score
    }

def load_excel(file_path, sheet_name, header):
    if not file_path.exists():
        logger.error(f"ファイルが見つかりません: {file_path}")
        return None
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header, engine='openpyxl')
        logger.info(f"ファイル読み込み完了: {file_path} ({len(df)}行)")
        return df
    except Exception as e:
        logger.error(f"ファイルの読み込み中にエラーが発生しました: {file_path}", exc_info=True)
        return None

def main():
    logger.info("高精度データマッチング処理を開始します。")
    logger.info(f"スコア重み設定: Levenshtein={WEIGHT_LEVENSHTEIN}, Jaccard={WEIGHT_JACCARD}")

    # --- 1. ファイル読み込み (変更なし) ---
    df_yuragi_orig = load_excel(YURAGI_FILE, YURAGI_SHEET_NAME, YURAGI_HEADER_ROW)
    df_correct = load_excel(CORRECT_FILE, CORRECT_SHEET_NAME, CORRECT_HEADER_ROW)
    if df_yuragi_orig is None or df_correct is None: sys.exit(1)
    if YURAGI_COL_NAME not in df_yuragi_orig.columns:
        logger.error(f"'{YURAGI_FILE}' に必須列 '{YURAGI_COL_NAME}' が見つかりません。")
        sys.exit(1)
    if CORRECT_COL_NAME not in df_correct.columns:
        logger.error(f"'{CORRECT_FILE}' に必須列 '{CORRECT_COL_NAME}' が見つかりません。")
        sys.exit(1)

    # --- 2. 前処理 ---
    logger.info("--- ステップ2: yuragiデータのNFKC正規化 ---")
    df_yuragi_normalized = df_yuragi_orig.copy()
    df_yuragi_normalized = df_yuragi_normalized.applymap(normalize_cell_nfkc)
    logger.info("NFKC正規化が完了しました。")

    # --- 3. 名寄せ処理 ---
    logger.info("--- ステップ3: ハイブリッドスコアによる名寄せ処理 ---")
    yuragi_names_orig = df_yuragi_orig[YURAGI_COL_NAME].astype(str).tolist()
    yuragi_names_to_compare = df_yuragi_normalized[YURAGI_COL_NAME].astype(str).tolist()
    correct_names = df_correct[CORRECT_COL_NAME].astype(str).tolist()

    logger.info("標準商品名リストの正規化キャッシュを作成中...")
    correct_cache = [
        {"original": s, "normalized": normalize_for_comparison(s)}
        for s in correct_names
    ]
    logger.info("キャッシュ作成完了。")

    results_list = []
    total_count = len(yuragi_names_to_compare)
    logger.info(f"マッチング処理を開始します... (対象: {total_count}件)")

    for i, name_to_compare in enumerate(yuragi_names_to_compare):
        normalized_yuragi = normalize_for_comparison(name_to_compare)
        
        if (i + 1) % 100 == 0 or (i + 1) == total_count:
            logger.info(f"処理中: {i + 1}/{total_count}")

        best_match_original = None
        best_score_info = {"hybrid": -1.0} # スコアは高いほど良いので-1で初期化

        for std_entry in correct_cache:
            std_normalized = std_entry["normalized"]
            
            # ハイブリッドスコアを計算
            current_score_info = calculate_hybrid_score(normalized_yuragi, std_normalized)
            
            # 最高スコアを更新
            if current_score_info["hybrid"] > best_score_info["hybrid"]:
                best_score_info = current_score_info
                best_match_original = std_entry["original"]
                # スコアが1.0なら完全一致なので探索終了
                if best_score_info["hybrid"] >= 1.0:
                    break
        
        # マッチング結果を詳細に出力
        result_message = (
            f"  [マッチング結果] 「{yuragi_names_orig[i]}」=>「{best_match_original}」 "
            f"(Hybrid: {best_score_info['hybrid']:.2f}, "
            f"Lev: {best_score_info.get('levenshtein', 0):.2f}, "
            f"Jac: {best_score_info.get('jaccard', 0):.2f})"
        )
        logger.info(result_message)
        
        results_list.append({
            '元のゆらぎ名': yuragi_names_orig[i],
            'マッチした正規名': best_match_original,
            'ハイブリッドスコア': best_score_info['hybrid'],
            'レーベンシュタイン類似度': best_score_info.get('levenshtein'),
            'Jaccard係数': best_score_info.get('jaccard'),
        })

    logger.info("マッチング処理が完了しました。")

    # --- 4. 結果の出力 ---
    logger.info("--- ステップ4: 処理結果のファイル出力 ---")
    try:
        df_results = pd.DataFrame(results_list)
        # スコアを小数点以下3桁でフォーマット
        df_results['ハイブリッドスコア'] = df_results['ハイブリッドスコア'].map('{:.3f}'.format)
        df_results['レーベンシュタイン類似度'] = df_results['レーベンシュタイン類似度'].map('{:.3f}'.format)
        df_results['Jaccard係数'] = df_results['Jaccard係数'].map('{:.3f}'.format)

        df_results.to_excel(RESULT_FILE, sheet_name=RESULT_SHEET_NAME, index=False, engine='openpyxl')
        logger.info(f"結果を正常に保存しました: {RESULT_FILE}")
    except Exception as e:
        logger.error(f"結果ファイルの保存中にエラーが発生しました。", exc_info=True)
        sys.exit(1)

    logger.info(" すべての処理が正常に完了しました。")

if __name__ == "__main__":
    start_time = datetime.now()
    try:
        main()
    except Exception as e:
        logger.critical("予期せぬエラーによりスクリプトが異常終了しました。", exc_info=True)
    finally:
        end_time = datetime.now()

        logger.info(f"総実行時間: {end_time - start_time}")
