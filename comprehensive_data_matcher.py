# -*- coding: utf-8 -*-
"""
Comprehensive Data Matcher

高精度な名寄せツール。レーベンシュタイン距離とJaccard係数を
組み合わせたハイブリッドスコアリング方式を使用。

Author: Yuki0906
Version: 2.3.0 (Web UI / CLI Hybrid対応)
License: MIT
GitHub: https://github.com/Yuki-M0906/comprehensive-data-matcher
"""

import pandas as pd
import unicodedata
from rapidfuzz import fuzz
import logging
from pathlib import Path
from datetime import datetime
import sys
import re 

# --- 設定項目（CLIのデフォルト値として残します） ---
BASE_DIR = Path(__file__).parent.resolve()
YURAGI_FILE = BASE_DIR / "yuragi.xlsx"
CORRECT_FILE = BASE_DIR / "correct.xlsx"
RESULT_FILE = BASE_DIR / "result_combined_high_accuracy.xlsx"
LOG_FILE = BASE_DIR / "matcher_log_high_accuracy.txt"

YURAGI_SHEET_NAME = "Sheet1"
CORRECT_SHEET_NAME = "Sheet1"
RESULT_SHEET_NAME = "マッチング結果"
YURAGI_HEADER_ROW = 0
CORRECT_HEADER_ROW = 0
YURAGI_COL_NAME = '商品名'
CORRECT_COL_NAME = '商品名'

# デフォルトの重み
DEFAULT_WEIGHT_LEVENSHTEIN = 0.7
DEFAULT_WEIGHT_JACCARD = 0.3

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

# --- 正規化関数  ---
def normalize_cell_nfkc(value):
    if isinstance(value, str):
        return unicodedata.normalize('NFKC', value)
    return value

def normalize_for_comparison(s: str) -> str:
    if not isinstance(s, str):
        s = str(s)
    return s.upper()

def tokenize(s: str) -> set:
    tokens = re.split(r'[ 　/]', s)
    return set(filter(None, tokens))

# 【変更点】引数として weight_lev, weight_jac を受け取るようにしました
def calculate_hybrid_score(s1: str, s2: str, weight_lev: float, weight_jac: float) -> dict:
    lev_score = fuzz.ratio(s1, s2) / 100.0 
    
    tokens1 = tokenize(s1)
    tokens2 = tokenize(s2)
    
    if not tokens1 and not tokens2:
        jac_score = 1.0 
    elif not tokens1 or not tokens2:
        jac_score = 0.0 
    else:
        intersection = len(tokens1.intersection(tokens2))
        union = len(tokens1.union(tokens2))
        jac_score = intersection / union

    hybrid_score = (lev_score * weight_lev) + (jac_score * weight_jac)

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

# ==============================================================================
# ▼▼▼ 新設：コアロジック関数（Web版からもCLI版からも呼ばれる共通の頭脳） ▼▼▼
# ==============================================================================
def process_matching(df_yuragi_orig, df_correct, yuragi_col, correct_col, weight_lev, weight_jac):
    logger.info("高精度データマッチング処理を開始します。")
    logger.info(f"スコア重み設定: Levenshtein={weight_lev}, Jaccard={weight_jac}")

    # 前処理
    logger.info("--- ステップ2: yuragiデータのNFKC正規化 ---")
    df_yuragi_normalized = df_yuragi_orig.copy()
    # 最新のpandas対応としてapplymapを維持しつつエラー回避
    df_yuragi_normalized = df_yuragi_normalized.applymap(normalize_cell_nfkc)

    # 名寄せ処理
    logger.info("--- ステップ3: ハイブリッドスコアによる名寄せ処理 ---")
    yuragi_names_orig = df_yuragi_orig[yuragi_col].astype(str).tolist()
    yuragi_names_to_compare = df_yuragi_normalized[yuragi_col].astype(str).tolist()
    correct_names = df_correct[correct_col].astype(str).tolist()

    correct_cache = [
        {"original": s, "normalized": normalize_for_comparison(s)}
        for s in correct_names
    ]

    results_list = []
    total_count = len(yuragi_names_to_compare)
    
    for i, name_to_compare in enumerate(yuragi_names_to_compare):
        normalized_yuragi = normalize_for_comparison(name_to_compare)
        
        if (i + 1) % 100 == 0 or (i + 1) == total_count:
            logger.info(f"処理中: {i + 1}/{total_count}")

        best_match_original = None
        best_score_info = {"hybrid": -1.0} 

        for std_entry in correct_cache:
            std_normalized = std_entry["normalized"]
            
            # 【変更点】引数で受け取った重みを渡す
            current_score_info = calculate_hybrid_score(
                normalized_yuragi, std_normalized, weight_lev, weight_jac
            )
            
            if current_score_info["hybrid"] > best_score_info["hybrid"]:
                best_score_info = current_score_info
                best_match_original = std_entry["original"]
                if best_score_info["hybrid"] >= 1.0:
                    break
        
        results_list.append({
            '元のゆらぎ名': yuragi_names_orig[i],
            'マッチした正規名': best_match_original,
            'ハイブリッドスコア': best_score_info['hybrid'],
            'レーベンシュタイン類似度': best_score_info.get('levenshtein'),
            'Jaccard係数': best_score_info.get('jaccard'),
        })

    logger.info("マッチング処理が完了しました。")

    # フォーマット調整したデータフレームを作成して返す
    df_results = pd.DataFrame(results_list)
    df_results['ハイブリッドスコア'] = df_results['ハイブリッドスコア'].map('{:.3f}'.format)
    df_results['レーベンシュタイン類似度'] = df_results['レーベンシュタイン類似度'].map('{:.3f}'.format)
    df_results['Jaccard係数'] = df_results['Jaccard係数'].map('{:.3f}'.format)
    
    return df_results


# ==============================================================================
# ▼▼▼ ローカル実行(CLI)用のエントリーポイント ▼▼▼
# ==============================================================================
def main_cli():
    df_yuragi = load_excel(YURAGI_FILE, YURAGI_SHEET_NAME, YURAGI_HEADER_ROW)
    df_correct = load_excel(CORRECT_FILE, CORRECT_SHEET_NAME, CORRECT_HEADER_ROW)
    
    if df_yuragi is None or df_correct is None: 
        sys.exit(1)
        
    if YURAGI_COL_NAME not in df_yuragi.columns or CORRECT_COL_NAME not in df_correct.columns:
        logger.error("必須列が見つかりません。")
        sys.exit(1)

    # コアロジック呼び出し
    df_results = process_matching(
        df_yuragi_orig=df_yuragi, 
        df_correct=df_correct, 
        yuragi_col=YURAGI_COL_NAME, 
        correct_col=CORRECT_COL_NAME, 
        weight_lev=DEFAULT_WEIGHT_LEVENSHTEIN, 
        weight_jac=DEFAULT_WEIGHT_JACCARD
    )

    try:
        df_results.to_excel(RESULT_FILE, sheet_name=RESULT_SHEET_NAME, index=False, engine='openpyxl')
        logger.info(f"結果を正常に保存しました: {RESULT_FILE}")
        logger.info("すべての処理が正常に完了しました。")
    except Exception as e:
        logger.error("結果ファイルの保存中にエラーが発生しました。", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    start_time = datetime.now()
    try:
        main_cli()
    except Exception as e:
        logger.critical("予期せぬエラーによりスクリプトが異常終了しました。", exc_info=True)
    finally:
        end_time = datetime.now()
        logger.info(f"総実行時間: {end_time - start_time}")
