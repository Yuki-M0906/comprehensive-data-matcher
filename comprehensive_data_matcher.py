# -*- coding: utf-8 -*-
"""
Comprehensive Data Matcher

高精度な名寄せツール。レーベンシュタイン距離とJaccard係数を
組み合わせたハイブリッドスコアリング方式を使用。

Author: Yuki-M0906
Version: 2.4.0
License: MIT
GitHub: https://github.com/Yuki-M0906/comprehensive-data-matcher
"""

import logging
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from rapidfuzz import fuzz, process

__version__ = "2.4.0"

# --- 設定項目（CLIのデフォルト値） ---
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

# 一度に計算する組み合わせの数の目安（メモリーの使用量を抑えるため、ゆらぎ側を区切って計算する）
_CHUNK_CELLS = 2_000_000

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


def setup_cli_logging(log_file=LOG_FILE):
    """CLI で実行するときだけ、コンソールとログファイルへの出力を設定する。

    読み込んだだけ（Web 版など）ではログファイルを作らない。
    """
    log_formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    logger.handlers.clear()
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(log_formatter)
    logger.addHandler(stream_handler)
    try:
        file_handler = logging.FileHandler(log_file, encoding='utf-8', mode='w')
        file_handler.setFormatter(log_formatter)
        logger.addHandler(file_handler)
    except OSError as e:
        logger.warning(f"ログファイルを作れませんでした（コンソールにだけ出します）: {e}")


# --- 正規化関数 ---
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


def is_blank(value) -> bool:
    """空欄（None・NaN・空白だけの文字列）なら True。"""
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def prepare_name(value) -> str:
    """比較に使う形にする（NFKC 正規化して大文字にする）。"""
    return normalize_for_comparison(normalize_cell_nfkc(value))


def calculate_hybrid_score(s1: str, s2: str, weight_lev: float, weight_jac: float) -> dict:
    """2 つの文字列（比較用に整えた後のもの）のスコアを 1 組だけ計算する。"""
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


def _token_index(choice_norm):
    """Jaccard 用の索引。単語 → その単語を含むマスターの位置の配列、と、マスターごとの単語の数。"""
    vocab = {}
    postings = []
    sizes = np.empty(len(choice_norm), dtype=np.float64)
    for j, s in enumerate(choice_norm):
        toks = tokenize(s)
        sizes[j] = len(toks)
        for t in toks:
            tid = vocab.get(t)
            if tid is None:
                tid = len(postings)
                vocab[t] = tid
                postings.append([])
            postings[tid].append(j)
    return vocab, [np.asarray(p, dtype=np.int64) for p in postings], sizes


def _jaccard_matrix(part, vocab, postings, sizes):
    """ゆらぎの名前（part）× マスターの Jaccard 係数。calculate_hybrid_score() と同じ値。"""
    m = len(sizes)
    empty = sizes == 0
    jac = np.empty((len(part), m), dtype=np.float64)
    for a, s in enumerate(part):
        toks = tokenize(s)
        if not toks:
            # 両方とも単語なし → 1.0、片方だけ → 0.0
            jac[a] = np.where(empty, 1.0, 0.0)
            continue
        hits = [postings[vocab[t]] for t in toks if t in vocab]
        inter = (np.bincount(np.concatenate(hits), minlength=m).astype(np.float64)
                 if hits else np.zeros(m, dtype=np.float64))
        with np.errstate(divide='ignore', invalid='ignore'):
            jac[a] = np.where(empty, 0.0, inter / (len(toks) + sizes - inter))
    return jac


def _pick_top(hybrid, k):
    """行ごとに上位 k 件の列の位置を選ぶ。"""
    r, m = hybrid.shape
    rows = np.arange(r)
    # 1 位：1 件ずつ比べる処理と同じ選び方（最初に 1.0 以上になったもの、無ければ最初の最大値）
    reach = hybrid >= 1.0
    best = np.where(reach.any(axis=1), reach.argmax(axis=1), hybrid.argmax(axis=1))
    picks = np.empty((r, k), dtype=np.int64)
    picks[:, 0] = best
    if k > 1:
        rest = hybrid.copy()
        rest[rows, best] = -np.inf
        if k - 1 < m - 1:
            cand = np.argpartition(-rest, k - 2, axis=1)[:, :k - 1]
        else:
            cand = np.tile(np.arange(m), (r, 1))
            cand = cand[cand != best[:, None]].reshape(r, m - 1)
        cand_scores = rest[rows[:, None], cand]
        # スコアの高い順。同じスコアならマスターの前のほうを先に
        order = np.lexsort((cand, -cand_scores), axis=1)
        picks[:, 1:] = np.take_along_axis(cand, order, axis=1)
    return picks


def match_names(queries, choices, weight_lev=DEFAULT_WEIGHT_LEVENSHTEIN, weight_jac=DEFAULT_WEIGHT_JACCARD,
                top_k=1, progress=None):
    """ゆらぎの名前ごとに、マスターの中からスコアの高い候補を探す。

    1 組ずつ calculate_hybrid_score() を呼ぶのと同じスコア・同じ選び方になるように、
    まとめて（行列で）計算する。

    Args:
        queries: ゆらぎの名前の並び（元の値。空欄は「候補なし」になる）
        choices: マスターの名前の並び（元の値。空欄は候補から外す）
        weight_lev / weight_jac: スコアの重み
        top_k: 候補をいくつまで返すか（1 以上）
        progress: 進み具合を受け取る関数 progress(済んだ数, 全体の数)。省略できる

    Returns:
        dict:
            "index": (ゆらぎの件数, top_k) の int 配列。choices の中の位置。候補が無いところは -1
            "hybrid" / "levenshtein" / "jaccard": 同じ形の float 配列。候補が無いところは NaN
    """
    queries = list(queries)
    choices = list(choices)
    top_k = max(1, int(top_k))
    n = len(queries)
    out = {
        "index": np.full((n, top_k), -1, dtype=np.int64),
        "hybrid": np.full((n, top_k), np.nan),
        "levenshtein": np.full((n, top_k), np.nan),
        "jaccard": np.full((n, top_k), np.nan),
    }

    # マスター：空欄を除き、比較用に整える
    choice_pos = np.asarray([j for j, v in enumerate(choices) if not is_blank(v)], dtype=np.int64)
    choice_norm = [prepare_name(choices[j]) for j in choice_pos]
    m = len(choice_norm)
    k = min(top_k, m)

    # ゆらぎ：空欄を除き、同じ名前は 1 回だけ計算する
    rows_of = {}
    for i, v in enumerate(queries):
        if not is_blank(v):
            rows_of.setdefault(prepare_name(v), []).append(i)
    uniq = list(rows_of)
    total = len(uniq)

    if m == 0 or total == 0:
        if progress:
            progress(total, total)
        return out

    vocab, postings, sizes = _token_index(choice_norm)
    chunk = max(1, _CHUNK_CELLS // m)
    for start in range(0, total, chunk):
        part = uniq[start:start + chunk]
        lev = process.cdist(part, choice_norm, scorer=fuzz.ratio, dtype=np.float64, workers=-1) / 100.0
        jac = _jaccard_matrix(part, vocab, postings, sizes)
        hybrid = (lev * weight_lev) + (jac * weight_jac)
        picks = _pick_top(hybrid, k)

        rows = np.arange(len(part))[:, None]
        values = {"index": choice_pos[picks], "hybrid": hybrid[rows, picks],
                  "levenshtein": lev[rows, picks], "jaccard": jac[rows, picks]}
        for a, key in enumerate(part):
            for name, arr in values.items():
                out[name][rows_of[key], :k] = arr[a]

        if progress:
            progress(min(start + chunk, total), total)

    return out


def load_excel(file_path, sheet_name, header):
    if not file_path.exists():
        logger.error(f"ファイルが見つかりません: {file_path}")
        return None
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header, engine='openpyxl')
        logger.info(f"ファイル読み込み完了: {file_path} ({len(df)}行)")
        return df
    except Exception:
        logger.error(f"ファイルの読み込み中にエラーが発生しました: {file_path}", exc_info=True)
        return None


# ==============================================================================
# コアロジック（Web 版・CLI 版の共通）
# ==============================================================================
def process_matching(df_yuragi_orig, df_correct, yuragi_col, correct_col, weight_lev, weight_jac):
    """ゆらぎの各行に、マスターの中で最もスコアの高い名前を対応づける。

    戻り値の列（v2.3.0 と同じ）：元のゆらぎ名・マッチした正規名・ハイブリッドスコア・
    レーベンシュタイン類似度・Jaccard係数（スコアは小数第 3 位までの文字列）。
    空欄の行は、正規名とスコアを空にする。
    """
    logger.info("高精度データマッチング処理を開始します。")
    logger.info(f"スコア重み設定: Levenshtein={weight_lev}, Jaccard={weight_jac}")

    queries = df_yuragi_orig[yuragi_col].tolist()
    choices = df_correct[correct_col].tolist()
    logger.info(f"ゆらぎ {len(queries)} 件 × マスター {len(choices)} 件")

    def _progress(done, total):
        logger.info(f"処理中: {done}/{total}（重複を除いた件数）")

    res = match_names(queries, choices, weight_lev, weight_jac, top_k=1, progress=_progress)
    logger.info("マッチング処理が完了しました。")

    def _fmt(x):
        return '' if np.isnan(x) else '{:.3f}'.format(x)

    idx = res["index"][:, 0]
    df_results = pd.DataFrame({
        '元のゆらぎ名': ['' if is_blank(v) else str(v) for v in queries],
        'マッチした正規名': [str(choices[j]) if j >= 0 else None for j in idx],
        'ハイブリッドスコア': [_fmt(x) for x in res["hybrid"][:, 0]],
        'レーベンシュタイン類似度': [_fmt(x) for x in res["levenshtein"][:, 0]],
        'Jaccard係数': [_fmt(x) for x in res["jaccard"][:, 0]],
    })
    return df_results


# ==============================================================================
# ローカル実行（CLI）用のエントリーポイント
# ==============================================================================
def main_cli():
    df_yuragi = load_excel(YURAGI_FILE, YURAGI_SHEET_NAME, YURAGI_HEADER_ROW)
    df_correct = load_excel(CORRECT_FILE, CORRECT_SHEET_NAME, CORRECT_HEADER_ROW)

    if df_yuragi is None or df_correct is None:
        sys.exit(1)

    if YURAGI_COL_NAME not in df_yuragi.columns or CORRECT_COL_NAME not in df_correct.columns:
        logger.error("必須列が見つかりません。")
        sys.exit(1)

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
    except Exception:
        logger.error("結果ファイルの保存中にエラーが発生しました。", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    setup_cli_logging()
    start_time = datetime.now()
    try:
        main_cli()
    except Exception:
        logger.critical("予期せぬエラーによりスクリプトが異常終了しました。", exc_info=True)
    finally:
        end_time = datetime.now()
        logger.info(f"総実行時間: {end_time - start_time}")
