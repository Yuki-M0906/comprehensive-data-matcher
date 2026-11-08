# -*- coding: utf-8 -*-
"""処理の本体（comprehensive_data_matcher.py）のテスト。"""

import logging
import math
import random

import numpy as np
import pandas as pd
import pytest

import comprehensive_data_matcher as cdm
from sample_data import APPAREL_MASTER, APPAREL_VARIANTS, SAMPLES


def reference(queries, choices, weight_lev, weight_jac):
    """1 組ずつ calculate_hybrid_score() で比べる素朴な方法（v2.3.0 の処理の流れ）。"""
    cache = [(j, cdm.prepare_name(c)) for j, c in enumerate(choices) if not cdm.is_blank(c)]
    out = []
    for q in queries:
        if cdm.is_blank(q) or not cache:
            out.append((-1, None))
            continue
        nq = cdm.prepare_name(q)
        best, info = None, {"hybrid": -1.0}
        for j, s in cache:
            cur = cdm.calculate_hybrid_score(nq, s, weight_lev, weight_jac)
            if cur["hybrid"] > info["hybrid"]:
                best, info = j, cur
                if info["hybrid"] >= 1.0:
                    break
        out.append((best, info))
    return out


WORDS = ["白", "黒", "Tシャツ", "パンツ", "デニム", "ニット", "ｺｯﾄﾝ", "コットン", "ＶＮＥＣＫ", "vneck",
         "/", "　", " ", "グレー", "ブルー", "a", "A", "1"]


def random_name(rng):
    r = rng.random()
    if r < 0.05:
        return None
    if r < 0.08:
        return float("nan")
    if r < 0.10:
        return "   "
    if r < 0.12:
        return 123
    return "".join(rng.choice(WORDS) for _ in range(rng.randint(1, 6)))


@pytest.mark.parametrize("seed", range(40))
def test_same_result_as_one_by_one(seed):
    rng = random.Random(seed)
    queries = [random_name(rng) for _ in range(rng.randint(0, 40))]
    choices = [random_name(rng) for _ in range(rng.randint(0, 30))]
    wl = rng.choice([0.7, 0.0, 1.0, 0.35, rng.random()])
    wj = rng.choice([1 - wl, 0.3, 0.5, rng.random()])  # 合計が 1 を超える重みも試す
    top_k = rng.choice([1, 2, 3, 5])

    res = cdm.match_names(queries, choices, wl, wj, top_k=top_k)
    for i, (j, info) in enumerate(reference(queries, choices, wl, wj)):
        if j == -1:
            assert res["index"][i, 0] == -1
            assert np.isnan(res["hybrid"][i, 0])
            continue
        assert res["index"][i, 0] == j
        assert res["hybrid"][i, 0] == pytest.approx(info["hybrid"], abs=1e-12)
        assert res["levenshtein"][i, 0] == pytest.approx(info["levenshtein"], abs=1e-12)
        assert res["jaccard"][i, 0] == pytest.approx(info["jaccard"], abs=1e-12)
        # 2 位以下：重複なし・スコアの高い順
        idx = res["index"][i]
        picked = idx[idx >= 0]
        assert len(set(picked.tolist())) == len(picked)
        rest = res["hybrid"][i, 1:len(picked)]
        assert all(rest[t] >= rest[t + 1] for t in range(len(rest) - 1))


def test_top_candidates_are_the_best_scores():
    rng = random.Random(99)
    queries = [random_name(rng) or "x" for _ in range(30)]
    choices = [random_name(rng) or "y" for _ in range(25)]
    res = cdm.match_names(queries, choices, 0.7, 0.3, top_k=3)
    for i, q in enumerate(queries):
        if cdm.is_blank(q):
            continue
        nq = cdm.prepare_name(q)
        scores = sorted(
            (cdm.calculate_hybrid_score(nq, cdm.prepare_name(c), 0.7, 0.3)["hybrid"]
             for c in choices if not cdm.is_blank(c)), reverse=True)
        assert res["hybrid"][i, :3] == pytest.approx(scores[:3], abs=1e-12)


def test_sample_scores_match_v230():
    """v2.3.0 で見本データを実行したときの値（表記ゆれを含むマスターではないので、正規化の修正の影響なし）。"""
    df = cdm.process_matching(pd.DataFrame({"商品名": APPAREL_VARIANTS}), pd.DataFrame({"商品名": APPAREL_MASTER}),
                              "商品名", "商品名", 0.7, 0.3)
    expected = {
        0: ("オーバーサイズ コットンTシャツ 白", "1.000", "1.000", "1.000"),
        1: ("ラウンドネック ロングTシャツ 白", "0.272", "0.389", "0.000"),
        5: ("スリムフィット デニムパンツ ブルー", "0.536", "0.595", "0.400"),
        24: ("クルーネック スウェット 杢グレー", "0.120", "0.171", "0.000"),
        33: ("ラウンドネック ロングTシャツ 白", "0.477", "0.467", "0.500"),
    }
    assert list(df.columns) == ["元のゆらぎ名", "マッチした正規名", "ハイブリッドスコア", "レーベンシュタイン類似度", "Jaccard係数"]
    for row, (name, h, lev, jac) in expected.items():
        assert tuple(df.iloc[row, 1:]) == (name, h, lev, jac)


def test_blank_rows_and_blank_master_entries():
    queries = ["白 Tシャツ", None, "", "   ", float("nan"), "黒 パンツ"]
    choices = ["", None, "白 Tシャツ", "黒 パンツ"]
    res = cdm.match_names(queries, choices, 0.7, 0.3, top_k=3)
    assert res["index"][:, 0].tolist() == [2, -1, -1, -1, -1, 3]
    # 空欄のマスターは候補にならない（2 位までしか無い）
    assert res["index"][0].tolist() == [2, 3, -1]
    df = cdm.process_matching(pd.DataFrame({"n": queries}), pd.DataFrame({"n": choices}), "n", "n", 0.7, 0.3)
    assert df["元のゆらぎ名"].tolist()[1:5] == ["", "", "", ""]
    assert all(pd.isna(v) for v in df["マッチした正規名"].tolist()[1:5])
    assert df["ハイブリッドスコア"].tolist()[1:5] == [""] * 4


def test_master_is_normalized_too():
    """マスターも NFKC でそろえる（v2.4.0 で修正）。全角のマスターでも完全一致は 1.000。"""
    res = cdm.match_names(["Vネック ニット"], ["ほかの名前", "Ｖネック ニット"], 0.7, 0.3)
    assert res["index"][0, 0] == 1
    assert res["hybrid"][0, 0] == pytest.approx(1.0)


def test_duplicates_are_computed_once_and_spread_back():
    calls = []
    queries = ["白 Tシャツ"] * 5 + ["黒 パンツ"] * 3
    res = cdm.match_names(queries, ["白 Tシャツ", "黒 パンツ"], 0.7, 0.3,
                          progress=lambda done, total: calls.append((done, total)))
    assert res["index"][:, 0].tolist() == [0] * 5 + [1] * 3
    assert calls[-1] == (2, 2)  # 同じ名前は 1 回だけ数える


def test_empty_inputs():
    res = cdm.match_names([], ["a"], 0.7, 0.3, top_k=2)
    assert res["index"].shape == (0, 2)
    res = cdm.match_names(["a", "b"], [], 0.7, 0.3, top_k=2)
    assert (res["index"] == -1).all()
    assert np.isnan(res["hybrid"]).all()


def test_large_input_is_chunked(monkeypatch):
    monkeypatch.setattr(cdm, "_CHUNK_CELLS", 50)  # 区切って計算しても同じ結果になる
    rng = random.Random(5)
    queries = [random_name(rng) for _ in range(60)]
    choices = [random_name(rng) for _ in range(20)]
    chunked = cdm.match_names(queries, choices, 0.6, 0.4, top_k=3)
    monkeypatch.setattr(cdm, "_CHUNK_CELLS", 10_000_000)
    whole = cdm.match_names(queries, choices, 0.6, 0.4, top_k=3)
    assert (chunked["index"] == whole["index"]).all()
    assert np.allclose(chunked["hybrid"], whole["hybrid"], equal_nan=True)


def test_company_sample_matches_as_intended():
    """見本の会社名は、全部が意図したマスターに対応づく（並びの順に 4・3・3・2・2・2・2・2・2・2 件ずつ）。"""
    s = SAMPLES["companies"]
    expected = [0] * 4 + [1] * 3 + [2] * 3 + [3] * 2 + [4] * 2 + [5] * 2 + [6] * 2 + [7] * 2 + [8] * 2 + [9] * 2
    res = cdm.match_names(s["variants"], s["master"], 0.7, 0.3)
    assert res["index"][:, 0].tolist() == expected


def test_importing_does_not_add_a_log_file_handler():
    """Web 版などで読み込んだだけではログファイルを作らない（CLI の setup_cli_logging() だけが作る）。"""
    assert not any(isinstance(h, logging.FileHandler) for h in cdm.logger.handlers)


def test_is_blank():
    assert cdm.is_blank(None) and cdm.is_blank("") and cdm.is_blank(" 　") and cdm.is_blank(math.nan)
    assert cdm.is_blank(pd.NA)
    assert not cdm.is_blank(0) and not cdm.is_blank("a") and not cdm.is_blank([1, 2])
