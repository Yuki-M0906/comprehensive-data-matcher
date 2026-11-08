# -*- coding: utf-8 -*-
"""
Comprehensive Data Matcher - Web 版（Streamlit）

    streamlit run app.py

処理の本体は comprehensive_data_matcher.py（CLI と共通）。画面の文言は app_text.py（日本語・英語）。
"""

import hashlib
import inspect
import io
import os
import re
import time
from datetime import datetime

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

import comprehensive_data_matcher as cdm
from app_text import TEXT
from sample_data import SAMPLES

REPO_URL = "https://github.com/Yuki-M0906/comprehensive-data-matcher"
ALGO_URL = REPO_URL + "/blob/main/docs/ALGORITHM.md"

# 公開デモを守るための上限（手元で動かすときは環境変数で変えられる）
MAX_ROWS = int(os.environ.get("CDM_MAX_ROWS", "50000"))
MAX_PAIRS = int(os.environ.get("CDM_MAX_PAIRS", "100000000"))

MAX_K = 5                      # 計算しておく候補の数（表に出す数は設定で 1〜5）
MAX_SELECT_OPTIONS = 2000      # マスターがこれより多いと、対応づけた名前を一覧から選ばせず文字で入れてもらう
AUTO_DOWNLOAD_ROWS = 5000      # これより多いと、保存用のファイルはボタンを押してから作る
SHOW_PROGRESS_PAIRS = 2_000_000
CACHE_ENTRIES = 6

SETTINGS = {"w_lev": 70, "threshold": 0.50, "margin": 0.05, "top_k": 3}
NAME_HINTS = ("商品名", "品名", "名称", "名前", "会社名", "取引先名", "氏名",
              "name", "Name", "NAME", "product", "Product", "company", "Company")
LANGS = {"ja": "日本語", "en": "English"}
COLOR_OK = "#2563EB"
COLOR_REVIEW = "#F59E0B"
STATUS_KEYS = ("ok", "low", "close", "edited", "blank", "none")

ss = st.session_state


# ---------------------------------------------------------------------------
# 状態
# ---------------------------------------------------------------------------
def k(name):
    """言語ごとのウィジェットのキー。言語を切り替えるときに値を引き継ぐ（_on_lang）。"""
    return f"{name}__{ss.lang}"


def init_state():
    if "lang" not in ss:
        ss.lang = "en" if st.query_params.get("lang") == "en" else "ja"
    ss.setdefault("lang_pick", LANGS[ss.lang])
    ss.setdefault("active", "sample")
    ss.setdefault("overrides", {})
    ss.setdefault("results", {})
    ss.setdefault("files", {})
    ss.setdefault("downloads", {})
    for name, value in SETTINGS.items():
        ss.setdefault(k(name), value)
    ss.setdefault(k("sample"), "apparel")
    ss.setdefault(k("paste_v"), "\n".join(SAMPLES["companies"]["variants"][:6]))
    ss.setdefault(k("paste_m"), "\n".join(SAMPLES["companies"]["master"][:4]))
    ss.setdefault(k("filter"), "all")
    ss.setdefault(k("search"), "")
    ss.setdefault(k("try_a"), "デニム スリムフィット ブルー パンツ")
    ss.setdefault(k("try_b"), "スリムフィット デニムパンツ ブルー")


def _on_lang():
    picked = ss.get("lang_pick")
    if picked is None:  # 選択を外したときは元に戻す
        ss.lang_pick = LANGS[ss.lang]
        return
    new = "en" if picked == LANGS["en"] else "ja"
    old = ss.lang
    if new == old:
        return
    suffix = f"__{old}"
    for key in list(ss.keys()):
        if isinstance(key, str) and key.endswith(suffix):
            ss[key[: -len(old)] + new] = ss[key]
    ss.lang = new
    st.query_params["lang"] = new


def _reset_settings():
    for name, value in SETTINGS.items():
        ss[k(name)] = value


def _activate(source):
    ss.active = source


def _sha(*parts):
    h = hashlib.sha1()
    for p in parts:
        h.update(repr(p).encode("utf-8", "surrogatepass"))
        h.update(b"\x1f")
    return h.hexdigest()


# ---------------------------------------------------------------------------
# 入力
# ---------------------------------------------------------------------------
def read_upload(uploaded, sheet=None):
    """アップロードされたファイルを表にする。(df, シート名の一覧, エラー) を返す。同じファイルは読み直さない。"""
    key = (uploaded.file_id, sheet)
    if key in ss.files:
        return ss.files[key]
    data = uploaded.getvalue()
    name = uploaded.name.lower()
    result = (None, [], None)
    try:
        if name.endswith(".csv"):
            df, err = None, None
            for enc in ("utf-8-sig", "cp932"):
                try:
                    df = pd.read_csv(io.BytesIO(data), encoding=enc, dtype=str, sep=None, engine="python")
                    break
                except UnicodeDecodeError as e:
                    err = e
            if df is None:
                raise err
            result = (df, [], None)
        else:
            book = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
            sheets = list(book.sheet_names)
            use = sheet if sheet in sheets else sheets[0]
            result = (book.parse(use), sheets, None)
    except Exception as e:  # 読めないファイルは、理由を画面に出す
        result = (None, [], f"{type(e).__name__}: {e}")
    if len(ss.files) > 8:
        ss.files.clear()
    ss.files[key] = result
    return result


def guess_column(df):
    cols = list(df.columns)
    for hint in NAME_HINTS:
        for c in cols:
            if str(c).strip() == hint:
                return c
    for c in cols:
        if df[c].dtype == object or str(df[c].dtype).startswith("str"):
            return c
    return cols[0]


def file_side(T, side, title):
    """ファイル 1 つ分の入力欄。(df, 列, ファイル名, 問題の文) を返す。"""
    st.markdown(f"**{title}**")
    uploaded = st.file_uploader("Excel / CSV", type=["xlsx", "csv"], key=f"upload_{side}",
                                label_visibility="collapsed")
    if uploaded is None:
        return None, None, None, None
    df, sheets, err = read_upload(uploaded)
    if err:
        return None, None, uploaded.name, T["read_error"].format(name=uploaded.name, err=err)
    if len(sheets) > 1:
        sheet_key = k(f"sheet_{side}")
        if ss.get(sheet_key) not in sheets:
            ss[sheet_key] = sheets[0]
        sheet = st.selectbox(T["sheet"], sheets, key=sheet_key)
        df, _, err = read_upload(uploaded, sheet)
        if err:
            return None, None, uploaded.name, T["read_error"].format(name=uploaded.name, err=err)
    if df is None or len(df.columns) == 0:
        return None, None, uploaded.name, T["no_columns"].format(name=uploaded.name)
    col_key = k(f"col_{side}_{uploaded.file_id}")
    cols = list(df.columns)
    if ss.get(col_key) not in cols:
        ss[col_key] = guess_column(df)
    col = st.selectbox(T["column"], cols, key=col_key, format_func=str)
    st.caption(T["rows"].format(n=len(df)))
    with st.expander(T["preview"]):
        st.dataframe(df.head(5), hide_index=True, width="stretch")
    if len(df) > MAX_ROWS:
        return None, None, uploaded.name, T["too_many_rows"].format(name=uploaded.name, n=len(df), max=MAX_ROWS)
    return df, col, uploaded.name, None


def render_inputs(T):
    """データを選ぶ欄。今の対象（ss.active）の入力を返す。"""
    st.subheader(T["step1"])
    tab_sample, tab_paste, tab_file = st.tabs([T["src_sample"], T["src_paste"], T["src_file"]])
    inputs = {}

    with tab_sample:
        picked = st.segmented_control(
            T["src_sample"], ["apparel", "companies"], key=k("sample"),
            format_func=lambda x: T[f"sample_{x}"], label_visibility="collapsed")
        name = picked or "apparel"
        s = SAMPLES[name]
        c1, c2 = st.columns(2)
        for col, label, values in ((c1, T["variants_label"], s["variants"]), (c2, T["master_label"], s["master"])):
            with col:
                st.markdown(f"**{label}** ・ {T['count'].format(n=len(values))}")
                st.dataframe(pd.DataFrame({T["col_name"]: values}), hide_index=True, height=200, width="stretch")
        st.button(T["run_sample"], icon=":material/play_arrow:", on_click=_activate, args=("sample",),
                  type="primary" if ss.active != "sample" else "secondary")
        inputs["sample"] = {
            "variants": list(s["variants"]), "master": list(s["master"]), "base_df": None,
            "input_col": s["column"], "label": T["source_sample"].format(name=T[f"sample_{name}"]),
            "sig": _sha("sample", name),
        }

    with tab_paste:
        st.caption(T["paste_hint"])
        c1, c2 = st.columns(2)
        text_v = c1.text_area(T["paste_variants"], key=k("paste_v"), height=230)
        text_m = c2.text_area(T["paste_master"], key=k("paste_m"), height=230)
        variants = [line.strip() for line in text_v.splitlines() if line.strip()]
        master = [line.strip() for line in text_m.splitlines() if line.strip()]
        c1.caption(T["count"].format(n=len(variants)))
        c2.caption(T["count"].format(n=len(master)))
        ready = bool(variants and master)
        if not ready:
            st.info(T["paste_empty"])
        if len(variants) > MAX_ROWS or len(master) > MAX_ROWS:
            st.error(T["too_many_rows"].format(name=T["src_paste"], n=max(len(variants), len(master)), max=MAX_ROWS))
            ready = False
        st.button(T["run_paste"], icon=":material/play_arrow:", on_click=_activate, args=("paste",),
                  disabled=not ready, type="primary" if ss.active != "paste" else "secondary")
        if ready:
            inputs["paste"] = {
                "variants": variants, "master": master, "base_df": None, "input_col": None,
                "label": T["source_paste"], "sig": _sha("paste", variants, master),
            }

    with tab_file:
        st.caption(T["file_hint"])
        c1, c2 = st.columns(2)
        with c1:
            df_v, col_v, name_v, err_v = file_side(T, "v", T["file_variants"])
        with c2:
            df_m, col_m, name_m, err_m = file_side(T, "m", T["file_master"])
        for err in (err_v, err_m):
            if err:
                st.error(err)
        ready = df_v is not None and df_m is not None
        if not ready and not (err_v or err_m):
            st.info(T["file_wait"])
        st.button(T["run_file"], icon=":material/play_arrow:", on_click=_activate, args=("file",),
                  disabled=not ready, type="primary" if ss.active != "file" else "secondary")
        if ready:
            variants = df_v[col_v].tolist()
            master = df_m[col_m].tolist()
            inputs["file"] = {
                "variants": variants, "master": master, "base_df": df_v, "input_col": None,
                "label": T["source_file"].format(a=name_v, b=name_m),
                "sig": _sha("file", variants, master),
            }

    return inputs.get(ss.active)


# ---------------------------------------------------------------------------
# 計算と表
# ---------------------------------------------------------------------------
def count_pairs(inp):
    queries = {cdm.prepare_name(v) for v in inp["variants"] if not cdm.is_blank(v)}
    choices = sum(1 for v in inp["master"] if not cdm.is_blank(v))
    return len(queries) * choices


def compute(inp, wl, wj, T):
    """名寄せの計算。同じ入力・同じ重みなら、この画面の中で前の結果を使う。"""
    key = (inp["sig"], wl, wj)
    if key in ss.results:
        return ss.results[key]
    pairs = count_pairs(inp)
    if pairs > MAX_PAIRS:
        return {"error": T["too_many_pairs"].format(pairs=pairs, max=MAX_PAIRS)}
    bar = st.progress(0.0, text=T["progress"].format(done=0, total=0)) if pairs >= SHOW_PROGRESS_PAIRS else None

    def _progress(done, total):
        if bar is not None and total:
            bar.progress(done / total, text=T["progress"].format(done=done, total=total))

    start = time.perf_counter()
    res = cdm.match_names(inp["variants"], inp["master"], wl, wj, top_k=MAX_K, progress=_progress)
    res["elapsed"] = time.perf_counter() - start
    res["pairs"] = pairs
    res["blank"] = np.array([cdm.is_blank(v) for v in inp["variants"]], dtype=bool)
    if bar is not None:
        bar.empty()
    while len(ss.results) >= CACHE_ENTRIES:
        ss.results.pop(next(iter(ss.results)))
    ss.results[key] = res
    return res


def build_table(inp, res, top_k, th, margin, overrides, T):
    """画面の表を作る。(表, 1 位の名前の一覧, 行ごとの判定の配列の dict) を返す。"""
    variants, master = inp["variants"], inp["master"]
    n = len(variants)
    idx, hy = res["index"], res["hybrid"]
    blank = res["blank"]
    best = idx[:, 0]
    none = (best < 0) & ~blank
    valid = ~blank & ~none
    top = np.round(hy[:, 0], 3)
    second = np.round(hy[:, 1], 3)
    with np.errstate(invalid="ignore"):
        low = valid & (top < th)
        close = valid & ~low & ~np.isnan(second) & (np.round(top - second, 3) < margin)
    review = low | close
    ok = valid & ~review

    master_str = ["" if cdm.is_blank(v) else str(v) for v in master]

    def name_at(j):
        return master_str[j] if j >= 0 else None

    best_names = [name_at(j) for j in best]
    matched = list(best_names)
    status = np.select([blank, none, low, close], ["blank", "none", "low", "close"], "ok").astype(object)
    edited = np.zeros(n, dtype=bool)
    for row, value in overrides.items():
        if 0 <= row < n:
            matched[row] = value
            status[row] = "edited"
            edited[row] = True

    labels = {key: T[f"st_{key}"] for key in STATUS_KEYS}
    data = {
        T["col_no"]: np.arange(1, n + 1),
        T["col_input"]: ["" if b else str(v) for v, b in zip(variants, blank)],
        T["col_match"]: matched,
        T["col_status"]: [labels[s] for s in status],
        T["col_score"]: top,
    }
    for c in range(1, top_k):
        data[T["col_cand"].format(n=c + 1)] = [name_at(j) for j in idx[:, c]]
        data[T["col_cand_score"].format(n=c + 1)] = np.round(hy[:, c], 3)
    data[T["col_lev"]] = np.round(res["levenshtein"][:, 0], 3)
    data[T["col_jac"]] = np.round(res["jaccard"][:, 0], 3)
    table = pd.DataFrame(data)
    flags = {"review": review, "ok": ok, "edited": edited, "valid": valid, "status": status}
    return table, best_names, flags


def score_chart(scores, groups, th, T):
    df = pd.DataFrame({"score": scores, "group": groups})
    bars = alt.Chart(df).mark_bar(cornerRadiusTopLeft=2, cornerRadiusTopRight=2).encode(
        x=alt.X("score:Q", bin=alt.Bin(extent=[0, 1], step=0.05), title=T["chart_x"],
                scale=alt.Scale(domain=[0, 1])),
        y=alt.Y("count():Q", title=T["chart_y"], stack=True),
        color=alt.Color("group:N", title=None,
                        scale=alt.Scale(domain=[T["chart_ok"], T["chart_review"]], range=[COLOR_OK, COLOR_REVIEW]),
                        legend=alt.Legend(orient="top")),
        tooltip=[alt.Tooltip("group:N", title=T["col_status"]), alt.Tooltip("count():Q", title=T["chart_y"])],
    )
    rule_df = pd.DataFrame({"x": [th], "label": [T["chart_rule"].format(th=th)]})
    rule = alt.Chart(rule_df).mark_rule(strokeDash=[4, 3], color="#6B7280").encode(x="x:Q")
    text = alt.Chart(rule_df).mark_text(align="left", dx=4, dy=8, color="#6B7280").encode(
        x="x:Q", y=alt.value(0), text="label:N")
    return (bars + rule + text).properties(height=200)


_ALTAIR_HAS_WIDTH = "width" in inspect.signature(st.altair_chart).parameters


def show_chart(chart):
    if _ALTAIR_HAS_WIDTH:
        st.altair_chart(chart, width="stretch")
    else:
        st.altair_chart(chart, use_container_width=True)


def make_downloads(inp, table, flags, T, meta):
    """保存用の Excel と CSV のバイト列を作る。判定は記号を外した文にする。"""
    if inp["base_df"] is not None:
        out = inp["base_df"].reset_index(drop=True).copy()
    else:
        out = pd.DataFrame({inp["input_col"] or T["col_input"]: inp["variants"]})
    plain = {key: T[f"st_{key}"].split(" ", 1)[-1] for key in STATUS_KEYS}
    result = table.drop(columns=[T["col_no"], T["col_input"]]).copy()
    result[T["col_status"]] = [plain[s] for s in flags["status"]]
    for col in result.columns:
        name = col
        while name in out.columns:
            name = f"{name}_2"
        out[name] = result[col].values

    csv_bytes = out.to_csv(index=False).encode("utf-8-sig")

    settings = pd.DataFrame({T["set_item"]: list(meta.keys()), T["set_value"]: [str(v) for v in meta.values()]})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name=T["sheet_results"], index=False)
        settings.to_excel(writer, sheet_name=T["sheet_settings"], index=False)
        ws = writer.sheets[T["sheet_results"]]
        ws.freeze_panes = "A2"
        for i, col in enumerate(out.columns, start=1):
            sample = [str(col)] + [str(v) for v in out[col].head(200).tolist()]
            width = min(60, max(8, max(len(s) for s in sample) * 1.6))
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
        ws2 = writer.sheets[T["sheet_settings"]]
        ws2.column_dimensions["A"].width = 30
        ws2.column_dimensions["B"].width = 50
    return buf.getvalue(), csv_bytes


def render_summary(table, flags, res, th, T):
    """件数などの数字と、点数の分布の図。数字の dict を返す。"""
    n = len(table)
    counts = {
        "rows": n,
        "ok": int((flags["ok"] & ~flags["edited"]).sum()),
        "review": int((flags["review"] & ~flags["edited"]).sum()),
        "edited": int(flags["edited"].sum()),
    }
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(T["m_rows"], f"{n:,}", border=True)
    m2.metric(T["m_ok"], T["m_ok_value"].format(n=counts["ok"], p=counts["ok"] / n) if n else "0", border=True)
    m3.metric(T["m_review"], f"{counts['review']:,}", border=True)
    m4.metric(T["m_time"], T["m_time_value"].format(s=res["elapsed"]), border=True,
              help=T["m_pairs"].format(pairs=res["pairs"]))
    if counts["edited"]:
        st.caption(f"{T['st_edited']}：{counts['edited']:,}")

    valid = flags["valid"]
    if valid.any():
        groups = np.where(flags["review"][valid], T["chart_review"], T["chart_ok"])
        show_chart(score_chart(table[T["col_score"]].to_numpy()[valid], groups, th, T))
    return counts


def filter_rows(table, flags, T):
    """絞り込みの欄。(絞り込んだ表, 絞り込みの種類, 検索の文字) を返す。"""
    c1, c2 = st.columns([3, 2], vertical_alignment="bottom")
    with c1:
        flt = st.segmented_control(
            T["filter"], ["all", "review", "ok", "edited"], key=k("filter"),
            format_func=lambda x: T[f"f_{x}"]) or "all"
    with c2:
        query = st.text_input(T["search"], key=k("search"), placeholder="🔍")
    mask = flags[flt].copy() if flt in ("review", "ok", "edited") else np.ones(len(table), dtype=bool)
    q = query.strip()
    if q:
        hit = (table[T["col_input"]].str.contains(q, case=False, regex=False, na=False)
               | table[T["col_match"]].fillna("").astype(str).str.contains(q, case=False, regex=False))
        mask &= hit.to_numpy()
    return table[mask], flt, query


def editor_config(inp, top_k, T):
    options = list(dict.fromkeys(str(x) for x in inp["master"] if not cdm.is_blank(x)))
    if len(options) <= MAX_SELECT_OPTIONS:
        match_col = st.column_config.SelectboxColumn(T["col_match"], options=options, required=False, width="medium")
    else:
        match_col = st.column_config.TextColumn(T["col_match"], width="medium")
    config = {
        T["col_no"]: st.column_config.NumberColumn(T["col_no"], format="%d", width="small"),
        T["col_input"]: st.column_config.TextColumn(T["col_input"], width="medium"),
        T["col_match"]: match_col,
        T["col_score"]: st.column_config.ProgressColumn(T["col_score"], min_value=0.0, max_value=1.0,
                                                        format="%.3f", width="small"),
        T["col_status"]: st.column_config.TextColumn(T["col_status"], width="medium"),
        T["col_lev"]: st.column_config.NumberColumn(T["col_lev"], format="%.3f", width="small"),
        T["col_jac"]: st.column_config.NumberColumn(T["col_jac"], format="%.3f", width="small"),
    }
    for c in range(1, top_k):
        config[T["col_cand_score"].format(n=c + 1)] = st.column_config.NumberColumn(format="%.3f", width="small")
    return config


def render_editor(inp, view, best_names, overrides, editor_key, top_k, T):
    """結果の表。「対応づけた名前」を直したら overrides に覚えて、画面を描き直す。"""
    if view.empty:
        st.info(T["no_rows"])
        return
    st.caption(T["edit_hint"])
    edited = st.data_editor(
        view, key=editor_key, hide_index=True, column_config=editor_config(inp, top_k, T), num_rows="fixed",
        disabled=[c for c in view.columns if c != T["col_match"]],
        height=min(38 + 35 * len(view), 560), width="stretch")

    before = view[T["col_match"]].fillna("").astype(str)
    after = edited[T["col_match"]].fillna("").astype(str)
    changed = before.index[(before != after).to_numpy()]
    if len(changed) == 0:
        return
    for row in changed:
        value = after.loc[row] or None
        if value == best_names[row]:
            overrides.pop(int(row), None)
        else:
            overrides[int(row)] = value
    ss.edit_ver = ss.get("edit_ver", 0) + 1
    st.rerun()


def render_downloads(inp, table, flags, counts, cfg, overrides, T):
    st.subheader(T["step3"])
    meta = {
        T["set_version"]: cdm.__version__,
        T["set_created"]: datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        T["set_source"]: inp["label"],
        T["set_w_lev"]: f"{cfg['wl']:.2f}",
        T["set_w_jac"]: f"{cfg['wj']:.2f}",
        T["set_th"]: f"{cfg['th']:.2f}",
        T["set_margin"]: f"{cfg['margin']:.2f}",
        T["set_rows"]: counts["rows"],
        T["set_ok"]: counts["ok"],
        T["set_review"]: counts["review"],
        T["set_edited"]: counts["edited"],
    }
    dl_key = _sha(inp["sig"], sorted(cfg.items()), ss.lang, sorted(overrides.items()))
    files = ss.downloads.get(dl_key)
    if files is None and (counts["rows"] <= AUTO_DOWNLOAD_ROWS or ss.get("dl_requested") == dl_key):
        files = make_downloads(inp, table, flags, T, meta)
        ss.downloads = {dl_key: files}
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    c1, c2, c3 = st.columns([1, 1, 3], vertical_alignment="center")
    if files is None:
        c1.button(T["dl_prepare"], icon=":material/build:", on_click=lambda: ss.update(dl_requested=dl_key))
    else:
        c1.download_button(T["dl_excel"], files[0], file_name=f"matching_result_{stamp}.xlsx", icon=":material/download:",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        c2.download_button(T["dl_csv"], files[1], file_name=f"matching_result_{stamp}.csv", icon=":material/download:",
                           mime="text/csv")
    c3.caption(T["dl_note"])


def render_results(inp, T):
    st.subheader(T["step2"])
    if inp is None:
        st.info(T["file_wait"] if ss.active == "file" else T["paste_empty"])
        return
    st.caption(T["result_of"].format(src=inp["label"]) + "　" + T["auto_update"])

    w = ss[k("w_lev")]
    cfg = {"wl": w / 100, "wj": (100 - w) / 100, "th": ss[k("threshold")], "margin": ss[k("margin")],
           "top_k": ss[k("top_k")]}
    res = compute(inp, cfg["wl"], cfg["wj"], T)
    if "error" in res:
        st.error(res["error"])
        return

    overrides = ss.overrides.setdefault(inp["sig"], {})
    table, best_names, flags = build_table(inp, res, cfg["top_k"], cfg["th"], cfg["margin"], overrides, T)
    counts = render_summary(table, flags, res, cfg["th"], T)
    view, flt, query = filter_rows(table, flags, T)
    # 表の中身が変わる設定（重み・下限・絞り込みなど）が変わったら、表を作り直す
    editor_key = "editor_" + _sha(inp["sig"], sorted(cfg.items()), flt, query, ss.lang,
                                  ss.get("edit_ver", 0) if flt == "edited" else 0)[:16]
    render_editor(inp, view, best_names, overrides, editor_key, cfg["top_k"], T)
    render_downloads(inp, table, flags, counts, cfg, overrides, T)


# ---------------------------------------------------------------------------
# スコアの仕組み
# ---------------------------------------------------------------------------
def ordered_tokens(s):
    return list(dict.fromkeys(t for t in re.split(r'[ 　/]', s) if t))


def render_how(T):
    st.markdown(T["how_intro"])
    st.markdown(T["how_formula"])
    st.latex(r"\mathrm{score} = w \times \mathrm{Levenshtein} + (1 - w) \times \mathrm{Jaccard}")
    st.markdown(T["how_normalize"])

    st.subheader(T["try_title"])
    w = ss[k("w_lev")]
    wl, wj = w / 100, (100 - w) / 100
    c1, c2 = st.columns(2)
    a = c1.text_input(T["try_a"], key=k("try_a"))
    b = c2.text_input(T["try_b"], key=k("try_b"))
    na, nb = cdm.prepare_name(a), cdm.prepare_name(b)
    for col, norm in ((c1, na), (c2, nb)):
        with col:
            st.caption(T["try_norm"])
            st.code(norm or " ", language=None)
            st.caption(T["try_tokens"])
            st.code("  |  ".join(ordered_tokens(norm)) or T["try_none"], language=None)
    sc = cdm.calculate_hybrid_score(na, nb, wl, wj)
    m1, m2, m3 = st.columns(3)
    m1.metric(T["col_lev"], f"{sc['levenshtein']:.3f}", border=True)
    m2.metric(T["col_jac"], f"{sc['jaccard']:.3f}", border=True)
    m3.metric(T["try_total"], f"{sc['hybrid']:.3f}", border=True)
    st.markdown(T["try_calc"].format(wl=wl, wj=wj, lev=sc["levenshtein"], jac=sc["jaccard"], h=sc["hybrid"]))


# ---------------------------------------------------------------------------
# 画面
# ---------------------------------------------------------------------------
def render_sidebar(T):
    with st.sidebar:
        st.segmented_control("Language", list(LANGS.values()), key="lang_pick", on_change=_on_lang,
                             label_visibility="collapsed")
        st.header(T["settings"])
        w = st.slider(T["w_label"], 0, 100, step=5, format="%d%%", key=k("w_lev"), help=T["w_help"])
        st.caption(T["w_caption"].format(lev=w, jac=100 - w))
        st.slider(T["th_label"], 0.0, 1.0, step=0.05, format="%.2f", key=k("threshold"), help=T["th_help"])
        st.slider(T["margin_label"], 0.0, 0.30, step=0.01, format="%.2f", key=k("margin"), help=T["margin_help"])
        st.select_slider(T["k_label"], options=list(range(1, MAX_K + 1)), key=k("top_k"))
        st.button(T["reset"], icon=":material/restart_alt:", on_click=_reset_settings)
        st.divider()
        st.subheader(T["about"])
        st.caption(T["about_text"])
        st.markdown(f"- [{T['link_repo']}]({REPO_URL})\n- [{T['link_algo']}]({ALGO_URL})")


def main():
    init_state()
    T = TEXT[ss.lang]
    st.set_page_config(page_title=T["page_title"], page_icon=":material/join_inner:", layout="wide")
    render_sidebar(T)

    st.title(T["title"])
    st.markdown(T["lead"])
    st.markdown(T["tags"])

    tab_match, tab_how = st.tabs([T["tab_match"], T["tab_how"]])
    with tab_match:
        inp = render_inputs(T)
        st.divider()
        render_results(inp, T)
    with tab_how:
        render_how(T)

    st.divider()
    st.caption(T["footer"].format(ver=cdm.__version__, url=REPO_URL))


main()
