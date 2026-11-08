# -*- coding: utf-8 -*-
"""Web 版（app.py）の画面の文言。日本語（ja）と英語（en）。"""

TEXT = {
    "ja": {
        "page_title": "名寄せツール | Comprehensive Data Matcher",
        "title": "名寄せツール",
        "lead": (
            "表記のゆれた名前を、正しい名前の一覧（マスター）に自動で対応づけます。"
            "**文字の近さ**と**単語の一致**を組み合わせて点数を付け、迷うものは「要確認」として人の目に回します。"
        ),
        "tags": ":blue-badge[文字の近さ（レーベンシュタイン）] :blue-badge[単語の一致（Jaccard）] "
                ":gray-badge[全角・半角をそろえる（NFKC）] :gray-badge[Excel / CSV]",
        "tab_match": "名寄せ",
        "tab_how": "スコアの仕組み",

        # サイドバー
        "settings": "設定",
        "w_label": "文字の近さを重く見る割合",
        "w_help": "大きくするほど誤字・脱字に強く、小さくするほど語順の違いに強くなります。",
        "w_caption": "文字の近さ **{lev}%**（誤字・脱字に強い）／単語の一致 **{jac}%**（語順の違いに強い）",
        "th_label": "OK とする点数の下限",
        "th_help": "1 位の点数がこれより低いと「要確認」にします。",
        "margin_label": "2 位との差の下限",
        "margin_help": "1 位と 2 位の点数の差がこれより小さいと、取り違えのおそれがあるので「要確認」にします。",
        "k_label": "表に出す候補の数",
        "reset": "既定に戻す",
        "about": "このデモについて",
        "about_text": (
            "データは、この画面を開いている間だけサーバーのメモリーで処理し、ファイルには保存しません。"
            "機密のデータは、手元の PC で動かしてください（README の「使い方」）。"
        ),
        "link_repo": "GitHub のリポジトリ",
        "link_algo": "アルゴリズムの説明",

        # 1. データ
        "step1": "1. データを選ぶ",
        "src_sample": "見本で試す",
        "src_paste": "貼り付ける",
        "src_file": "ファイル（Excel / CSV）",
        "sample_apparel": "衣料品の商品名（日本語）",
        "sample_companies": "取引先の会社名（英語）",
        "variants_label": "ゆらぎのある名前",
        "master_label": "マスター（正しい名前）",
        "count": "{n:,} 件",
        "run_sample": "この見本で名寄せする",
        "paste_hint": "1 行に 1 件ずつ入れます。Excel の列をコピーして貼り付けても使えます。",
        "paste_variants": "ゆらぎのある名前",
        "paste_master": "マスター（正しい名前）",
        "run_paste": "貼り付けた内容で名寄せする",
        "paste_empty": "両方に 1 件以上入れてください。",
        "file_hint": "1 行目を列の名前として読みます。CSV は UTF-8 と Shift_JIS（cp932）に対応します。",
        "file_variants": "ゆらぎのあるデータ",
        "file_master": "マスター",
        "sheet": "シート",
        "column": "名前の列",
        "rows": "{n:,} 行",
        "preview": "先頭の 5 行を見る",
        "run_file": "ファイルで名寄せする",
        "file_wait": "両方のファイルを選ぶと実行できます。",
        "read_error": "「{name}」を読めませんでした：{err}",
        "no_columns": "「{name}」に列がありません。",
        "too_many_rows": (
            "「{name}」は {n:,} 行あります。このデモでは 1 つにつき {max:,} 行までです。"
            "手元の PC で動かせば上限を変えられます（環境変数 CDM_MAX_ROWS）。"
        ),
        "too_many_pairs": (
            "組み合わせが {pairs:,} 通り（同じ名前を除いた数）になり、このデモの上限 {max:,} を超えます。"
            "手元の PC で動かせば上限を変えられます（環境変数 CDM_MAX_PAIRS）。"
        ),

        # 2. 結果
        "step2": "2. 結果",
        "source_sample": "見本「{name}」",
        "source_paste": "貼り付けた内容",
        "source_file": "ファイル「{a}」と「{b}」",
        "result_of": "対象：{src}",
        "auto_update": "左の設定を変えると、結果もすぐに変わります。",
        "progress": "計算しています… {done:,} / {total:,}",
        "m_rows": "件数",
        "m_ok": "OK",
        "m_ok_value": "{n:,}（{p:.0%}）",
        "m_review": "要確認",
        "m_time": "処理時間",
        "m_time_value": "{s:.2f} 秒",
        "m_pairs": "{pairs:,} 通りを比較",
        "chart_x": "1 位の点数",
        "chart_y": "件数",
        "chart_ok": "OK",
        "chart_review": "要確認",
        "chart_rule": "下限 {th:.2f}",
        "filter": "表示",
        "f_all": "すべて",
        "f_review": "要確認だけ",
        "f_ok": "OK だけ",
        "f_edited": "手で直したもの",
        "search": "名前で絞り込む",
        "edit_hint": "「対応づけた名前」は、セルを選んで選び直せます。直した内容はダウンロードに入ります。",
        "no_rows": "当てはまる行はありません。",
        "col_no": "No.",
        "col_name": "名前",
        "col_input": "入力の名前",
        "col_match": "対応づけた名前",
        "col_score": "点数",
        "col_status": "判定",
        "col_lev": "文字の近さ",
        "col_jac": "単語の一致",
        "col_cand": "{n} 位の候補",
        "col_cand_score": "{n} 位の点数",
        "st_ok": "✅ OK",
        "st_low": "⚠️ 要確認：点数が低い",
        "st_close": "⚠️ 要確認：2 位と僅差",
        "st_edited": "✏️ 手で修正",
        "st_blank": "— 空欄",
        "st_none": "— 候補なし",

        # 3. 保存
        "step3": "3. 保存する",
        "dl_prepare": "保存用のファイルを作る",
        "dl_excel": "Excel で保存",
        "dl_csv": "CSV で保存",
        "dl_note": (
            "Excel には、元の列のあとに結果の列を足したシートと、設定のシートが入ります。"
            "CSV は Excel でそのまま開ける形（UTF-8・BOM 付き）です。"
        ),
        "sheet_results": "結果",
        "sheet_settings": "設定",
        "set_item": "項目",
        "set_value": "値",
        "set_version": "ツールの版",
        "set_created": "作成日時",
        "set_source": "対象",
        "set_w_lev": "文字の近さの重み",
        "set_w_jac": "単語の一致の重み",
        "set_th": "OK とする点数の下限",
        "set_margin": "2 位との差の下限",
        "set_rows": "件数",
        "set_ok": "OK",
        "set_review": "要確認",
        "set_edited": "手で修正",

        # スコアの仕組み
        "how_intro": """
1 つの名前とマスターの全部の名前を比べて、いちばん点数の高いものを対応づけます。
点数は 2 つの物差しを組み合わせて決めます。

| 物差し | 測るもの | 得意なゆれ |
|---|---|---|
| **文字の近さ**（レーベンシュタイン類似度） | 一方をもう一方に直すのに要る、文字の足し引きの少なさ | 誤字・脱字、スペースの有無 |
| **単語の一致**（Jaccard 係数） | 空白と「/」で区切った単語のうち、共通する単語の割合 | 語順の入れ替わり |
""",
        "how_formula": "点数は次の式で計算します（重み w は左の設定の「文字の近さを重く見る割合」）。",
        "how_normalize": """
**比べる前にそろえること**

- NFKC 正規化：全角の英数字・記号を半角に、半角カナを全角にそろえます（例：`Ｔシャツ` → `Tシャツ`）。
- 英字を大文字にそろえます（例：`navy` → `NAVY`）。

**判定**

- 1 位の点数が「OK とする点数の下限」より低い → **要確認：点数が低い**
- 1 位と 2 位の点数の差が「2 位との差の下限」より小さい → **要確認：2 位と僅差**
- どちらでもない → **OK**
""",
        "try_title": "2 つの名前で試す",
        "try_a": "名前 A",
        "try_b": "名前 B",
        "try_norm": "比較に使う形",
        "try_tokens": "単語",
        "try_total": "点数",
        "try_calc": "{wl:.2f} × {lev:.3f} ＋ {wj:.2f} × {jac:.3f} ＝ **{h:.3f}**",
        "try_none": "（なし）",

        "footer": "Comprehensive Data Matcher v{ver} ・ MIT License ・ [GitHub]({url})",
    },

    "en": {
        "page_title": "Name Matcher | Comprehensive Data Matcher",
        "title": "Name Matcher",
        "lead": (
            "Match messy, inconsistent names to a clean master list. "
            "Each pair is scored by combining **character similarity** and **word overlap**, "
            "and uncertain matches are flagged for human review."
        ),
        "tags": ":blue-badge[Character similarity (Levenshtein)] :blue-badge[Word overlap (Jaccard)] "
                ":gray-badge[Width normalization (NFKC)] :gray-badge[Excel / CSV]",
        "tab_match": "Match",
        "tab_how": "How scoring works",

        "settings": "Settings",
        "w_label": "Weight on character similarity",
        "w_help": "Higher values tolerate typos better; lower values tolerate different word order better.",
        "w_caption": "Character similarity **{lev}%** (typos) / word overlap **{jac}%** (word order)",
        "th_label": "Minimum score for OK",
        "th_help": "If the top score is below this value, the row is flagged for review.",
        "margin_label": "Minimum lead over 2nd place",
        "margin_help": "If the top two scores are closer than this, the match may be wrong, so the row is flagged for review.",
        "k_label": "Candidates shown in the table",
        "reset": "Reset to defaults",
        "about": "About this demo",
        "about_text": (
            "Data is processed in server memory only while this page is open and is never written to disk. "
            "For confidential data, run the app on your own computer (see Usage in the README)."
        ),
        "link_repo": "GitHub repository",
        "link_algo": "Algorithm details (Japanese)",

        "step1": "1. Choose data",
        "src_sample": "Try a sample",
        "src_paste": "Paste",
        "src_file": "Files (Excel / CSV)",
        "sample_apparel": "Apparel products (Japanese)",
        "sample_companies": "Company names (English)",
        "variants_label": "Messy names",
        "master_label": "Master list (clean names)",
        "count": "{n:,} names",
        "run_sample": "Match this sample",
        "paste_hint": "One name per line. You can paste a column copied from Excel.",
        "paste_variants": "Messy names",
        "paste_master": "Master list (clean names)",
        "run_paste": "Match pasted names",
        "paste_empty": "Enter at least one name in each box.",
        "file_hint": "The first row is read as column names. CSV files may be UTF-8 or Shift_JIS (cp932).",
        "file_variants": "Messy data",
        "file_master": "Master list",
        "sheet": "Sheet",
        "column": "Name column",
        "rows": "{n:,} rows",
        "preview": "Show the first 5 rows",
        "run_file": "Match files",
        "file_wait": "Choose both files to run.",
        "read_error": "Could not read “{name}”: {err}",
        "no_columns": "“{name}” has no columns.",
        "too_many_rows": (
            "“{name}” has {n:,} rows. This demo accepts up to {max:,} rows per file. "
            "Run it on your own computer to raise the limit (environment variable CDM_MAX_ROWS)."
        ),
        "too_many_pairs": (
            "This would compare {pairs:,} pairs (after removing duplicates), above this demo's limit of {max:,}. "
            "Run it on your own computer to raise the limit (environment variable CDM_MAX_PAIRS)."
        ),

        "step2": "2. Results",
        "source_sample": "sample “{name}”",
        "source_paste": "pasted names",
        "source_file": "files “{a}” and “{b}”",
        "result_of": "Source: {src}",
        "auto_update": "Changing the settings on the left updates the results immediately.",
        "progress": "Matching… {done:,} / {total:,}",
        "m_rows": "Rows",
        "m_ok": "OK",
        "m_ok_value": "{n:,} ({p:.0%})",
        "m_review": "Review",
        "m_time": "Time",
        "m_time_value": "{s:.2f} s",
        "m_pairs": "{pairs:,} pairs compared",
        "chart_x": "Top score",
        "chart_y": "Rows",
        "chart_ok": "OK",
        "chart_review": "Review",
        "chart_rule": "min {th:.2f}",
        "filter": "Show",
        "f_all": "All",
        "f_review": "Review only",
        "f_ok": "OK only",
        "f_edited": "Edited",
        "search": "Filter by name",
        "edit_hint": "Select a cell in “Matched name” to choose a different name. Your edits are included in the downloads.",
        "no_rows": "No matching rows.",
        "col_no": "No.",
        "col_name": "Name",
        "col_input": "Input name",
        "col_match": "Matched name",
        "col_score": "Score",
        "col_status": "Status",
        "col_lev": "Character similarity",
        "col_jac": "Word overlap",
        "col_cand": "Candidate {n}",
        "col_cand_score": "Score {n}",
        "st_ok": "✅ OK",
        "st_low": "⚠️ Review: low score",
        "st_close": "⚠️ Review: close call",
        "st_edited": "✏️ Edited",
        "st_blank": "— Blank",
        "st_none": "— No candidate",

        "step3": "3. Download",
        "dl_prepare": "Prepare files for download",
        "dl_excel": "Download Excel",
        "dl_csv": "Download CSV",
        "dl_note": (
            "The Excel file has your original columns followed by the result columns, plus a settings sheet. "
            "The CSV opens directly in Excel (UTF-8 with BOM)."
        ),
        "sheet_results": "Results",
        "sheet_settings": "Settings",
        "set_item": "Item",
        "set_value": "Value",
        "set_version": "Tool version",
        "set_created": "Created",
        "set_source": "Source",
        "set_w_lev": "Weight: character similarity",
        "set_w_jac": "Weight: word overlap",
        "set_th": "Minimum score for OK",
        "set_margin": "Minimum lead over 2nd place",
        "set_rows": "Rows",
        "set_ok": "OK",
        "set_review": "Review",
        "set_edited": "Edited",

        "how_intro": """
Each input name is compared with every name in the master list, and the highest-scoring one is chosen.
The score combines two measures.

| Measure | What it measures | Good at |
|---|---|---|
| **Character similarity** (Levenshtein ratio) | How few character insertions and deletions turn one name into the other | Typos, missing characters, spacing |
| **Word overlap** (Jaccard index) | The share of words in common, splitting on spaces and "/" | Different word order |
""",
        "how_formula": "The score is calculated as follows (w is “Weight on character similarity” in the settings).",
        "how_normalize": """
**Normalization before comparing**

- NFKC normalization: full-width letters, digits and symbols become half-width, and half-width katakana becomes full-width (e.g. `Ｔシャツ` → `Tシャツ`).
- Letters are converted to upper case (e.g. `navy` → `NAVY`).

**Status**

- Top score below “Minimum score for OK” → **Review: low score**
- Top two scores closer than “Minimum lead over 2nd place” → **Review: close call**
- Otherwise → **OK**
""",
        "try_title": "Try two names",
        "try_a": "Name A",
        "try_b": "Name B",
        "try_norm": "Normalized",
        "try_tokens": "Words",
        "try_total": "Score",
        "try_calc": "{wl:.2f} × {lev:.3f} + {wj:.2f} × {jac:.3f} = **{h:.3f}**",
        "try_none": "(none)",

        "footer": "Comprehensive Data Matcher v{ver} · MIT License · [GitHub]({url})",
    },
}
