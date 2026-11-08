# Comprehensive Data Matcher

[※日本語の解説は下記にあります]

**[Try it Live (Web Demo)]**
[https://comprehensive-data-matcher.streamlit.app/](https://comprehensive-data-matcher.streamlit.app/)

A high-precision data deduplication and matching tool using a hybrid scoring method (Levenshtein distance and Jaccard similarity). Now supports both Web UI (Streamlit) and CLI.

## Overview

This tool automatically matches inconsistent data (e.g., product names with typos or different word orders) to a standardized master list.

With the recent update, we added a Web UI powered by Streamlit, making it a hybrid architecture that supports both an intuitive browser-based operation and local CLI batch processing.

### Key Features

* **Hybrid Scoring:** Combines Levenshtein similarity (character level) and Jaccard similarity (word level).
* **Web UI and CLI:** Use the browser-based Web UI, or the CLI for batch processing.
* **Try It Instantly:** The Web UI opens with built-in sample data (Japanese apparel names and English company names). You can also paste names or upload Excel / CSV files.
* **Review-Friendly Results:** Rows with a low score, or with a close second candidate, are flagged for review. You can see the top candidates, correct a match in the table, and download the result as Excel or CSV.
* **Flexible Weight Adjustment:** Adjust the weights with a slider (Web UI) or in the script (CLI).
* **Unicode Normalization:** NFKC normalization absorbs full-width/half-width differences on both sides.
* **Fast:** All pairs are scored at once with rapidfuzz on all CPU cores. 10,000 × 10,000 names take a few seconds on a typical PC.
* **Bilingual UI:** Japanese and English, with light and dark themes.

## Examples

### Input Data Example

**yuragi.xlsx** (Data with inconsistencies)
| Product Name (商品名) |
| --- |
| オーバーサイズ　コットンＴシャツ　白 |
| ビッグシルエット 綿Tシャツ ホワイト |
| デニム スリムフィット ブルー パンツ |
| パーカー フーディー 黒 スウェット |
| チノパンツ ストレッチ　ベージュ |

**correct.xlsx** (Standardized Master Data)
| Product Name (商品名) |
| --- |
| オーバーサイズ コットンTシャツ 白 |
| スリムフィット デニムパンツ ブルー |
| ケーブルニット セーター グレー |
| フード付き スウェットパーカー 黒 |
| ストレッチ チノパン ベージュ |

### Output Result Example

**result_combined_high_accuracy.xlsx**
| Original Name | Matched Standard Name | Hybrid Score | Levenshtein | Jaccard |
| --- | --- | --- | --- | --- |
| オーバーサイズ　コットンＴシャツ　白 | オーバーサイズ コットンTシャツ 白 | 1.000 | 1.000 | 1.000 |
| ビッグシルエット 綿Tシャツ ホワイト | ラウンドネック ロングTシャツ 白 | 0.272 | 0.389 | 0.000 |
| デニム スリムフィット ブルー パンツ | スリムフィット デニムパンツ ブルー | 0.536 | 0.595 | 0.400 |

The second row is a wrong match: the name is a paraphrase ("big silhouette" for "oversize", "綿" for "cotton"), which neither measure can recognize. Its low score is why the Web UI flags such rows for review.

## Installation

### Prerequisites

* Python 3.9+
* pip

### Steps

1. Clone the repository

```bash
git clone https://github.com/Yuki-M0906/comprehensive-data-matcher.git
cd comprehensive-data-matcher

```

2. Install required packages

```bash
pip install -r requirements.txt

```

## Usage

You can choose between two interfaces depending on your needs.

### Option 1: Web UI (Streamlit)

Match names in your browser without writing any code.

* **Use the Cloud Environment:**
Open the "Web Demo" URL above. No setup is needed.
* **Run Locally:**
Execute the following command to open the Web UI in your local browser:
```bash
streamlit run app.py

```

**How to use**

1. **Choose data** — try a sample, paste names (one per line), or upload two files (Excel `.xlsx`, or CSV in UTF-8 or Shift_JIS). For files, choose the sheet and the name column.
2. **Check the results** — the counts, a score histogram, and a table with the top candidates. A row is flagged *Review* when the top score is below the minimum, or when the second candidate is too close.
3. **Fix and download** — if needed, choose a different name in the "Matched name" column. Then download Excel (your original columns, the results, and a settings sheet) or CSV.

The public demo accepts up to 50,000 rows per file and 100 million comparisons. When running locally, you can change these limits with the environment variables `CDM_MAX_ROWS` and `CDM_MAX_PAIRS`. Data is processed in memory only and is not written to disk. For confidential data, run the app locally.


### Option 2: CLI (Command Line Interface)

Ideal for automating routine tasks or batch processing tens of thousands of records.

1. Place your `yuragi.xlsx` and `correct.xlsx` in the same directory as the script.
2. Run the script:

```bash
python comprehensive_data_matcher.py

```

3. The result will be generated as `result_combined_high_accuracy.xlsx`.

### Score Weight Adjustment

You can adjust the matching algorithm weights via the sliders on the Web UI, or by modifying the following settings in the script for the CLI version:

```python
# Levenshtein distance weight (character-level similarity)
DEFAULT_WEIGHT_LEVENSHTEIN = 0.7

# Jaccard similarity weight (word-level similarity)
DEFAULT_WEIGHT_JACCARD = 0.3

```

**Adjustment Guide:**

* To prioritize fixing typos -> Increase `DEFAULT_WEIGHT_LEVENSHTEIN`
* To prioritize word order differences -> Increase `DEFAULT_WEIGHT_JACCARD`

## Algorithm Details

The final score is calculated using the following formula:
`Final Score = (Levenshtein Similarity × 0.7) + (Jaccard Similarity × 0.3)`

For more details, please refer to `docs/ALGORITHM.md`.

## File Structure

```text
comprehensive-data-matcher/
├── app.py                         # Web UI (Streamlit)
├── app_text.py                    # Web UI text (Japanese / English)
├── comprehensive_data_matcher.py  # Core logic & CLI script
├── sample_data.py                 # Sample data (Web UI and generate_sample_data.py)
├── generate_sample_data.py        # Writes the sample Excel files to examples/
├── requirements.txt               # Required packages
├── pytest.ini                     # Test settings
├── .streamlit/config.toml         # Streamlit settings (theme, upload size limit)
├── tests/                         # Automated tests (pytest)
├── README.md                      # This file
├── LICENSE                        # License
├── .gitignore                     # Git ignore settings
├── examples/                      # Sample data
└── docs/                          # Additional documentation

```

## Notes

* Excel files must be in `.xlsx` format (`.xls` is not supported). The Web UI also accepts CSV.
* The CLI reads `yuragi.xlsx` and `correct.xlsx` (Excel only).

## Tests

```bash
pip install pytest
pytest

```

The tests check that the fast matching gives exactly the same scores as comparing pairs one by one, and run the Web UI without a browser.

## License

This project is licensed under the MIT License - see the `LICENSE` file for details.

## Author

Yuki-M0906

---

---

<a id="japanese"></a>

# Comprehensive Data Matcher (日本語)

高精度な名寄せ（データマッチング）を実現するPythonツール

**[ブラウザ版デモ環境]**
環境構築不要で、以下のURLからすぐにWebアプリ版をお試しいただけます。

[https://comprehensive-data-matcher.streamlit.app/](https://comprehensive-data-matcher.streamlit.app/)

## 概要

このツールは、表記ゆらぎのある商品名やデータを、標準的な名称に自動的にマッチングする名寄せツールです。レーベンシュタイン距離とJaccard係数を組み合わせたハイブリッドスコアリング方式により、高精度なマッチングを実現します。

今回、新たにWebブラウザ上で動作するUI（Streamlit）を追加し、ローカル環境でのバッチ処理（CLI）とWeb画面での直感的な操作の両方に対応するハイブリッド構成へとアップデートしました。

### 主な特徴

* **ハイブリッドスコアリング**: レーベンシュタイン類似度（文字の近さ）とJaccard係数（単語の一致）を組み合わせて点数を付けます
* **Web UI と CLI**: ブラウザーで使う Web 版と、まとめて処理するコマンドライン版（CLI）があります
* **すぐに試せる**: Web 版は見本データ（衣料品の商品名・英語の会社名）で開きます。名前の貼り付けや、Excel・CSV ファイルも使えます
* **確かめやすい結果**: 点数が低いもの、2 位と僅差のものを「要確認」にします。上位の候補を見比べ、表の中で対応づけを直し、Excel・CSV で保存できます
* **柔軟な重み調整**: Web 版はスライダーで、CLI 版はスクリプトの設定で、重みを変えられます
* **Unicode正規化**: NFKC正規化により、全角/半角の違いを両方の側で吸収します
* **高速処理**: 全部の組み合わせを rapidfuzz でまとめて計算します（CPU の全コアを使用）。1 万件 × 1 万件でも数秒です
* **日英の画面**: 日本語と英語を切り替えられます。明るい表示・暗い表示の両方に対応します

## 使用例

### 入力データの例

**yuragi.xlsx** (表記ゆらぎのあるデータ)
| 商品名 |
| --- |
| オーバーサイズ　コットンＴシャツ　白 |
| ビッグシルエット 綿Tシャツ ホワイト |
| デニム スリムフィット ブルー パンツ |
| パーカー フーディー 黒 スウェット |
| チノパンツ ストレッチ　ベージュ |

**correct.xlsx** (標準的な商品名)
| 商品名 |
| --- |
| オーバーサイズ コットンTシャツ 白 |
| スリムフィット デニムパンツ ブルー |
| ケーブルニット セーター グレー |
| フード付き スウェットパーカー 黒 |
| ストレッチ チノパン ベージュ |

### 出力結果の例

**result_combined_high_accuracy.xlsx**
| 元のゆらぎ名 | マッチした正規名 | ハイブリッドスコア | レーベンシュタイン類似度 | Jaccard係数 |
| --- | --- | --- | --- | --- |
| オーバーサイズ　コットンＴシャツ　白 | オーバーサイズ コットンTシャツ 白 | 1.000 | 1.000 | 1.000 |
| ビッグシルエット 綿Tシャツ ホワイト | ラウンドネック ロングTシャツ 白 | 0.272 | 0.389 | 0.000 |
| デニム スリムフィット ブルー パンツ | スリムフィット デニムパンツ ブルー | 0.536 | 0.595 | 0.400 |

2 行目は取り違えです（「ビッグシルエット」「綿」は「オーバーサイズ」「コットン」の言い換えで、どちらの物差しでも見抜けません）。このように点数の低い行を、Web 版は「要確認」にします。

## インストール

### 必要な環境

* Python 3.9以上
* pip

### インストール手順

1. リポジトリをクローン

```bash
git clone https://github.com/Yuki-M0906/comprehensive-data-matcher.git
cd comprehensive-data-matcher

```

2. 必要なパッケージをインストール

```bash
pip install -r requirements.txt

```

## 使い方

用途に合わせて、2つのインターフェースから選択できます。

### パターン1: Web UI版（Streamlit）を使う場合

プログラムを書かずに、ブラウザーで名寄せができます。クラウド環境またはローカル環境で利用可能です。

* **クラウド環境で使う:**
上記の「ブラウザ版デモ環境」URLにアクセスするだけで即座に利用可能です。
* **ローカル環境で立ち上げる:**
以下のコマンドを実行すると、ご自身のPCのブラウザが自動的に立ち上がります。
```bash
streamlit run app.py

```

**使い方**

1. **データを選ぶ** — 見本で試す、名前を貼り付ける（1 行に 1 件）、2 つのファイルを選ぶ（Excel の `.xlsx`、または UTF-8・Shift_JIS の CSV）のどれかです。ファイルのときは、シートと名前の列を選びます。
2. **結果を見る** — 件数、点数の分布の図、上位の候補つきの表が出ます。1 位の点数が下限より低いとき、または 2 位との差が小さいときは「要確認」になります。
3. **直して保存する** — 必要なら「対応づけた名前」を選び直します。Excel（元の列＋結果＋設定のシート）か CSV で保存できます。

公開デモでは、1 つのファイルにつき 50,000 行、比べる組み合わせは 1 億通りまでです。手元で動かすときは、環境変数 `CDM_MAX_ROWS`・`CDM_MAX_PAIRS` で上限を変えられます。データはメモリーの中だけで処理し、ファイルには保存しません。機密のデータは、手元の PC で動かしてください。


### パターン2: コマンドライン版（CLI）を使う場合

定常業務の自動化や、数万件規模の大量データのバッチ処理に適しています。

1. `yuragi.xlsx` と `correct.xlsx` をスクリプトと同じディレクトリに配置
2. スクリプトを実行

```bash
python comprehensive_data_matcher.py

```

3. 結果が `result_combined_high_accuracy.xlsx` に出力されます

### スコアの重み調整

マッチングアルゴリズムの重みは、Web UI画面上のスライダー、またはCLI版の場合はスクリプト内の以下の設定を変更することで調整できます。

```python
# レーベンシュタイン距離の重み (文字レベルの類似度)
DEFAULT_WEIGHT_LEVENSHTEIN = 0.7

# Jaccard係数の重み (単語レベルの類似度)
DEFAULT_WEIGHT_JACCARD = 0.3

```

**調整の指針:**

* 誤字脱字に強くしたい → `DEFAULT_WEIGHT_LEVENSHTEIN` を高く
* 語順の違いに強くしたい → `DEFAULT_WEIGHT_JACCARD` を高く

## アルゴリズムの説明

### ハイブリッドスコアリング方式

このツールは2つのアルゴリズムを組み合わせています:

1. **レーベンシュタイン類似度**
* 文字単位での編集距離を計算
* 誤字脱字、文字の挿入・削除に強い
* 例: "オーバーサイズ Tシャツ" ↔ "オーバーサイズTシャツ"


2. **Jaccard係数**
* 単語の集合の類似度を計算
* 語順の違いに強い
* 例: "スリムフィット デニムパンツ" ↔ "デニムパンツ スリムフィット"



最終スコア = (レーベンシュタイン類似度 × 0.7) + (Jaccard係数 × 0.3)

詳細は `docs/ALGORITHM.md` をご覧ください。

## ファイル構成

```text
comprehensive-data-matcher/
├── app.py                         # Web 版（Streamlit）
├── app_text.py                    # Web 版の画面の文言（日本語・英語）
├── comprehensive_data_matcher.py  # コアロジックおよびCLI実行用スクリプト
├── sample_data.py                 # 見本データ（Web 版と generate_sample_data.py で使う）
├── generate_sample_data.py        # 見本の Excel を examples/ に作る
├── requirements.txt               # 必要なパッケージ
├── pytest.ini                     # テストの設定
├── .streamlit/config.toml         # Streamlit の設定（色・アップロードの上限）
├── tests/                         # 自動テスト（pytest）
├── README.md                      # このファイル
├── LICENSE                        # ライセンス
├── .gitignore                     # Git除外設定
├── examples/                      # サンプルデータ
│   ├── yuragi_sample.xlsx         # generate_sample_data.py で作る
│   └── correct_sample.xlsx
└── docs/                          # 追加ドキュメント
    ├── ALGORITHM.md               # アルゴリズム詳細説明
    └── SETUP_GUIDE.md             # セットアップガイド

```

## カスタマイズ

### 設定項目

CLI版で実行する場合、スクリプト内で以下の設定を変更できます:

```python
# 入力ファイル名
YURAGI_FILE = BASE_DIR / "yuragi.xlsx"
CORRECT_FILE = BASE_DIR / "correct.xlsx"

# 出力ファイル名
RESULT_FILE = BASE_DIR / "result_combined_high_accuracy.xlsx"

# シート名と列名
YURAGI_SHEET_NAME = "Sheet1"
YURAGI_COL_NAME = '商品名'

```

## ログ出力

CLI版の実行中のログは以下の2箇所に出力されます:

* コンソール: リアルタイムで進捗を確認
* `matcher_log_high_accuracy.txt`: 詳細なログファイル

## 注意事項

* Excelファイルは `.xlsx` 形式である必要があります（`.xls` は非対応）。Web 版は CSV も読めます
* CLI 版が読むのは `yuragi.xlsx` と `correct.xlsx`（Excel だけ）です
* CLI版で入力ファイルが存在しない場合、エラーメッセージが表示されます

## テスト

```bash
pip install pytest
pytest

```

まとめて計算した結果が 1 組ずつ比べた結果と同じ点数になること、Web 版がブラウザーなしで動くことを確かめます。

## 貢献

バグ報告、機能リクエスト、プルリクエストを歓迎します！
詳細は `CONTRIBUTING.md` をご覧ください。

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細は `LICENSE` ファイルをご覧ください。

## 作成者

Yuki-M0906

## お問い合わせ

質問や提案がありましたら、Issueを作成してください。

---

このプロジェクトが役に立ったら、スターをつけていただけると嬉しいです！

---

