# Comprehensive Data Matcher

[※日本語の解説は下記にあります]

**[Try it Live (Web Demo)]**
[https://comprehensive-data-matcher.streamlit.app/](https://comprehensive-data-matcher.streamlit.app/)

A high-precision data deduplication and matching tool using a hybrid scoring method (Levenshtein distance and Jaccard similarity). Now supports both Web UI (Streamlit) and CLI.

## Overview

This tool automatically matches inconsistent data (e.g., product names with typos or different word orders) to a standardized master list.

With the recent update, we added a Web UI powered by Streamlit, making it a hybrid architecture that supports both an intuitive browser-based operation and local CLI batch processing.

### Key Features

* **Hybrid Scoring:** A unique scoring method combining Levenshtein distance and Jaccard similarity for high accuracy.
* **Dual Interface:** Supports both an intuitive Web UI and a CLI suitable for automation.
* **Flexible Weight Adjustment:** Adjust the algorithm weights according to your data (dynamically adjustable via sliders on the Web UI).
* **Detailed Logging:** Verify matching results one by one in real-time.
* **Unicode Normalization:** Automatically absorbs full-width/half-width differences using NFKC normalization.
* **High-Speed Processing:** Efficient string matching utilizing the rapidfuzz library.

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
| オーバーサイズ　コットンＴシャツ　白 | オーバーサイズ コットンTシャツ 白 | 0.967 | 0.950 | 1.000 |
| ビッグシルエット 綿Tシャツ ホワイト | オーバーサイズ コットンTシャツ 白 | 0.725 | 0.607 | 1.000 |
| デニム スリムフィット ブルー パンツ | スリムフィット デニムパンツ ブルー | 0.883 | 0.794 | 1.000 |

## Installation

### Prerequisites

* Python 3.8+
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

Allows you to intuitively upload Excel files from your browser and execute the matching process.

* **Use the Cloud Environment:**
Simply access the "Web Demo" URL above to use the app immediately without any setup.
* **Run Locally:**
Execute the following command to automatically open the Web UI in your local browser:
```bash
streamlit run app.py

```



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
WEIGHT_LEVENSHTEIN = 0.7

# Jaccard similarity weight (word-level similarity)
WEIGHT_JACCARD = 0.3

```

**Adjustment Guide:**

* To prioritize fixing typos -> Increase `WEIGHT_LEVENSHTEIN`
* To prioritize word order differences -> Increase `WEIGHT_JACCARD`

## Algorithm Details

The final score is calculated using the following formula:
`Final Score = (Levenshtein Similarity × 0.7) + (Jaccard Similarity × 0.3)`

For more details, please refer to `docs/ALGORITHM.md`.

## File Structure

```text
comprehensive-data-matcher/
├── app.py                         # Web UI script (New)
├── comprehensive_data_matcher.py  # Core logic & CLI script
├── requirements.txt               # Required packages
├── README.md                      # This file
├── LICENSE                        # License
├── .gitignore                     # Git ignore settings
├── examples/                      # Sample data
└── docs/                          # Additional documentation

```

## Notes

* Processing large datasets (10,000+ rows) via CLI may take some time.
* Excel files must be in `.xlsx` format (`.xls` is not supported).

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

* **ハイブリッドスコアリング**: レーベンシュタイン類似度とJaccard係数を組み合わせた独自のスコアリング方式
* **デュアルインターフェース対応**: 直感的なWeb UIと、自動化に適したコマンドライン（CLI）の両方をサポート
* **柔軟な重み調整**: 用途に応じてアルゴリズムの重みを調整可能（Web画面からはスライダーで動的に調整可能）
* **詳細なログ出力**: マッチング結果を1件ずつリアルタイムで確認可能
* **Unicode正規化**: NFKC正規化により全角/半角の違いを吸収
* **高速処理**: rapidfuzzライブラリによる効率的な文字列マッチング

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
| オーバーサイズ　コットンＴシャツ　白 | オーバーサイズ コットンTシャツ 白 | 0.967 | 0.950 | 1.000 |
| ビッグシルエット 綿Tシャツ ホワイト | オーバーサイズ コットンTシャツ 白 | 0.725 | 0.607 | 1.000 |
| デニム スリムフィット ブルー パンツ | スリムフィット デニムパンツ ブルー | 0.883 | 0.794 | 1.000 |

## インストール

### 必要な環境

* Python 3.8以上
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

ブラウザから直感的にExcelファイルをアップロードして名寄せを実行できます。クラウド環境またはローカル環境で利用可能です。

* **クラウド環境で使う:**
上記の「ブラウザ版デモ環境」URLにアクセスするだけで即座に利用可能です。
* **ローカル環境で立ち上げる:**
以下のコマンドを実行すると、ご自身のPCのブラウザが自動的に立ち上がります。
```bash
streamlit run app.py

```



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
WEIGHT_LEVENSHTEIN = 0.7

# Jaccard係数の重み (単語レベルの類似度)
WEIGHT_JACCARD = 0.3

```

**調整の指針:**

* 誤字脱字に強くしたい → `WEIGHT_LEVENSHTEIN` を高く
* 語順の違いに強くしたい → `WEIGHT_JACCARD` を高く

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
├── app.py                         # Web UI用スクリプト (新規追加)
├── comprehensive_data_matcher.py  # コアロジックおよびCLI実行用スクリプト
├── requirements.txt               # 必要なパッケージ
├── README.md                      # このファイル
├── LICENSE                        # ライセンス
├── .gitignore                     # Git除外設定
├── examples/                      # サンプルデータ
│   ├── yuragi_sample.xlsx
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

* 大量のデータ（10,000件以上）をCLIで処理する場合、実行に時間がかかることがあります
* Excelファイルは `.xlsx` 形式である必要があります（`.xls` は非対応）
* CLI版で入力ファイルが存在しない場合、エラーメッセージが表示されます

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

