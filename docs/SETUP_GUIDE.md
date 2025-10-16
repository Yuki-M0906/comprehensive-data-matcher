# セットアップガイド

このガイドでは、Comprehensive Data Matcherの詳細なセットアップ手順を説明します。

## 目次

1. [システム要件](#システム要件)
2. [Pythonのインストール](#pythonのインストール)
3. [プロジェクトのセットアップ](#プロジェクトのセットアップ)
4. [動作確認](#動作確認)
5. [トラブルシューティング](#トラブルシューティング)

## システム要件

### 必須要件
- Python 3.8以上
- pip（Pythonパッケージマネージャー）
- 512MB以上の空きメモリ
- 100MB以上の空きディスク容量

### 推奨環境
- Python 3.9以上
- 2GB以上の空きメモリ（大量データ処理時）
- SSD（処理速度向上のため）

## Pythonのインストール

### Windows

1. [Python公式サイト](https://www.python.org/downloads/)から最新版をダウンロード
2. インストーラーを実行
3. **重要**: 「Add Python to PATH」にチェックを入れる
4. インストール完了後、コマンドプロンプトで確認:
```cmd
python --version
```

### macOS

#### Homebrewを使用する場合（推奨）
```bash
# Homebrewのインストール（未インストールの場合）
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Pythonのインストール
brew install python@3.11
```

#### 公式インストーラーを使用する場合
1. [Python公式サイト](https://www.python.org/downloads/)からmacOS用をダウンロード
2. .pkgファイルを実行してインストール

### Linux (Ubuntu/Debian)

```bash
# システムのアップデート
sudo apt update
sudo apt upgrade

# Pythonのインストール
sudo apt install python3 python3-pip

# バージョン確認
python3 --version
pip3 --version
```

## プロジェクトのセットアップ

### ステップ1: リポジトリの取得

#### GitHubからクローン
```bash
git clone https://github.com/Yuki-M0906/comprehensive-data-matcher.git
cd comprehensive-data-matcher
```

#### ZIPファイルからダウンロード
1. GitHubリポジトリページの「Code」→「Download ZIP」
2. ダウンロードしたZIPを解凍
3. ターミナル/コマンドプロンプトで解凍したディレクトリに移動

### ステップ2: 仮想環境の作成（推奨）

仮想環境を使用することで、プロジェクトごとに独立したPython環境を作成できます。

#### Windows
```cmd
# 仮想環境の作成
python -m venv venv

# 仮想環境の有効化
venv\Scripts\activate
```

#### macOS/Linux
```bash
# 仮想環境の作成
python3 -m venv venv

# 仮想環境の有効化
source venv/bin/activate
```

仮想環境が有効化されると、コマンドラインに `(venv)` が表示されます。

### ステップ3: 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

インストールが成功すると、以下のパッケージがインストールされます:
- pandas (データ処理)
- openpyxl (Excel操作)
- rapidfuzz (高速文字列マッチング)

### ステップ4: サンプルデータの生成

初めて使用する場合は、サンプルデータを生成して動作確認できます:

```bash
python generate_sample_data.py
```

`examples/` ディレクトリに以下のファイルが作成されます:
- `yuragi_sample.xlsx`
- `correct_sample.xlsx`

## 動作確認

### サンプルデータでのテスト

1. サンプルデータをメインディレクトリにコピー:
```bash
# Windows
copy examples\yuragi_sample.xlsx yuragi.xlsx
copy examples\correct_sample.xlsx correct.xlsx

# macOS/Linux
cp examples/yuragi_sample.xlsx yuragi.xlsx
cp examples/correct_sample.xlsx correct.xlsx
```

2. スクリプトを実行:
```bash
python comprehensive_data_matcher.py
```

3. 実行結果の確認:
- コンソールにマッチング結果が表示される
- `result_combined_high_accuracy.xlsx` が生成される
- `matcher_log_high_accuracy.txt` にログが保存される

### 正常に動作している場合の出力例

```
[2025-04-30 10:00:00] [INFO] 高精度データマッチング処理を開始します。
[2025-04-30 10:00:00] [INFO] スコア重み設定: Levenshtein=0.7, Jaccard=0.3
[2025-04-30 10:00:01] [INFO] ファイル読み込み完了: yuragi.xlsx (30行)
[2025-04-30 10:00:01] [INFO] ファイル読み込み完了: correct.xlsx (15行)
...
[2025-04-30 10:00:05] [INFO] すべての処理が正常に完了しました。
```
## トラブルシューティング

### よくある問題と解決方法

#### 問題1: `python: command not found`

**原因**: Pythonがインストールされていない、またはPATHが通っていない

**解決方法**:
- Windowsの場合: Pythonを再インストールし、「Add Python to PATH」にチェック
- macOS/Linuxの場合: `python3` コマンドを使用してみる

#### 問題2: `pip: command not found`

**原因**: pipがインストールされていない

**解決方法**:
```bash
# Windows/macOS/Linux
python -m ensurepip --upgrade
```

#### 問題3: パッケージインストール時のエラー

**エラー例**: `error: Microsoft Visual C++ 14.0 is required`

**解決方法**:
- Windows: [Microsoft C++ Build Tools](https://visualstudio.microsoft.com/ja/visual-cpp-build-tools/)をインストール
- macOS: Xcode Command Line Toolsをインストール `xcode-select --install`
- Linux: build-essentialをインストール `sudo apt install build-essential`

#### 問題4: `ModuleNotFoundError: No module named 'xxx'`

**原因**: 必要なパッケージがインストールされていない

**解決方法**:
```bash
pip install -r requirements.txt
```

または個別にインストール:
```bash
pip install pandas openpyxl rapidfuzz
```

#### 問題5: Excelファイルが開けない

**エラー**: `File not found: yuragi.xlsx`

**解決方法**:
- ファイルがスクリプトと同じディレクトリにあるか確認
- ファイル名のスペルミスがないか確認
- ファイルが `.xlsx` 形式であることを確認（`.xls` は非対応）

#### 問題6: 文字化け

**原因**: エンコーディングの問題

**解決方法**:
- Windows: コマンドプロンプトで `chcp 65001` を実行してUTF-8に設定
- ログファイルをUTF-8対応のエディタで開く（VSCode、Notepad++など）

## 仮想環境の終了

作業が終わったら、仮想環境を終了できます:

```bash
deactivate
```
## 次のステップ

セットアップが完了したら:

1. [README.md](../README.md) で基本的な使い方を確認
2. [ALGORITHM.md](ALGORITHM.md) でアルゴリズムの詳細を学習
3. 自分のデータで試してみる

## ヒント

### パフォーマンス最適化

大量のデータを処理する場合:
- SSDを使用する
- 十分なメモリを確保する（最低2GB推奨）
- 不要なアプリケーションを閉じる

### データの準備

- Excelファイルは可能な限りシンプルにする
- 不要な書式設定を削除する
- データは1つのシートにまとめる

## それでも解決しない場合

1. [GitHubのIssue](https://github.com/Yuki-M0906/comprehensive-data-matcher/issues)を確認
2. 新しいIssueを作成して質問
3. エラーメッセージの全文を含める
4. 使用しているOSとPythonバージョンを記載

---

セットアップでお困りの際は、お気軽にお問い合わせください！
