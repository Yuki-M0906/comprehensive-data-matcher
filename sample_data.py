# -*- coding: utf-8 -*-
"""見本データ（Web 版の「見本で試す」と generate_sample_data.py で使う）。

どれも架空の名前です。
"""

# --- 衣料品の商品名（日本語） ---
APPAREL_MASTER = [
    'オーバーサイズ コットンTシャツ 白',
    'スリムフィット デニムパンツ ブルー',
    'ケーブルニット セーター グレー',
    'フード付き スウェットパーカー 黒',
    'ストレッチ チノパン ベージュ',
    'Vネック カーディガン ネイビー',
    'クルーネック スウェット 杢グレー',
    'テーパード カーゴパンツ カーキ',
    'ボタンダウン オックスフォードシャツ 白',
    'リブ編み タートルネック 黒',
    'ワイド ストレートパンツ ブラック',
    'ジップアップ フリースジャケット グレー',
    'ロング丈 トレンチコート ベージュ',
    'ラウンドネック ロングTシャツ 白',
    'ハイウエスト スキニーパンツ ブラック',
]

APPAREL_VARIANTS = [
    # オーバーサイズTシャツのゆらぎ（文字種の違い）
    'オーバーサイズ　コットンＴシャツ　白',  # 全角スペース・全角英字
    'ビッグシルエット 綿Tシャツ ホワイト',  # 言い換え
    'オーバーサイズ・コットン・Tシャツ（白）',  # 中黒と括弧
    'oversizeコットンティーシャツ白',  # 英字混在・スペースなし

    # デニムパンツのゆらぎ（語順の違い）
    'スリムフィット　デニム　パンツ　blue',  # 英語混在
    'デニム スリムフィット ブルー パンツ',  # 語順バラバラ
    'スリム・デニムパンツ／ブルー',  # 記号混在
    'ブルー　スリムデニム',  # 省略形

    # ニットセーターのゆらぎ（表現の違い）
    'ケーブル編みニットセーター　グレー',  # 詳細表記
    'ケーブルニット　灰色',  # 色の言い換え
    'cable knit セーター gray',  # 英語表記
    'ケーブルセーター　グレイ',  # カタカナ表記ゆれ

    # パーカーのゆらぎ（別名・語順）
    'パーカー フーディー 黒 スウェット',  # 語順バラバラ + 別名
    'フード付スウェット　ブラック',  # 漢字表記
    'スウェットパーカ　黒　フード',  # 長音なし
    'hoodie black フード付き',  # 英語混在

    # チノパンのゆらぎ
    'チノパンツ ストレッチ　ベージュ',  # 語順違い
    'ストレッチチノ　ベージュ色',  # 省略形
    'CHINO PANTS ベージュ ストレッチ',  # 英語表記

    # カーディガンのゆらぎ
    'Ｖネックカーディガン　ネイビー',  # 全角英字・スペースなし
    'ブイネック カーデ 紺',  # 略称・色の言い換え
    'V-NECK カーディガン navy',  # 記号・英語混在

    # スウェットのゆらぎ
    'クルーネック　スェット　杢グレー',  # 表記ゆれ
    '丸首スウェット　モクグレー',  # 言い換え
    'crew neck sweat 杢灰',  # 英語・漢字混在

    # カーゴパンツのゆらぎ
    'テーパードカーゴパンツカーキ',  # スペースなし
    'cargo pants カーキ テーパード',  # 語順違い
    'テーパード型　カーゴ　カーキ色',  # 型・色の付加

    # その他のゆらぎ例
    'ボタンダウンシャツ　白　オックスフォード',  # 語順違い
    'タートルネック リブ 黒',  # 語順違い
    'ワイドストレート　ブラック　パンツ',  # 語順違い
    'フリースジップアップ　グレー',  # 語順違い
    'ロング丈　トレンチ　ベージュ',  # 省略
    'ロンT 白 ラウンドネック',  # 略称・語順違い
]

# --- 取引先の会社名（英語・架空） ---
COMPANY_MASTER = [
    'Bluefield Logistics Inc.',
    'Harbor Point Foods Co., Ltd.',
    'Sakura Precision Electronics Corp.',
    'Greenleaf Pharmacy Group',
    'Northgate Software Solutions LLC',
    'Maple & Stone Architects',
    'Silverline Transport Co.',
    'Kitakaze Trading Company',
    'Summit Peak Outdoor Gear',
    'Aoba Medical Clinic',
]

COMPANY_VARIANTS = [
    'BLUEFIELD LOGISTICS INC',  # 大文字・ピリオドなし
    'Bluefield Logistics, Inc.',  # カンマ
    'Blue Field Logistics',  # 分かち書き・社名の種類なし
    'ＢＬＵＥＦＩＥＬＤ ＬＯＧＩＳＴＩＣＳ',  # 全角英字
    'Harbor Point Foods Co Ltd',  # 記号なし
    'HarborPoint Foods',  # 続け書き
    'Harbour Point Foods Co., Ltd.',  # 英国式のつづり
    'Sakura Precision Electronics Corporation',  # 略さない
    'SAKURA PRECISION ELEC. CORP.',  # 略語
    'Sakura Electronics',  # 省略
    'Greenleaf Pharmacy Grp',  # 略語
    'Green Leaf Pharmacy Group',  # 分かち書き
    'Northgate Software Solutions, LLC',  # カンマ
    'North Gate Software',  # 省略
    'Maple and Stone Architects',  # & → and
    'Maple & Stone Architect',  # 単数形
    'Silverline Transport Company',  # 略さない
    'Silver Line Transport Co',  # 分かち書き
    'Kitakaze Trading Co.',  # 略語
    'KITAKAZE TRADING COMPANY',  # 大文字
    'Summit Peak Outdoor',  # 省略
    'Summit-Peak Outdoor Gear',  # ハイフン
    'Aoba Medical Clinic',  # 完全一致
    'Aoba Med. Clinic',  # 略語
]

SAMPLES = {
    "apparel": {"variants": APPAREL_VARIANTS, "master": APPAREL_MASTER, "column": "商品名"},
    "companies": {"variants": COMPANY_VARIANTS, "master": COMPANY_MASTER, "column": "Company"},
}
