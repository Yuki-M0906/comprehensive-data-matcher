import streamlit as st
import pandas as pd
import io
from comprehensive_data_matcher import process_matching # 分離したコアロジックをインポート

st.set_page_config(page_title="高精度名寄せツール", layout="wide")

st.title("高精度名寄せツール (Web版)")
st.write("Excelファイルをアップロードするだけで、レーベンシュタイン距離とJaccard係数を用いたハイブリッド名寄せを実行します。")

# サイドバーで重みを動的に調整できるようにする
st.sidebar.header("⚙️ アルゴリズム設定")
weight_lev = st.sidebar.slider("レーベンシュタイン距離の重み (誤字脱字に強い)", 0.0, 1.0, 0.7, 0.1)
weight_jac = st.sidebar.slider("Jaccard係数の重み (語順の違いに強い)", 0.0, 1.0, 0.3, 0.1)

st.sidebar.info("2つの重みの合計が1.0になるように設定するのがおすすめです。")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. ゆらぎデータ")
    yuragi_file = st.file_uploader("ゆらぎデータ (Excel) をアップロード", type=["xlsx"], key="yuragi")
    yuragi_col = st.text_input("対象の列名", value="商品名", key="yuragi_col")

with col2:
    st.subheader("2. マスターデータ")
    correct_file = st.file_uploader("マスターデータ (Excel) をアップロード", type=["xlsx"], key="correct")
    correct_col = st.text_input("対象の列名", value="商品名", key="correct_col")

if yuragi_file and correct_file:
    if st.button("名寄せを実行する", type="primary"):
        try:
            # Excelをメモリ上に読み込み
            df_y = pd.read_excel(yuragi_file)
            df_c = pd.read_excel(correct_file)
            
            with st.spinner('AIが名寄せ計算を行っています...'):
                # ステップ1で作った関数に丸投げ！
                result_df = process_matching(
                    df_yuragi_orig=df_y,
                    df_correct=df_c,
                    yuragi_col=yuragi_col,
                    correct_col=correct_col,
                    weight_lev=weight_lev,
                    weight_jac=weight_jac
                )
            
            st.success("名寄せが完了しました！")
            
            # 結果のプレビュー
            st.dataframe(result_df.head(50)) # 最初の50件を表示
            
            # ダウンロードボタンの生成
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, sheet_name="マッチング結果", index=False)
            
            st.download_button(
                label="結果をExcelでダウンロード",
                data=output.getvalue(),
                file_name="result_combined_high_accuracy.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
