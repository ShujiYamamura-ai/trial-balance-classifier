# streamlit_app.py

import streamlit as st
import pandas as pd
import openai
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from tempfile import NamedTemporaryFile

st.set_page_config(page_title="即席ZBTベンチマーク用支援", layout="wide")
st.title("即席ZBTベンチマーク用支援")

st.markdown("""
このアプリは、Excel形式の残高試算表（A列のみ）に記載された項目を、OpenAIを用いて分類辞書に基づき即座にコスト分類できる支援ツールです。

🔴 分類辞書ファイルのアップロードは必須です！

* 残高試算表ファイル：A列の項目名のみ使用（最大500行）
* 分類辞書ファイル：5行目がヘッダ、B〜F列を使用
* 出力形式：分類結果付きExcel（Lv1#, Lv1name, Lv2#, Lv2name, 理由）
""")

api_key = st.text_input("OpenAI APIキーを入力してください", type="password")
if api_key:
    openai.api_key = api_key

category_dict = st.file_uploader("分類辞書ファイルをアップロード（5行目がヘッダ／必須）", type="xlsx")
if not category_dict:
    st.warning("⚠ 分類辞書ファイルがアップロードされていません。必ずアップロードしてください。")

@st.cache_data(show_spinner=False)
def load_category_table(file):
    df = pd.read_excel(file, header=4, usecols="B:F")
    df.columns = ['Lv1#', 'Lv1name', 'Lv2#', 'Lv2name', '説明']
    df = df[df['Lv1#'].notna()]
    df.fillna('', inplace=True)
    return df.astype(str)

@st.cache_data(show_spinner=False)
def generate_category_prompt(df):
    header = "分類表:\n    Lv1#;Lv1name;Lv2#;Lv2name;説明"
    rows = [f"    {';'.join(row)}" for row in df[['Lv1#', 'Lv1name', 'Lv2#', 'Lv2name', '説明']].values]
    return header + "\n" + "\n".join(rows)

def classify_text(text, category_prompt):
    if pd.isna(text) or str(text).strip() == '':
        return '', '', '', '', ''

    prompt = f"""
以下の情報を基にコスト費目を分類してください：

概要: {str(text).strip()}

{category_prompt}

指示:
* 上記の分類表の中から該当するものを選んでください。
* 出力形式は以下としてください：
  分類: Lv1#,Lv1name,Lv2#,Lv2name
  理由: <分類の根拠（任意）>
"""
    try:
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "あなたはコスト分類に長けた優秀な業務プロフェッショナルです。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )
        content = response.choices[0].message.content
        classification, reason = '', ''
        for line in content.splitlines():
            if "分類:" in line:
                classification = line.split("分類:")[1].strip()
            elif "理由:" in line:
                reason = line.split("理由:")[1].strip()
        parts = [x.strip() for x in classification.split(',')]
        if len(parts) == 4:
            return parts[0], parts[1], parts[2], parts[3], reason
        return '', '', '', '', reason
    except Exception as e:
        return 'エラー', '', '', '', str(e)

def adjust_excel_width(df, output):
    wb = load_workbook(output)
    ws = wb.active
    fixed_widths = {'A': 40, 'B': 6, 'C': 12, 'D': 6, 'E': 18, 'F': 80}
    for col, width in fixed_widths.items():
        ws.column_dimensions[col].width = width
    tempf = NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tempf.name)
    return tempf.name

col1, col2 = st.columns(2)

with col1:
    st.subheader("一括分類")
    uploaded_data = st.file_uploader("残高試算表ファイルのアップロード（A列）", type="xlsx")
    if st.button("ファイル全体分析開始"):
        if not category_dict:
            st.error("❌ 分類辞書ファイルがアップロードされていません。")
        elif not uploaded_data:
            st.error("❌ 残高試算表ファイルがアップロードされていません。")
        elif not api_key:
            st.error("❌ OpenAI APIキーが未入力です。")
        else:
            df = pd.read_excel(uploaded_data, usecols=[0], header=None)
            df.columns = ['テキスト']

            if len(df) > 500:
                st.error("最大500行まで処理可能です。ファイルを分割してください。")
            else:
                cat_df = load_category_table(category_dict)
                cat_prompt = generate_category_prompt(cat_df)

                results = []
                progress_bar = st.progress(0, text="GPTで分類中...")

                for i, text in enumerate(df['テキスト']):
                    result = classify_text(text, cat_prompt)
                    results.append(result)
                    progress_bar.progress((i + 1) / len(df))

                df[['Lv1#', 'Lv1name', 'Lv2#', 'Lv2name', '理由']] = pd.DataFrame(results, index=df.index)

                buffer = BytesIO()
                df.to_excel(buffer, index=False)
                buffer.seek(0)
                output_path = adjust_excel_width(df, buffer)

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="分類結果をダウンロード（Excel）",
                        data=f.read(),
                        file_name=f"classified_trial_balance_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

with col2:
    st.subheader("サンプルテスト（1件のみ分類）")
    sample_text = st.text_input("試したい費目を入力してください")
    if st.button("1件だけ分類開始"):
        if not category_dict:
            st.error("❌ 分類辞書ファイルがアップロードされていません。")
        elif not sample_text.strip():
            st.error("❌ 費目名が未入力です。")
        elif not api_key:
            st.error("❌ OpenAI APIキーが未入力です。")
        else:
            try:
                cat_df = load_category_table(category_dict)
                prompt = generate_category_prompt(cat_df)
                lv1, lv1name, lv2, lv2name, reason = classify_text(sample_text, prompt)
                st.success("分類結果：")
                st.write(f"Lv1#: {lv1}, Lv1name: {lv1name}")
                st.write(f"Lv2#: {lv2}, Lv2name: {lv2name}")
                if reason:
                    st.write("理由:")
                    st.markdown(reason)
            except Exception as e:
                st.error(f"エラーが発生しました: {str(e)}")
