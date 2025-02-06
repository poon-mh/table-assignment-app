import streamlit as st
import pandas as pd
import string
import io
import xlsxwriter
from datetime import datetime

# カスタムCSSでselectboxの幅を調整
st.markdown("""
    <style>
    div[data-testid="stSelectbox"] {
        width: 100% !important;
    }
    </style>
    """, unsafe_allow_html=True)

# セッション状態の初期化
if 'num_tables' not in st.session_state:
    st.session_state.num_tables = 7  # デフォルトのテーブル数
if 'tables' not in st.session_state:
    st.session_state.tables = {key: [] for key in string.ascii_uppercase[:st.session_state.num_tables]}
if 'unassigned' not in st.session_state:
    st.session_state.unassigned = []

# 🔧 **未振り分けリストから振り分け済みの人を削除**
assigned_people = {name for table in st.session_state.tables.values() for name in table}
st.session_state.unassigned = [name for name in st.session_state.unassigned if name not in assigned_people]

# 🔴 リセットボタン（画面右上）
st.markdown("""
    <style>
        .reset-container {
            position: absolute;
            top: 10px;
            right: 10px;
        }
        .reset-button {
            background-color: red;
            color: white;
            font-size: 16px;
            font-weight: bold;
            padding: 8px 16px;
            border-radius: 8px;
            border: none;
            cursor: pointer;
        }
        .reset-button:hover {
            background-color: darkred;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="reset-container">', unsafe_allow_html=True)
confirm_reset = st.checkbox("⚠ リセット確認")
if st.button("🔄 リセット", key="reset_confirm", help="リセットすると全データが消えます", disabled=not confirm_reset):
    st.session_state.tables = {key: [] for key in string.ascii_uppercase[:st.session_state.num_tables]}
    st.session_state.unassigned = []
    st.rerun()
st.markdown("</div>", unsafe_allow_html=True)

st.title("📌 テーブル振り分けシステム")

# 🔧 **テーブル数の設定**
num_tables = st.slider("🔢 作成するテーブル数（1～26）", 1, 26, st.session_state.num_tables)
if num_tables != st.session_state.num_tables:
    st.session_state.tables = {key: [] for key in string.ascii_uppercase[:num_tables]}
    st.session_state.num_tables = num_tables
    st.rerun()

# 📂 **名前リストのアップロード**
uploaded_file = st.file_uploader("📄 名前リスト（.txt または .xlsx）をアップロードしてください", type=["txt", "xlsx"])
if uploaded_file:
    try:
        if uploaded_file.name.endswith(".txt"):
            names = uploaded_file.getvalue().decode("utf-8").splitlines()
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, header=None)
            names = df[0].dropna().tolist()

        # すでに振り分けられている名前を除外
        st.session_state.unassigned = [name for name in names if name not in assigned_people]
        st.success(f"✅ {len(st.session_state.unassigned)} 人の名前を読み込みました！")
    except Exception as e:
        st.error(f"⚠ エラー: {e}")

# **振り分け済みの結果を読み込むボタン**
st.subheader("📤 振り分け済み結果の読み込み")
uploaded_assigned_file = st.file_uploader("🔄 振り分け結果を読み込む（.xlsx）", type=["xlsx"], key="assigned_file")
if uploaded_assigned_file:
    try:
        df_assigned = pd.read_excel(uploaded_assigned_file, sheet_name="振り分け結果")
        if not df_assigned.empty:
            assigned_data = df_assigned[["テーブル名", "名前"]].dropna()
            # 振り分け結果をセッションステートに反映させる際、既に同じ名前がテーブルに存在しない場合のみ追加
            for _, row in assigned_data.iterrows():
                table = row["テーブル名"]
                name = row["名前"]
                if name not in assigned_people:  # 名前がすでに振り分けられていない場合のみ追加
                    if table in st.session_state.tables:
                        st.session_state.tables[table].append(name)
                    else:
                        st.session_state.tables[table] = [name]
            st.success("✅ 振り分け結果を読み込みました！")
        else:
            st.warning("⚠ 振り分け結果が空です。")
    except Exception as e:
        st.error(f"⚠ エラー: {e}")

col1, col2 = st.columns(2)

# **未振り分けリスト**
with col1:
    st.subheader("📌 未振り分けの人")
    if st.session_state.unassigned:
        selected_name = st.selectbox("📝 振り分ける人を選択", st.session_state.unassigned, key="assign_name")
        assign_table = st.selectbox("📌 割り当てるテーブルを選択", list(st.session_state.tables.keys()), key="assign_table")

        if st.button("✅ 振り分け", key="assign_button"):
            st.session_state.tables[assign_table].append(selected_name)
            st.session_state.unassigned.remove(selected_name)  # 未振り分けリストから削除
            st.session_state.tables[assign_table].sort()  # 🔥 **振り分け後に自動ソート**
            st.rerun()

# **振り分け済みリスト**
with col2:
    st.subheader("🔄 振り分け済みの人の管理")
    selected_table = st.selectbox("📌 現在のテーブルを選択", list(st.session_state.tables.keys()), key="current_table")

    if selected_table and st.session_state.tables[selected_table]:
        selected_person = st.selectbox("📝 移動・削除する人を選択", st.session_state.tables[selected_table], key="selected_person")
        new_table = st.selectbox("📌 移動先のテーブルを選択", list(st.session_state.tables.keys()), key="new_table")

        col_move, col_remove = st.columns(2)
        with col_move:
            if st.button("🚀 移動", key="move_button"):
                st.session_state.tables[selected_table].remove(selected_person)
                st.session_state.tables[new_table].append(selected_person)
                st.session_state.tables[new_table].sort()  # 🔥 **移動後も自動ソート**
                st.rerun()

        with col_remove:
            if st.button("❌ 削除（未振り分けに戻す）", key="remove_button"):
                st.session_state.tables[selected_table].remove(selected_person)
                st.session_state.unassigned.append(selected_person)  # 未振り分けリストに戻す
                st.rerun()

# **各テーブルの表示**
st.subheader("📌 各テーブルの割り当て状況")
for table in st.session_state.tables:
    st.write(f"### テーブル {table}")
    if st.session_state.tables[table]:
        st.text("\n".join(sorted(st.session_state.tables[table])))  # 🔥 **テーブル内の名前を常にソート**
    else:
        st.write("（まだ誰もいません）")

# **Excel出力**
def export_to_excel():
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        all_data = [[table, name] for table, names in st.session_state.tables.items() for name in sorted(names)]
        df = pd.DataFrame(all_data, columns=["テーブル名", "名前"])
        df.to_excel(writer, sheet_name="振り分け結果", index=False)

    output.seek(0)
    return output

st.subheader("📤 振り分け結果の出力")
if st.button("📥 Excelダウンロード"):
    excel_data = export_to_excel()
    file_name = f"振り分け結果_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
    st.download_button(label="📂 Excelをダウンロード", data=excel_data, file_name=file_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
