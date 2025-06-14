import streamlit as st
import pdfplumber
import pandas as pd
import io
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Smart PDF to Excel", layout="wide")
st.title("📄 PDF to Excel Converter with Smart Charts in App & Excel")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")

chart_types = ["Bar Chart", "Line Chart", "Area Chart", "Histogram", "Box Plot", "Scatter"]

def summarize_dataframe(df):
    summary = df.describe(include='all').transpose()
    st.subheader("📌 Data Summary")
    st.dataframe(summary)

def plot_chart(df, x_col, y_col, chart_type):
    st.subheader(f"📊 {chart_type} of '{y_col}' vs '{x_col}'")
    fig, ax = plt.subplots()
    try:
        if chart_type == "Bar Chart":
            sns.barplot(x=x_col, y=y_col, data=df, ax=ax)
        elif chart_type == "Line Chart":
            sns.lineplot(x=x_col, y=y_col, data=df, ax=ax)
        elif chart_type == "Area Chart":
            df.plot.area(x=x_col, y=y_col, ax=ax)
        elif chart_type == "Histogram":
            df[y_col].hist(ax=ax, bins=20)
        elif chart_type == "Box Plot":
            sns.boxplot(y=df[y_col], ax=ax)
        elif chart_type == "Scatter":
            sns.scatterplot(x=x_col, y=y_col, data=df, ax=ax)
        plt.xticks(rotation=45)
        st.pyplot(fig)
    except Exception as e:
        st.warning(f"⚠️ Could not generate {chart_type}: {e}")

def clean_data(df):
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='ignore')
    return df

if uploaded_file:
    with pdfplumber.open(uploaded_file) as pdf:
        all_tables = []
        for i, page in enumerate(pdf.pages):
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                all_tables.append(df)

    if all_tables:
        st.success(f"✅ Found {len(all_tables)} table(s) in the PDF.")
        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            for idx, table in enumerate(all_tables):
                sheet_name = f'Table{idx + 1}'
                table = clean_data(table)
                table.to_excel(writer, sheet_name=sheet_name, index=False)
                st.markdown(f"---\n### 📑 Table {idx + 1}")
                st.dataframe(table)
                summarize_dataframe(table)

                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                numeric_cols = table.select_dtypes(include='number').columns.tolist()
                categorical_cols = table.select_dtypes(include='object').columns.tolist()

                if numeric_cols:
                    y_col = numeric_cols[0]
                    x_col = categorical_cols[0] if categorical_cols else None

                    if x_col:
                        max_row = len(table) + 1  # +1 for header
