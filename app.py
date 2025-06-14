import streamlit as st
import pdfplumber
import pandas as pd
import io
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Smart PDF to Excel", layout="wide")
st.title("📄 PDF to Excel Converter with Smart Chart Generator")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")

chart_types = ["Bar Chart", "Line Chart", "Area Chart", "Histogram", "Box Plot", "Scatter"]

def summarize_dataframe(df):
    """Generate basic summary of the dataframe"""
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
    """Try converting each column to numeric if possible"""
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
                table.to_excel(writer, sheet_name=sheet_name, index=False)

                st.markdown(f"---\n### 📑 Table {idx + 1}")
                st.dataframe(table)

                table = clean_data(table)
                summarize_dataframe(table)

                numeric_cols = table.select_dtypes(include=['number']).columns.tolist()
                categorical_cols = table.select_dtypes(include=['object']).columns.tolist()

                if numeric_cols:
                    chart_type = st.selectbox(
                        f"📈 Select chart type for Table {idx + 1}:", chart_types, key=f"chart_{idx}"
                    )
                    y_col = st.selectbox(f"🧮 Select Y-axis (numeric):", numeric_cols, key=f"ycol_{idx}")

                    if categorical_cols:
                        x_col = st.selectbox(f"🏷️ Select X-axis (category or index):", categorical_cols + ['Index'], key=f"xcol_{idx}")
                        if x_col == 'Index':
                            table['Index'] = table.index
                    else:
                        x_col = 'Index'
                        table['Index'] = table.index

                    plot_chart(table, x_col, y_col, chart_type)
                else:
                    st.info("ℹ️ No numeric columns found for chart generation.")

        buffer.seek(0)
        st.download_button("📥 Download Excel", buffer, file_name="converted.xlsx")
    else:
        st.warning("⚠️ No tables were found in the PDF.")
