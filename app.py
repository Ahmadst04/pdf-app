import streamlit as st
import pdfplumber
import pandas as pd
import io
import matplotlib.pyplot as plt
import seaborn as sns
from xlsxwriter.utility import xl_rowcol_to_cell

st.set_page_config(page_title="PDF to Excel with Charts", layout="wide")
st.title("📄 PDF to Excel Converter with Embedded Charts")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")

chart_types = ["Bar Chart", "Line Chart", "Area Chart", "Histogram", "Box Plot", "Scatter"]
user_selections = []
buffer = None

def clean_data(df):
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='ignore')
    return df

def create_excel_with_charts(tables, selections):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        for idx, df in enumerate(tables):
            sheet_name = f'Table{idx + 1}'
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]

            selection = selections[idx]
            if not selection:
                continue

            chart_type = selection["chart_type"]
            x_col = selection["x_col"]
            y_col = selection["y_col"]

            if x_col not in df.columns or y_col not in df.columns:
                continue

            chart_map = {
                "Bar Chart": "column",
                "Line Chart": "line",
                "Area Chart": "area",
                "Scatter": "scatter",
            }

            if chart_type not in chart_map:
                continue

            chart = workbook.add_chart({'type': chart_map[chart_type]})
            max_row = len(df)

            chart.add_series({
                'categories': f"='{sheet_name}'!${xl_rowcol_to_cell(1, df.columns.get_loc(x_col))}:${xl_rowcol_to_cell(max_row, df.columns.get_loc(x_col))}",
                'values':     f"='{sheet_name}'!${xl_rowcol_to_cell(1, df.columns.get_loc(y_col))}:${xl_rowcol_to_cell(max_row, df.columns.get_loc(y_col))}",
                'name':       f"{y_col} vs {x_col}"
            })

            chart.set_title({'name': f"{chart_type} of {y_col} vs {x_col}"})
            chart.set_x_axis({'name': x_col})
            chart.set_y_axis({'name': y_col})
            worksheet.insert_chart("J2", chart)

    output.seek(0)
    return output

if uploaded_file:
    with pdfplumber.open(uploaded_file) as pdf:
        all_tables = []
        for i, page in enumerate(pdf.pages):
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                df = clean_data(df)
                all_tables.append(df)

    if all_tables:
        st.success(f"✅ Found {len(all_tables)} table(s)")

        for idx, df in enumerate(all_tables):
            st.markdown(f"---\n### 📑 Table {idx + 1}")
            st.dataframe(df)

            numeric_cols = df.select_dtypes(include='number').columns.tolist()
            categorical_cols = df.select_dtypes(include='object').columns.tolist()

            if numeric_cols:
                chart_type = st.selectbox(
                    f"📈 Chart type for Table {idx + 1}:", chart_types, key=f"chart_{idx}"
                )
                y_col = st.selectbox(
                    f"🧮 Y-axis (numeric):", numeric_cols, key=f"ycol_{idx}"
                )

                x_col = 'Index'
                if categorical_cols:
                    x_col = st.selectbox(
                        f"🏷️ X-axis:", categorical_cols + ['Index'], key=f"xcol_{idx}"
                    )
                    if x_col == 'Index':
                        df['Index'] = df.index
                else:
                    df['Index'] = df.index
                    x_col = 'Index'

                user_selections.append({
                    "chart_type": chart_type,
                    "x_col": x_col,
                    "y_col": y_col
                })

                st.subheader(f"Preview: {chart_type}")
                fig, ax = plt.subplots()
                try:
                    if chart_type == "Bar Chart":
                        sns.barplot(x=x_col, y=y_col, data=df, ax=ax)
                    elif chart_type == "Line Chart":
                        sns.lineplot(x=x_col, y=y_col, data=df, ax=ax)
                    elif chart_type == "Area Chart":
                        df.plot.area(x=x_col, y=y_col, ax=ax)
                    elif chart_type == "Histogram":
                        df[y_col].plot.hist(ax=ax, bins=20)
                    elif chart_type == "Box Plot":
                        sns.boxplot(y=df[y_col], ax=ax)
                    elif chart_type == "Scatter":
                        sns.scatterplot(x=x_col, y=y_col, data=df, ax=ax)
                    plt.xticks(rotation=45)
                    st.pyplot(fig)
                except Exception as e:
                    st.warning(f"Chart preview failed: {e}")
            else:
                st.info("No numeric column found. Skipping chart options.")
                user_selections.append(None)

        st.markdown("---")
        if st.button("📥 Generate Excel with Charts"):
            buffer = create_excel_with_charts(all_tables, user_selections)

    else:
        st.warning("⚠️ No tables found in PDF.")

    if buffer:
        st.download_button("📥 Download Excel", buffer, file_name="converted_with_charts.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
