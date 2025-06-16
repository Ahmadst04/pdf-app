import streamlit as st
import pdfplumber
import pandas as pd
import io
import matplotlib.pyplot as plt
import seaborn as sns
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image
from xlsxwriter.utility import xl_rowcol_to_cell

import cv2
import numpy as np

st.set_page_config(page_title="PDF to Excel with Smart Charts", layout="wide")
st.title("📄 PDF to Excel Converter with AI Chart Suggestion + OCR")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")
buffer = None
user_selections = []

def clean_data(df):
    for col in df.columns:
        df[col] = df[col].replace(r'[^\d.-]', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='ignore')
    return df

def score_table(df):
    numeric_cols = df.select_dtypes(include='number').shape[1]
    row_count = len(df)
    col_names = df.columns
    meaningful_names = sum(1 for name in col_names if len(name.strip()) > 2 and not name.lower().startswith("unnamed"))
    return numeric_cols * 2 + row_count + meaningful_names

def recommend_chart(df):
    numeric = df.select_dtypes('number').columns.tolist()
    categorical = df.select_dtypes('object').columns.tolist()
    if len(numeric) >= 2:
        return "Scatter", numeric[0], numeric[1]
    elif len(numeric) >= 1 and len(categorical) >= 1:
        return "Bar Chart", categorical[0], numeric[0]
    elif len(numeric) >= 1:
        return "Histogram", 'Index', numeric[0]
    else:
        return None, None, None

def text_to_table(text):
    lines = text.split("\n")
    lines = [line.strip() for line in lines if line.strip()]
    if len(lines) < 2:
        return None

    headers = lines[0].split()
    data = []

    for line in lines[1:]:
        row = line.split()
        if len(row) < len(headers):
            row += [''] * (len(headers) - len(row))
        elif len(row) > len(headers):
            row = row[:len(headers)]
        data.append(row)

    if not data:
        return None

    return pd.DataFrame(data, columns=headers)

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

def preprocess_image_for_ocr(pil_image):
    # Convert PIL Image to OpenCV grayscale
    img = np.array(pil_image)
    if img.ndim == 3:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

    # Apply Otsu's thresholding (binarization)
    _, img_bin = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Median blur to reduce noise
    img_bin = cv2.medianBlur(img_bin, 3)

    return img_bin

def extract_text_with_ocr(file_bytes):
    images = convert_from_bytes(file_bytes)
    extracted = []

    custom_config = r'--oem 3 --psm 6'  # LSTM OCR, single uniform block of text

    for img in images:
        preprocessed_img = preprocess_image_for_ocr(img)
        text = pytesseract.image_to_string(preprocessed_img, config=custom_config)
        extracted.append(text)
    return extracted

if uploaded_file:
    file_bytes = uploaded_file.read()
    all_tables = []
    raw_text_tables = []

    # Step 1: Try pdfplumber text/tables
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        has_text = False
        for i, page in enumerate(pdf.pages):
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                df = clean_data(df)
                all_tables.append(df)
            else:
                text = page.extract_text()
                if text:
                    has_text = True
                    df_text = text_to_table(text)
                    if df_text is not None:
                        df_text = clean_data(df_text)
                        raw_text_tables.append(df_text)

    all_tables.extend(raw_text_tables)

    # Step 2: Fallback to OCR if no usable tables
    if not all_tables:
        st.warning("⚠️ No tables or text found. Trying OCR (scanned document fallback)...")
        ocr_text_pages = extract_text_with_ocr(file_bytes)

        for text in ocr_text_pages:
            df = text_to_table(text)
            if df is not None:
                df = clean_data(df)
                all_tables.append(df)

    # Step 3: Display results
    if all_tables:
        table_scores = [score_table(df) for df in all_tables]
        best_table_idx = table_scores.index(max(table_scores))
        st.success(f"✅ Found {len(all_tables)} usable tables. Best table: Table {best_table_idx + 1}")

        for idx, df in enumerate(all_tables):
            st.markdown(f"---\n### 📑 Table {idx + 1}" + (" ✅ **Best Table**" if idx == best_table_idx else ""))
            st.dataframe(df)

            numeric_cols = df.select_dtypes(include='number').columns.tolist()
            categorical_cols = df.select_dtypes(include='object').columns.tolist()

            if 'Index' not in df.columns:
                df['Index'] = df.index

            auto_chart, x_col, y_col = recommend_chart(df)
            if auto_chart:
                st.info(f"📊 Recommended: {auto_chart} for Table {idx + 1}")
                try:
                    st.subheader(f"Preview: {auto_chart}")
                    fig, ax = plt.subplots()
                    if auto_chart == "Bar Chart":
                        sns.barplot(x=x_col, y=y_col, data=df, ax=ax)
                    elif auto_chart == "Line Chart":
                        sns.lineplot(x=x_col, y=y_col, data=df, ax=ax)
                    elif auto_chart == "Area Chart":
                        df.plot.area(x=x_col, y=y_col, ax=ax)
                    elif auto_chart == "Histogram":
                        df[y_col].plot.hist(ax=ax, bins=20)
                    elif auto_chart == "Box Plot":
                        sns.boxplot(y=df[y_col], ax=ax)
                    elif auto_chart == "Scatter":
                        sns.scatterplot(x=x_col, y=y_col, data=df, ax=ax)
                    st.pyplot(fig)
                except Exception as e:
                    st.warning(f"⚠️ Chart preview failed: {e}")

                user_selections.append({
                    "chart_type": auto_chart,
                    "x_col": x_col,
                    "y_col": y_col
                })
            else:
                st.warning("⚠️ No numeric data to generate chart.")
                user_selections.append(None)

        if st.button("📥 Generate Excel with Charts"):
            buffer = create_excel_with_charts(all_tables, user_selections)

    else:
        st.error("❌ No usable data found even after OCR.")

    if buffer:
        st.download_button("📥 Download Excel", buffer, file_name="converted_with_charts.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
