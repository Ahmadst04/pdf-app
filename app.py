import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import cv2
import numpy as np
from PIL import Image
from collections import defaultdict

# --- Configuration and Setup ---
st.set_page_config(page_title="PDF OCR Table Extractor", layout="wide")
st.title("📸 Image-based PDF to Excel Converter (with OCR)")

# --- Global Variables ---
# Buffer to store the Excel file in memory
excel_buffer = None

# --- Helper Functions ---

def preprocess_image(pil_image):
    """
    Converts a PIL image to a pure black and white image using Otsu's thresholding
    for better OCR accuracy. Also applies a minor morphological opening for noise reduction.
    """
    img = np.array(pil_image)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    
    # Apply Otsu thresholding to binarize the image
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Optional: Morphological clean-up to remove small noise or gaps
    # A small kernel (e.g., 1x1 or 2x2) can help
    kernel = np.ones((1, 1), np.uint8) # Adjust kernel size if needed
    cleaned = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)

    return cleaned

def ocr_image_to_table(image, x_tolerance=10, y_tolerance=5):
    """
    Performs OCR on an image and attempts to extract structured table data
    by spatially grouping text into rows and columns.

    Args:
        image (numpy.ndarray): The preprocessed (e.g., binary) image.
        x_tolerance (int): Maximum horizontal distance (pixels) between words
                           to be considered part of the same column.
        y_tolerance (int): Maximum vertical distance (pixels) between lines
                           to be considered part of the same row.

    Returns:
        pandas.DataFrame: A DataFrame representing the extracted table, or an empty DataFrame if no text is found.
    """
    # Use psm 6 for a single uniform block of text (good for tables)
    # You might experiment with other psm values (e.g., 3 for default, 11 for sparse text)
    custom_config = r'--psm 6' 
    data = pytesseract.image_to_data(image, config=custom_config, output_type=pytesseract.Output.DATAFRAME)
    
    # Filter out rows with no detected text and drop NaN values
    data = data.dropna(subset=['text']).query('text.str.strip() != ""', engine='python')

    if data.empty:
        return pd.DataFrame() # Return empty DataFrame if no text is detected

    # Sort by top position first, then by left position
    data = data.sort_values(by=['top', 'left']).reset_index(drop=True)

    # Group words into potential rows based on vertical proximity
    rows = defaultdict(list)
    current_row_top = -1
    row_idx = 0

    for idx, row in data.iterrows():
        if current_row_top == -1 or abs(row['top'] - current_row_top) > y_tolerance:
            # Start a new row if the vertical distance is too large
            current_row_top = row['top']
            row_idx += 1
        rows[row_idx].append(row)

    # Process each row to align text into columns
    table_data = []
    
    # Determine a global set of column boundaries
    # This is a simplified approach. For highly complex tables, more advanced clustering
    # or line detection might be needed.
    all_left_coords = sorted(data['left'].unique())
    column_boundaries = []
    if all_left_coords:
        column_boundaries.append(all_left_coords[0])
        for i in range(1, len(all_left_coords)):
            if all_left_coords[i] - all_left_coords[i-1] > x_tolerance:
                column_boundaries.append(all_left_coords[i])

    # Refine column boundaries by merging close ones
    refined_column_boundaries = []
    if column_boundaries:
        refined_column_boundaries.append(column_boundaries[0])
        for i in range(1, len(column_boundaries)):
            if column_boundaries[i] - refined_column_boundaries[-1] > x_tolerance:
                refined_column_boundaries.append(column_boundaries[i])
            # else: merge, effectively skipping this boundary if too close

    # Now, populate the table data based on these column boundaries
    for r_idx in sorted(rows.keys()):
        row_words = sorted(rows[r_idx], key=lambda x: x['left'])
        
        current_table_row = [''] * len(refined_column_boundaries) # Initialize with empty strings
        
        for word_data in row_words:
            word_text = str(word_data['text']).strip()
            word_left = word_data['left']
            
            # Find the appropriate column for the word
            col_index = -1
            for i, boundary_left in enumerate(refined_column_boundaries):
                # If the word's left edge is within a column's tolerance
                if abs(word_left - boundary_left) <= x_tolerance:
                    col_index = i
                    break
                # If the word is past a boundary and closer to the next, assign to next
                elif word_left > boundary_left:
                    col_index = i # Tentatively assign to current column
                    if i + 1 < len(refined_column_boundaries) and word_left > refined_column_boundaries[i+1] - x_tolerance:
                        col_index = i + 1 # Assign to next if closer
            
            if col_index == -1 and refined_column_boundaries: # If still not assigned, assign to closest
                col_index = np.argmin(np.abs(np.array(refined_column_boundaries) - word_left))


            if col_index != -1 and col_index < len(current_table_row):
                # Append text to existing content in the cell, separated by space
                if current_table_row[col_index]:
                    current_table_row[col_index] += " " + word_text
                else:
                    current_table_row[col_index] = word_text
        
        table_data.append(current_table_row)

    # Ensure all rows have the same number of columns
    max_cols = max(len(row) for row in table_data) if table_data else 0
    final_table_data = [row + [''] * (max_cols - len(row)) for row in table_data]

    df = pd.DataFrame(final_table_data)
    
    # Drop rows/columns that are entirely empty after OCR processing
    df.replace('', np.nan, inplace=True) # Replace empty strings with NaN for dropna
    df.dropna(how='all', inplace=True) # Drop rows where all cells are NaN
    df.dropna(axis=1, how='all', inplace=True) # Drop columns where all cells are NaN
    df.replace(np.nan, '', inplace=True) # Convert NaN back to empty string for display

    return df

def create_excel_file(tables, original_filename="ocr_tables"):
    """
    Creates an Excel file from a list of Pandas DataFrames.
    Each DataFrame is written to a separate sheet.
    """
    output = io.BytesIO()
    # Use xlsxwriter engine for better compatibility and features
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for idx, df in enumerate(tables):
            # Sheet names are 'Page_1_Table_1', 'Page_2_Table_1', etc.
            # Assuming 'tables' might contain multiple tables per page,
            # or one table per page. Adjust sheet naming if multiple tables
            # per page need distinct names. For now, it's just Table{idx+1}.
            df.to_excel(writer, sheet_name=f'Table_{idx + 1}', index=False, header=False)
    output.seek(0)
    return output, original_filename.replace(".pdf", ".xlsx")

# --- Streamlit UI Layout ---

st.markdown("""
Upload a **scanned PDF** or a PDF with **complex layouts** where standard table extraction fails.
This tool uses Optical Character Recognition (OCR) to extract text and then attempts to structure it into tables.
""")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file:
    st.write(f"File uploaded: **{uploaded_file.name}** (Size: {round(len(uploaded_file.getbuffer()) / (1024 * 1024), 2)} MB)")
    
    # Use a placeholder for the spinner
    processing_status = st.empty()

    try:
        with processing_status.container():
            st.info("Converting PDF pages to images... (This might take a moment for large PDFs)")
        file_bytes = uploaded_file.read()
        images = convert_from_bytes(file_bytes, dpi=300) # Higher DPI for better OCR

        all_extracted_dfs = []

        for page_num, img in enumerate(images):
            with processing_status.container():
                st.info(f"Processing Page {page_num + 1} of {len(images)} for OCR...")
            
            st.markdown(f"---\n### 🖼 Page {page_num + 1} Visuals")
            
            # Display original image
            st.image(img, caption=f"Original Page {page_num + 1}", use_column_width=True)

            # Preprocess image for OCR
            processed_img_np = preprocess_image(img)
            # Convert NumPy array back to PIL Image for Streamlit's st.image
            processed_img_pil = Image.fromarray(processed_img_np)
            
            # Show preprocessed black & white version
            st.image(processed_img_pil, caption="🧼 Preprocessed (Black & White)", use_column_width=True, channels="GRAY")

            # Perform OCR and attempt table extraction
            df_page = ocr_image_to_table(processed_img_np)
            
            if not df_page.empty:
                st.success(f"✅ Table extracted from Page {page_num + 1}!")
                st.dataframe(df_page)
                all_extracted_dfs.append(df_page)
            else:
                st.warning(f"⚠️ No significant table text detected on Page {page_num + 1}.")
        
        # Clear the processing status message
        processing_status.empty()

        if all_extracted_dfs:
            if st.button("📥 Download All Extracted Tables as Excel"):
                with st.spinner("Generating Excel file..."):
                    excel_buffer, excel_filename = create_excel_file(all_extracted_dfs, uploaded_file.name)
                    st.download_button(
                        label="Click to Download Excel",
                        data=excel_buffer,
                        file_name=excel_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("Excel file ready for download!")
        else:
            st.error("❌ No usable tables could be extracted from the entire PDF via OCR.")

    except pytesseract.TesseractNotFoundError:
        st.error("""
            **Tesseract OCR engine not found!**
            Please install Tesseract OCR on your system.
            * **Windows:** Download from [Tesseract-OCR GitHub](https://github.com/UB-Mannheim/tesseract/wiki).
            * **macOS:** `brew install tesseract` (if you have Homebrew).
            * **Linux (Ubuntu/Debian):** `sudo apt-get install tesseract-ocr`
            After installation, you might need to specify the path to `tesseract.exe` in your code:
            `pytesseract.pytesseract.tesseract_cmd = r'<path_to_tesseract.exe>'`
            (e.g., `r'C:\Program Files\Tesseract-OCR\tesseract.exe'`)
        """)
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
        st.info("Please ensure you have `poppler` installed for `pdf2image` to work correctly.")
        st.markdown("""
            **Poppler Installation:**
            * **Windows:** Download binaries from [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases) and add to PATH.
            * **macOS:** `brew install poppler`
            * **Linux (Ubuntu/Debian):** `sudo apt-get install poppler-utils`
        """)


st.markdown("---")
st.markdown("### How this OCR-based converter works:")
st.markdown("""
1.  **PDF to Images:** Each page of your PDF is converted into a high-resolution image.
2.  **Image Preprocessing:** Images are converted to pure black and white to improve OCR accuracy.
3.  **OCR Text Extraction:** `pytesseract` performs OCR to detect all text and its bounding box (position) on each image.
4.  **Spatial Table Reconstruction:** The extracted text and its positions are then analyzed to group words into rows and columns, attempting to reconstruct the table structure.
5.  **Excel Output:** All reconstructed tables from different pages are compiled into a single Excel file, with each table on a separate sheet.
""")
st.markdown("---")
st.markdown("### Tips for better OCR results:")
st.markdown("""
* **High-Quality Scans:** The clearer and higher resolution your PDF scans are, the better the OCR accuracy.
* **Clean Layouts:** OCR struggles with highly complex or overlapping text. Simple, clear table structures yield the best results.
* **Adjust Tolerances:** In the `ocr_image_to_table` function, you can experiment with `x_tolerance` (for horizontal column alignment) and `y_tolerance` (for vertical row grouping) to fine-tune the table detection for your specific PDFs.
* **Tesseract Configuration (`--psm`):** The `pytesseract.image_to_data` function uses `--psm 6` (Page Segmentation Mode: Assume a single uniform block of text). For different PDF layouts, you might try other `psm` values (e.g., `3` for default, `11` for sparse text). This would require modifying the `custom_config` variable.
""")
