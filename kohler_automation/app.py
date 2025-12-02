import streamlit as st
import pandas as pd
import io
import asyncio
import sys
import os
import tkinter as tk
from tkinter import filedialog
from scrape_kohler import process_codes

# Fix for Windows Event Loop Policy
if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# Page Config
st.set_page_config(page_title="Kohler Automation", layout="wide")

# Sidebar Navigation
st.sidebar.title("Navigation")
page = st.sidebar.radio("Select Tool", ["Kohler Scraper", "Folder Scanner", "PDF Highlighter"])

if page == "Kohler Scraper":
    st.title("Kohler Product Scraper")
    st.markdown("Enter product codes below (one per line) to scrape data from Kohler.com.")

    # Input Area
    input_text = st.text_area("Product Codes", height=200, placeholder="K-23475-4-AF\nK-77748T-4-0")

    if st.button("Run Scraper"):
        if not input_text.strip():
            st.warning("Please enter at least one code.")
        else:
            # Parse input
            codes = [line.strip() for line in input_text.split('\n') if line.strip()]
            
            st.info(f"Processing {len(codes)} codes...")
            
            # Run Scraper
            try:
                with st.spinner('Scraping in progress... This may take a while.'):
                    result_df = process_codes(codes)
                
                st.success("Scraping Completed!")
                
                # Display Results
                st.dataframe(result_df)
                
                # Download Button
                # Convert DF to Excel in memory
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                output.seek(0)
                
                st.download_button(
                    label="Download Results (Excel)",
                    data=output,
                    file_name="kohler_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"An error occurred: {e}")

elif page == "Folder Scanner":
    st.title("Folder Scanner")
    st.markdown("Scan a folder and its subfolders to list all files with their paths.")

    # Ensure scan_df is in session state
    if 'scan_df' not in st.session_state:
        st.session_state['scan_df'] = None

    # Callback for Browse Button
    def browse_callback():
        try:
            root = tk.Tk()
            root.withdraw()
            root.wm_attributes('-topmost', 1)
            selected_folder = filedialog.askdirectory(master=root)
            root.destroy()
            if selected_folder:
                st.session_state['path_input'] = selected_folder
        except Exception as e:
            st.error(f"Could not open file dialog: {e}")

    # Callback for Save As Button
    def save_as_callback():
        if st.session_state['scan_df'] is not None:
            try:
                root = tk.Tk()
                root.withdraw()
                root.wm_attributes('-topmost', 1)
                file_path = filedialog.asksaveasfilename(
                    master=root,
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
                )
                root.destroy()
                if file_path:
                    st.session_state['scan_df'].to_excel(file_path, index=False)
                    st.session_state['save_message'] = {"type": "success", "text": f"Saved to {file_path}"}
            except Exception as e:
                st.session_state['save_message'] = {"type": "error", "text": f"Could not save file: {e}"}

    col1, col2 = st.columns([3, 1])
    
    with col1:
        # The text_input is bound to st.session_state['path_input']
        folder_path_input = st.text_input("Directory Path", key='path_input')

    with col2:
        # Add some spacing to align button with input box
        st.write("") 
        st.write("")
        st.button("Browse Folder", on_click=browse_callback)

    if st.button("Scan Folder"):
        # Use the value from the text input
        target_path = st.session_state['path_input']
        
        if not target_path or not os.path.isdir(target_path):
            st.error(f"Please enter a valid directory path. Invalid path: '{target_path}'")
        else:
            file_data = []
            try:
                with st.spinner(f"Scanning {target_path}..."):
                    for root, dirs, files in os.walk(target_path):
                        for file in files:
                            full_path = os.path.join(root, file)
                            file_data.append({"File Name": file, "File Path": full_path})
                
                if not file_data:
                    st.warning("No files found in the selected directory.")
                    st.session_state['scan_df'] = None
                else:
                    df = pd.DataFrame(file_data)
                    st.session_state['scan_df'] = df
                    st.success(f"Found {len(df)} files!")
                    
                    # Clear previous save messages
                    if 'save_message' in st.session_state:
                        del st.session_state['save_message']

            except Exception as e:
                st.error(f"An error occurred during scanning: {e}")
                st.session_state['scan_df'] = None

    # Display Results if available
    if st.session_state['scan_df'] is not None:
        st.dataframe(st.session_state['scan_df'])
        
        # Display save message if exists
        if 'save_message' in st.session_state:
            msg = st.session_state['save_message']
            if msg['type'] == 'success':
                st.success(msg['text'])
            else:
                st.error(msg['text'])

        col_d1, col_d2 = st.columns(2)
        
        with col_d1:
            # Download Button (Browser)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state['scan_df'].to_excel(writer, index=False)
            output.seek(0)
            
            st.download_button(
                label="Download Results (Browser)",
                data=output,
                file_name="folder_scan_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col_d2:
            # Save As Button (Native)
            st.button("Save to File... (Native)", on_click=save_as_callback)

elif page == "PDF Highlighter":
    st.title("PDF Highlighter")
    st.markdown("Search and highlight text in PDF files (supports scanned PDFs via OCR).")

    import fitz  # PyMuPDF
    import pytesseract
    from PIL import Image

    # Set Tesseract Path explicitly
    # Set Tesseract Path explicitly if on Windows and path exists
    tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
    # On Linux (Streamlit Cloud), it will use the default 'tesseract' command from PATH

    # File Uploader
    uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])
    search_text = st.text_input("Text to Search & Highlight")
    enable_ocr = st.checkbox("Enable OCR (for scanned files)", value=False)

    if uploaded_file and search_text:
        if st.button("Process PDF"):
            try:
                with st.spinner("Processing PDF..."):
                    # Open PDF from memory
                    pdf_document = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                    total_matches = 0
                    
                    for page_num in range(len(pdf_document)):
                        page = pdf_document[page_num]
                        
                        # 1. Standard Text Search
                        text_instances = page.search_for(search_text)
                        
                        # Add highlights for text matches
                        for inst in text_instances:
                            highlight = page.add_highlight_annot(inst)
                            highlight.update()
                            total_matches += 1
                        
                        # 2. OCR Search (if enabled)
                        if enable_ocr:
                            try:
                                # Render page to image (300 DPI for better OCR)
                                pix = page.get_pixmap(dpi=300)
                                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                                
                                # Get OCR data with bounding boxes
                                ocr_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
                                
                                n_boxes = len(ocr_data['text'])
                                for i in range(n_boxes):
                                    if search_text.lower() in ocr_data['text'][i].lower():
                                        (x, y, w, h) = (ocr_data['left'][i], ocr_data['top'][i], ocr_data['width'][i], ocr_data['height'][i])
                                        
                                        # Convert Image coordinates to PDF coordinates
                                        # PDF units are usually 72 DPI, Image is 300 DPI
                                        scale_x = page.rect.width / pix.width
                                        scale_y = page.rect.height / pix.height
                                        
                                        pdf_rect = fitz.Rect(
                                            x * scale_x, 
                                            y * scale_y, 
                                            (x + w) * scale_x, 
                                            (y + h) * scale_y
                                        )
                                        
                                        highlight = page.add_highlight_annot(pdf_rect)
                                        highlight.update()
                                        total_matches += 1
                                        
                            except pytesseract.TesseractNotFoundError:
                                st.error("Tesseract OCR is not found. Please install Tesseract and add it to your PATH to use OCR features.")
                                break # Stop processing pages if OCR is missing
                            except Exception as e:
                                st.warning(f"OCR failed on page {page_num + 1}: {e}")

                    if total_matches > 0:
                        st.success(f"Found and highlighted {total_matches} occurrences.")
                        
                        # Save to buffer
                        output_pdf = io.BytesIO()
                        pdf_document.save(output_pdf)
                        output_pdf.seek(0)
                        
                        st.download_button(
                            label="Download Highlighted PDF",
                            data=output_pdf,
                            file_name=f"highlighted_{uploaded_file.name}",
                            mime="application/pdf"
                        )
                    else:
                        st.warning("No matches found.")
                        
            except Exception as e:
                st.error(f"An error occurred: {e}")
