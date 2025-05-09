import concurrent.futures # For ThreadPoolExecutor
import datetime
# import chardet # If needed by your functions
import streamlit as st
import easyocr
import torch
import numpy as np
import pandas as pd
from PIL import Image
from io import BytesIO
import os
import pdfplumber
from pptx import Presentation
from docx import Document
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeoutError
import re
import time
import logging
import threading # For thread ID

# --- Setup Logging ---
logging.basicConfig(
    level=logging.INFO, # DEBUG for more details
    format='%(asctime)s - %(process)d - %(threadName)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# =====================================================================================
# >>> PASTE ALL YOUR EXISTING TOP-LEVEL HELPER FUNCTIONS AND GLOBAL VARIABLES HERE <<<
#
# Example:
# torch.classes.__path__ = []
# @st.cache_resource
# def load_easyocr_model(): ...
# def extract_text_easyocr(pil_img): ...
# # ... ALL your other data processing functions and global regex/patterns ...
# # ... identify_document_type, classify_document, mask_data ...
# classification_options = { ... }
#
# !!! THEIR INTERNAL LOGIC REMAINS UNTOUCHED BY ME !!!
# =====================================================================================


# ---- NEW WORKER ORCHESTRATOR FUNCTION ----
# YOU WILL PUT YOUR EXISTING SINGLE-FILE PROCESSING LOGIC INSIDE THIS FUNCTION.
# This function will be run in a separate thread.
def worker_orchestrator_function(file_content_bytes, original_filename, file_mime_type, selected_classifications_list_from_ui):
    """
    This function takes the data for ONE file and calls YOUR EXISTING sequence 
    of helper functions to process it and produce a report_row.
    """
    main_pid = os.getpid()
    thread_id = threading.get_ident()
    logger.info(f"PID {main_pid}, ThreadID {thread_id}: Orchestrator started for '{original_filename}'")

    file_buffer = BytesIO(file_content_bytes)
    ext = os.path.splitext(original_filename)[1].lower()
    
    # Initialize variables that your sequence of functions will populate
    raw_text = ""
    cleaned_text = ""
    text_upper = ""
    identified_doc_type = "Unknown"
    details_extracted_dict = {} 
    masked_details_dict = {}
    doc_name_classified = "Unknown"
    doc_type_classified = "Unclassified"
    sensitivity_classified = "Low"
    is_pii_classified = False
    final_report_row = {}

    try:
        # --- YOUR EXISTING LOGIC FOR PROCESSING ONE FILE STARTS HERE ---
        #    (Copy and paste your sequential calls from your original synchronous 'Scan Files' button block)

        # Example Step 1: Text Extraction (Call YOUR functions)
        if ext in [".png", ".jpg", ".jpeg"]:
            pil_img = Image.open(file_buffer) # Add .convert("RGB") if your easyocr function needs it
            raw_text = extract_text_easyocr(pil_img)
        elif ext == ".pdf" or file_mime_type == "application/pdf":
            raw_text = extract_data_from_pdf(file_buffer)
        # ... (YOUR elif for .txt, .pptx, .docx, .xlsx, .csv) ...
        # IMPORTANT: If you handle .xlsx/.csv differently and return early, replicate that here.
        # For example, if CSV/Excel processing completes and returns a report_row directly:
        elif ext in [".xlsx", ".csv"] or file_mime_type in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "text/csv"]:
            df = process_excel_csv(file_buffer) 
            columns = df.columns.tolist()
            # result_classified_csv = classify_csv_columns(df) # If you used this for the report
            doc_name_classified_csv, doc_type_classified_csv, sensitivity_classified_csv = "Spreadsheet/CSV", "Structured Data", "High"
            details_str_csv = str({"columns": columns})
            is_pii_csv = True
            final_report_row = {
                "document_name": original_filename, "document": doc_name_classified_csv,
                "details": details_str_csv, "sensitivity": sensitivity_classified_csv,
                "document_type": doc_type_classified_csv, "is_pii": is_pii_csv
            }
            logger.info(f"PID {main_pid}, ThreadID {thread_id}: Orchestrator finished CSV/Excel '{original_filename}'")
            return {"status": "success", "filename": original_filename, "report_row": final_report_row, "error_message": None}
        else:
            raw_text = f"Error: Unsupported file type '{ext}'."

        # Example Step 2: Preprocessing (Call YOUR function)
        if raw_text and "Error:" in raw_text:
            cleaned_text = raw_text; text_upper = ""
        else:
            cleaned_text = preprocess_text(raw_text)
            text_upper = cleaned_text.upper() if cleaned_text else ""

        # Example Step 3: Document Identification (Call YOUR function)
        if text_upper: identified_doc_type = identify_document_type(text_upper)
        else: identified_doc_type = "Unknown"
        
        # Example Step 4: Specific Details Extraction (YOUR if/elif logic)
        details_extracted_dict = {} # Default
        if "Error:" not in cleaned_text and identified_doc_type in selected_classifications_list_from_ui:
            # PASTE YOUR if identified_doc_type == "Aadhaar": details_extracted_dict = extract_aadhaar_details(cleaned_text)
            # ELIF identified_doc_type == "PAN": details_extracted_dict = extract_pan_details(cleaned_text)
            # ... AND SO ON FOR ALL YOUR DOCUMENT TYPES ...
            # Ensure this block exactly matches your original synchronous logic.
            if identified_doc_type == "Aadhaar": details_extracted_dict = extract_aadhaar_details(cleaned_text)
            elif identified_doc_type == "PAN": details_extracted_dict = extract_pan_details(cleaned_text)
            elif identified_doc_type == "Voter ID": details_extracted_dict = extract_voter_details(cleaned_text)
            elif identified_doc_type == "Passport": details_extracted_dict = extract_passport_details(cleaned_text)
            elif identified_doc_type == "Driving License": details_extracted_dict = extract_driving_license_details(cleaned_text)
            elif identified_doc_type == "Medical Document": details_extracted_dict = extract_medical_details(cleaned_text)
            elif identified_doc_type == "Financial Document": details_extracted_dict = extract_financial_details(cleaned_text)
            else: details_extracted_dict = {"Note": f"No specific extractor for '{identified_doc_type}'."}
        elif "Error:" in cleaned_text: details_extracted_dict = {"ExtractionError": cleaned_text}
        else: details_extracted_dict = {"Note": f"Type '{identified_doc_type}' not selected or Unknown."}


        # Example Step 5: Mask Data (Call YOUR function)
        if details_extracted_dict and not details_extracted_dict.get("ExtractionError") and not details_extracted_dict.get("Note"):
            masked_details_dict = mask_data(details_extracted_dict)
        else:
            masked_details_dict = details_extracted_dict # Pass error/note

        # Example Step 6: Document Classification (Call YOUR function)
        # !!! CRITICAL: Ensure the first argument to classify_document matches YOUR function's expectation !!!
        # If it expects cleaned_text (as in your original provided main loop):
        doc_name_classified, doc_type_classified, sensitivity_classified = classify_document(cleaned_text, text_upper)
        # If it expects the dictionary of extracted details:
        # doc_name_classified, doc_type_classified, sensitivity_classified = classify_document(details_extracted_dict, text_upper)
        
        is_pii_classified = doc_name_classified not in ["Unknown", "Other"] and sensitivity_classified in ["Medium", "High"]

        # Construct the final report row
        final_report_row = {
            "document_name": original_filename,
            "document": doc_name_classified,
            "details": str(masked_details_dict if masked_details_dict else "No Data Extracted!"), # Your original default
            "sensitivity": sensitivity_classified,
            "document_type": doc_type_classified,
            "is_pii": is_pii_classified
        }
        # --- YOUR EXISTING LOGIC FOR PROCESSING ONE FILE ENDS HERE ---
        
        logger.info(f"PID {main_pid}, ThreadID {thread_id}: Orchestrator successfully processed '{original_filename}'")
        return {"status": "success", "filename": original_filename, "report_row": final_report_row, "error_message": None}

    except Exception as e:
        logger.error(f"PID {main_pid}, ThreadID {thread_id}: Unhandled orchestrator error for '{original_filename}': {str(e)}", exc_info=True)
        error_report_row = {
            "document_name": original_filename, "document": "Processing Error",
            "details": f"Critical error in orchestrator: {str(e)}", "sensitivity": "Unknown",
            "document_type": "Error", "is_pii": False 
        }
        return {"status": "error", "filename": original_filename, "report_row": error_report_row, "error_message": str(e)}


# ---- Helper function to manage Executor (USING ThreadPoolExecutor) ----
def get_app_executor():
    executor_key = "my_app_thread_executor_final_demo_v2" # Use a distinct key
    if executor_key not in st.session_state:
        cpu_cores = os.cpu_count()
        num_workers = 4 
        if cpu_cores and cpu_cores > 1:
            num_workers = min(max(4, cpu_cores * 2), 16) 
        
        logger.info(f"Main App (PID {os.getpid()}): Initializing ThreadPoolExecutor with {num_workers} workers.")
        print(f"PRINT DEBUG: Main App (PID {os.getpid()}): Initializing ThreadPoolExecutor with {num_workers} workers.")
        st.session_state[executor_key] = ThreadPoolExecutor(max_workers=num_workers)
    return st.session_state[executor_key]

# ---- Main Streamlit Application (Structure from before, uses ThreadPoolExecutor now) ----
def main():
    app_executor = get_app_executor() 

    st.set_page_config(page_title="Document Scanner & Classifier (Demo vFinal)", layout="wide")
    st.sidebar.title("🗂️ Document Intelligence")
    page = st.sidebar.radio("Navigation", ["File Scanner", "DB Scanner (Placeholder)"])

    tasks_session_key = "file_tasks_status_demo_v_final"
    report_session_key = "report_rows_collected_demo_v_final"

    if tasks_session_key not in st.session_state:
        st.session_state[tasks_session_key] = {}
    if report_session_key not in st.session_state:
        st.session_state[report_session_key] = []
    
    if page == "File Scanner":
        st.header("🔍 File Scanner & Document Classifier")
        
        with st.form(key="file_upload_scan_form_demo_final"): # New key for form
            uploaded_files = st.file_uploader(
                "Choose files for scanning",
                type=["png", "jpg", "jpeg", "pdf", "txt", "xlsx", "docx", "csv", "pptx"],
                accept_multiple_files=True,
                help="Upload files. Processing will be non-blocking."
            )
            selected_classifications_list_from_ui = st.multiselect( 
                "Target document types for detailed field extraction:",
                options=list(classification_options.keys()), 
                default=list(classification_options.keys())
            )
            scan_button = st.form_submit_button(label="🚀 Scan Uploaded Files", use_container_width=True)

        if scan_button:
            if uploaded_files:
                queued_count = 0
                for up_file in uploaded_files:
                    file_id = getattr(up_file, 'file_id', f"{up_file.name}_{up_file.size}")
                    current_task = st.session_state[tasks_session_key].get(file_id)
                    if not current_task or current_task["status"] == "error":
                        logger.info(f"Main App: Queuing '{up_file.name}' (ID: {file_id}) with ThreadPoolExecutor.")
                        try:
                            file_bytes = up_file.getvalue()
                            future = app_executor.submit( # Submitting to the orchestrator
                                worker_orchestrator_function, # <<< NEW ORCHESTRATOR FUNCTION
                                file_bytes, up_file.name,
                                up_file.type, selected_classifications_list_from_ui
                            )
                            st.session_state[tasks_session_key][file_id] = {
                                "filename": up_file.name, "status": "queued", "future": future, 
                                "result": None, "submitted_at": datetime.datetime.now().isoformat()
                            }
                            queued_count += 1
                        except Exception as e:
                            logger.error(f"Main App: Error submitting '{up_file.name}': {e}", exc_info=True)
                            st.session_state[tasks_session_key][file_id] = {
                                "filename": up_file.name, "status": "error", "future": None, 
                                "result": {"error_message": f"Submission failed: {str(e)}"},
                                "submitted_at": datetime.datetime.now().isoformat()
                            }
                if queued_count > 0:
                    st.toast(f"{queued_count} file(s) sent for background processing.", icon="⏳")
                    st.rerun()
                elif not uploaded_files: st.warning("Please upload files.")
                else: st.info("Selected files are already processing or completed.")
            else:
                st.warning("No files uploaded to scan.")

        # --- Display Progress and Results (This UI logic should be mostly fine) ---
        active_tasks_running = False
        if st.session_state[tasks_session_key]:
            st.write("---"); st.subheader("📊 Processing Dashboard")
            sorted_ids = sorted(st.session_state[tasks_session_key].keys(), 
                                key=lambda k: st.session_state[tasks_session_key][k]["submitted_at"])

            for file_id in sorted_ids:
                task = st.session_state[tasks_session_key][file_id]
                if task["status"] == "queued" and task.get("future"): task["status"] = "processing"

                expander_open = task["status"] in ["queued", "processing", "error"]
                with st.expander(f"📁 {task['filename']}  |  Status: {task['status'].upper()}", expanded=expander_open):
                    st.caption(f"Submitted: {task['submitted_at']}")
                    if task["status"] == "processing":
                        active_tasks_running = True; st.progress(50) 
                        if task["future"] and task["future"].done():
                            needs_rerun = False
                            try:
                                worker_result_dict = task["future"].result(timeout=0.1) # result from orchestrator
                                task["result"] = worker_result_dict 
                                if worker_result_dict.get("status") == "success":
                                    task["status"] = "completed"
                                    if worker_result_dict.get("report_row") and not any(
                                        r["document_name"] == worker_result_dict["filename"] for r in st.session_state[report_session_key]):
                                        st.session_state[report_session_key].append(worker_result_dict["report_row"])
                                    st.success(f"'{task['filename']}' processed.")
                                else: 
                                    task["status"] = "error"
                                    if worker_result_dict.get("report_row") and not any(
                                        r["document_name"] == worker_result_dict["filename"] for r in st.session_state[report_session_key]):
                                        st.session_state[report_session_key].append(worker_result_dict["report_row"])
                                    st.warning(f"Issue processing '{task['filename']}'.")
                                needs_rerun = True
                            except FutureTimeoutError: pass 
                            except Exception as e: 
                                logger.error(f"Main App: Failed to get result for '{task['filename']}': {e}", exc_info=True)
                                task["status"] = "error"
                                task["result"] = {"error_message": f"System error retrieving result: {str(e)}"}
                                err_row = {"document_name": task['filename'], "document": "System Error", "details": str(e), "sensitivity": "Unknown", "document_type": "Error", "is_pii": False}
                                if not any(r["document_name"] == task['filename'] for r in st.session_state[report_session_key]):
                                    st.session_state[report_session_key].append(err_row)
                                needs_rerun = True
                            if needs_rerun: st.rerun()
                    
                    if task["status"] == "completed": st.success("✅ Processing complete.")
                    elif task["status"] == "error":
                        st.error("❌ Error during processing:")
                        err_msg_display = "Unknown error."
                        if task.get("result"):
                            if task["result"].get("error_message"): err_msg_display = task["result"]["error_message"]
                            elif task["result"].get("report_row") and "details" in task["result"]["report_row"]: err_msg_display = task["result"]["report_row"]["details"]
                        st.text_area("Error Details:", value=err_msg_display, height=75, disabled=True, key=f"error_display_final_{file_id}") # Unique key
                    elif task["status"] == "queued": st.info("⏳ In queue...")

            if st.session_state[report_session_key]:
                st.write("---"); st.subheader("📋 Consolidated Report")
                df_report = pd.DataFrame(st.session_state[report_session_key])
                st.dataframe(df_report, use_container_width=True)
                pii_count = sum(1 for row in st.session_state[report_session_key] if row.get("is_pii", False))
                st.metric("Files in Report with PII", pii_count)
                try:
                    excel_bytes = save_to_excel(st.session_state[report_session_key])
                    st.download_button("📥 Download Report (Excel)", excel_bytes, "ScanReport.xlsx", 
                                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                       key="dl_report_final_v2", use_container_width=True) # New key
                except Exception as e: st.error(f"Excel generation failed: {e}")
            
            if active_tasks_running:
                time.sleep(3) 
                logger.debug("Main App: Active tasks running, rerunning for UI poll.")
                st.rerun()

        st.write("---")
        if st.button("🧹 Clear All Data & Reset", key="clear_all_final_demo_v2_btn", use_container_width=True): #New key
            st.session_state[tasks_session_key] = {}
            st.session_state[report_session_key] = []
            if "file_uploader_widget" in st.session_state: st.session_state.file_uploader_widget = []
            st.toast("Cleared all data.", icon="♻️"); st.rerun()

    elif page == "DB Scanner (Placeholder)":
        st.header("🗄️ Database Scanner")
        st.warning("This feature is under development.", icon="⚠️")


# CRITICAL: This guard is ESSENTIAL for any multiprocessing/threading to behave well,
# especially on Windows when scripts might be re-imported by worker initializers.
if __name__ == "__main__":
    main()