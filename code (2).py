# =====================================================================================
# >>> PASTE ALL YOUR EXISTING TOP-LEVEL HELPER FUNCTIONS AND GLOBAL VARIABLES HERE <<<
#
# Example:
# import ... (all your necessary imports for these functions)
# torch.classes.__path__ = []
# @st.cache_resource
# def load_easyocr_model(): ...
# def extract_text_easyocr(pil_img): ...
# # ... ALL other data processing functions and global regex/patterns ...
# classification_options = { ... }
#
# !!! IMPORTANT: THE LOGIC INSIDE THESE FUNCTIONS IS YOURS AND REMAINS UNTOUCHED !!!
# =====================================================================================


# ---- NEW WORKER FUNCTION for Background Processing ----
# This function calls YOUR existing sequence of functions for a single file.
def worker_process_file(file_content_bytes, original_filename, file_mime_type, selected_classifications_list_from_ui):
    """
    Orchestrates the processing of a single file using YOUR existing helper functions.
    This function is run in a separate process.
    """
    pid = os.getpid()
    logger.info(f"PID {pid}: Worker processing '{original_filename}'")

    file_buffer = BytesIO(file_content_bytes)
    ext = os.path.splitext(original_filename)[1].lower()
    
    # Variables to hold results from YOUR functions, mimicking your synchronous flow
    raw_text = "" # From your text extraction stage
    cleaned_text = "" # From your preprocess_text
    text_upper = "" # Uppercase version of cleaned_text
    identified_doc_type = "Unknown" # From your identify_document_type
    details_extracted_by_your_logic = {} # From your extract_..._details if/elif chain
    masked_details_final = {} # From your mask_data
    # Final classification results from your classify_document
    doc_name_classified, doc_type_classified, sensitivity_classified = "Unknown", "Unclassified", "Low"
    is_pii_classified = False
    report_row_for_this_file = {}

    try:
        # --- STAGE 1: Text Extraction (using YOUR functions) ---
        if ext in [".png", ".jpg", ".jpeg"]:
            pil_img = Image.open(file_buffer) # .convert("RGB") if your easyocr function needs it
            raw_text = extract_text_easyocr(pil_img)
        elif ext == ".pdf" or file_mime_type == "application/pdf":
            raw_text = extract_data_from_pdf(file_buffer)
        # ... (Add elif for .txt, .pptx, .docx as in your original logic, calling YOUR functions) ...
        elif ext in [".txt", ".py", ".java", ".c", ".cpp", ".js", ".json", ".xml"]: # From your original code
            raw_text = extract_text_plain(file_buffer)
        elif ext == ".pptx" or file_mime_type == "application/vnd.openxmlformats-officedocument.presentationml.presentation": # From your original code
            raw_text = extract_text_from_pptx(file_buffer)
        elif ext == ".docx" or file_mime_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document": # From your original code
            raw_text = extract_text_from_docx(file_buffer)
        elif ext in [".xlsx", ".csv"] or file_mime_type in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "text/csv"]:
            # This is how your original code handled CSV/Excel directly for the report
            df = process_excel_csv(file_buffer) # YOUR function
            columns = df.columns.tolist()
            # result_classified_csv = classify_csv_columns(df) # YOUR function if used for report

            # Construct report row for CSV/Excel as per your original logic
            doc_name_final, doc_type_final, sensitivity_final = "Spreadsheet/CSV", "Structured Data", "High" # As per your code
            details_for_report_str = str({"columns": columns}) # As per your code
            is_pii_final = True # As per your code logic for CSV/Excel

            report_row_for_this_file = {
                "document_name": original_filename,
                "document": doc_name_final,
                "details": details_for_report_str,
                "sensitivity": sensitivity_final,
                "document_type": doc_type_final,
                "is_pii": is_pii_final # If you track this
            }
            logger.info(f"PID {pid}: Worker finished CSV/Excel '{original_filename}'")
            return {"status": "success", "filename": original_filename, "report_row": report_row_for_this_file, "error_message": None}
        else:
            raw_text = f"Error: Unsupported file type '{ext}' for direct text extraction." # Your original handling
            logger.warning(f"PID {pid}: {raw_text} for '{original_filename}'")

        # --- STAGE 2: Preprocessing (using YOUR function) ---
        # This part assumes raw_text might be an error string from extraction.
        if raw_text and "Error:" in raw_text:
            cleaned_text = raw_text # Keep error message
            text_upper = "" # No valid text for further processing
        else:
            cleaned_text = preprocess_text(raw_text) # YOUR function
            text_upper = cleaned_text.upper() if cleaned_text else ""

        # --- STAGE 3: Document Identification (using YOUR function) ---
        if text_upper: # Only if text is valid
            identified_doc_type = identify_document_type(text_upper) # YOUR function
        else:
            identified_doc_type = "Unknown" # Or handle based on extraction error

        # --- STAGE 4: Specific Details Extraction (YOUR if/elif logic EXACTLY as in your synchronous code) ---
        # This is the critical part that must match your original File Scanner button logic.
        # It populates `details_extracted_by_your_logic`.
        # It must respect `selected_classifications_list_from_ui`.
        
        # Default if no specific extraction happened (as per your original implied logic)
        details_extracted_by_your_logic = {} # Start with an empty dict

        if "Error:" not in cleaned_text and identified_doc_type in selected_classifications_list_from_ui:
            # YOUR if/elif chain from your original synchronous code goes here:
            if identified_doc_type == "Aadhaar":
                details_extracted_by_your_logic = extract_aadhaar_details(cleaned_text)
            elif identified_doc_type == "PAN":
                details_extracted_by_your_logic = extract_pan_details(cleaned_text)
            elif identified_doc_type == "Voter ID":
                details_extracted_by_your_logic = extract_voter_details(cleaned_text)
            elif identified_doc_type == "Passport":
                details_extracted_by_your_logic = extract_passport_details(cleaned_text)
            elif identified_doc_type == "Driving License":
                details_extracted_by_your_logic = extract_driving_license_details(cleaned_text)
            elif identified_doc_type == "Medical Document":
                details_extracted_by_your_logic = extract_medical_details(cleaned_text)
            elif identified_doc_type == "Financial Document":
                details_extracted_by_your_logic = extract_financial_details(cleaned_text)
            else: # If identified_doc_type was something else unexpected but in selected_classifications
                details_extracted_by_your_logic = {"Note": f"No specific extractor for selected type '{identified_doc_type}'."}
        elif "Error:" in cleaned_text:
            details_extracted_by_your_logic = {"ExtractionError": cleaned_text}
        else: # Type not selected or Unknown
             details_extracted_by_your_logic = {"Note": f"Type '{identified_doc_type}' not selected or is Unknown."}


        # --- STAGE 5: Mask Data (using YOUR function) ---
        # Your mask_data function should operate on `details_extracted_by_your_logic`
        if details_extracted_by_your_logic and not details_extracted_by_your_logic.get("ExtractionError") and not details_extracted_by_your_logic.get("Note"):
            masked_details_final = mask_data(details_extracted_by_your_logic)
        else:
            masked_details_final = details_extracted_by_your_logic # Pass along note or error

        # --- STAGE 6: Document Classification (using YOUR function) ---
        # The first argument here MUST MATCH what your `classify_document` function expects
        # based on your original synchronous code (likely `cleaned_text` or `details_extracted_by_your_logic`).
        # Your original code was: `doc_name, doc_type, sensitivity = classify_document(cleaned_text, text_upper)`
        # So we will use that signature here:
        doc_name_classified, doc_type_classified, sensitivity_classified = classify_document(cleaned_text, text_upper)
        
        is_pii_classified = doc_name_classified not in ["Unknown", "Other"] # As per your original logic

        # --- Construct the final report row for this file ---
        # The "details" field should be the string representation of `masked_details_final`
        # as per your original logic: `details": str(masked_details)`
        report_row_for_this_file = {
            "document_name": original_filename,
            "document": doc_name_classified,
            "details": str(masked_details_final if masked_details_final else "No Data Extracted!"), # Your original default
            "sensitivity": sensitivity_classified,
            "document_type": doc_type_classified,
            "is_pii": is_pii_classified # If you track this
        }
        logger.info(f"PID {pid}: Worker successfully processed '{original_filename}'")
        return {"status": "success", "filename": original_filename, "report_row": report_row_for_this_file, "error_message": None}

    except Exception as e:
        logger.error(f"PID {pid}: Unhandled worker error for '{original_filename}': {str(e)}", exc_info=True)
        # Create an error report row
        error_report_row = {
            "document_name": original_filename, "document": "Processing Error",
            "details": f"Critical error in worker: {str(e)}", "sensitivity": "Unknown",
            "document_type": "Error", "is_pii": False 
        }
        return {"status": "error", "filename": original_filename, "report_row": error_report_row, "error_message": str(e)}


# ---- Helper function to manage ProcessPoolExecutor (UNCHANGED from previous good version) ----
def get_app_executor():
    executor_key = "my_app_executor_v2_final" # Use a unique key for this app
    if executor_key not in st.session_state:
        cpu_cores = os.cpu_count()
        num_workers = 1 
        if cpu_cores and cpu_cores > 1:
            num_workers = max(1, cpu_cores // 2) 
            if num_workers > 4: num_workers = 4 
        logger.info(f"Main App (PID {os.getpid()}): Initializing ProcessPoolExecutor with {num_workers} workers.")
        print(f"PRINT DEBUG: Main App (PID {os.getpid()}): Initializing ProcessPoolExecutor with {num_workers} workers.")
        st.session_state[executor_key] = ProcessPoolExecutor(max_workers=num_workers)
    return st.session_state[executor_key]

# ---- Main Streamlit Application (UNCHANGED from previous good version for UI and task submission) ----
def main():
    app_executor = get_app_executor() # Ensures executor is ready

    st.set_page_config(page_title="Document Scanner & Classifier", layout="wide")
    st.sidebar.title("🗂️ Document Intelligence")
    page = st.sidebar.radio("Navigation", ["File Scanner", "DB Scanner (Placeholder)"])

    # Use distinct session state keys to avoid conflicts if rerunning old versions
    tasks_session_key = "file_tasks_status_final_v1"
    report_session_key = "report_rows_collected_final_v1"

    if tasks_session_key not in st.session_state:
        st.session_state[tasks_session_key] = {}
    if report_session_key not in st.session_state:
        st.session_state[report_session_key] = []
    
    if page == "File Scanner":
        st.header("🔍 File Scanner & Document Classifier")
        
        with st.form(key="file_upload_scan_form_final"):
            uploaded_files = st.file_uploader(
                "Choose files for scanning",
                type=["png", "jpg", "jpeg", "pdf", "txt", "xlsx", "docx", "csv", "pptx"],
                accept_multiple_files=True,
                help="Upload one or more files. Processing will happen in the background."
            )
            # Renamed variable to be clear it's from UI
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
                        logger.info(f"Main App: Queuing '{up_file.name}' (ID: {file_id}).")
                        try:
                            file_bytes = up_file.getvalue()
                            future = app_executor.submit(
                                worker_process_file, # Our top-level worker
                                file_bytes,
                                up_file.name,
                                up_file.type,
                                selected_classifications_list_from_ui # Pass the list
                            )
                            st.session_state[tasks_session_key][file_id] = {
                                "filename": up_file.name, "status": "queued", 
                                "future": future, "result": None,
                                "submitted_at": datetime.datetime.now().isoformat()
                            }
                            queued_count += 1
                        except Exception as e:
                            logger.error(f"Main App: Error submitting '{up_file.name}': {e}", exc_info=True)
                            st.session_state[tasks_session_key][file_id] = {
                                "filename": up_file.name, "status": "error", 
                                "future": None, "result": {"error_message": f"Submission failed: {str(e)}"},
                                "submitted_at": datetime.datetime.now().isoformat()
                            }
                if queued_count > 0:
                    st.toast(f"{queued_count} file(s) sent for background processing.", icon="⏳")
                    st.rerun()
                elif not uploaded_files: st.warning("Please upload files.")
                else: st.info("Selected files are already processing or completed. Clear results to rescan.")
            else:
                st.warning("No files uploaded to scan.")

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
                    st.caption(f"Submitted: {task['submitted_at']}") # Display submission time
                    if task["status"] == "processing":
                        active_tasks_running = True; st.progress(50) # Indeterminate
                        if task["future"] and task["future"].done():
                            needs_rerun_after_check = False
                            try:
                                worker_result = task["future"].result(timeout=0.1)
                                task["result"] = worker_result
                                if worker_result.get("status") == "success":
                                    task["status"] = "completed"
                                    if worker_result.get("report_row") and not any(
                                        r["document_name"] == worker_result["filename"] for r in st.session_state[report_session_key]):
                                        st.session_state[report_session_key].append(worker_result["report_row"])
                                    st.success(f"'{task['filename']}' processed.") # Changed from toast for clarity inside expander
                                else: 
                                    task["status"] = "error"
                                    if worker_result.get("report_row") and not any(
                                        r["document_name"] == worker_result["filename"] for r in st.session_state[report_session_key]):
                                        st.session_state[report_session_key].append(worker_result["report_row"])
                                    st.warning(f"Issue processing '{task['filename']}'.") # Changed from toast
                                needs_rerun_after_check = True
                            except FutureTimeoutError: pass 
                            except Exception as e: 
                                logger.error(f"Main App: Failed to get result for '{task['filename']}': {e}", exc_info=True)
                                task["status"] = "error"
                                task["result"] = {"error_message": f"System error retrieving result: {str(e)}"}
                                err_row = {"document_name": task['filename'], "document": "System Error", "details": str(e), "sensitivity": "Unknown", "document_type": "Error", "is_pii": False}
                                if not any(r["document_name"] == task['filename'] for r in st.session_state[report_session_key]):
                                    st.session_state[report_session_key].append(err_row)
                                needs_rerun_after_check = True
                            if needs_rerun_after_check: st.rerun()
                    
                    if task["status"] == "completed": st.success("✅ Processing complete.")
                    elif task["status"] == "error":
                        st.error("❌ Error during processing:")
                        err_msg_display = "Unknown error."
                        if task.get("result"): # Check if result dict exists
                            if task["result"].get("error_message"): # Error from worker or system
                                err_msg_display = task["result"]["error_message"]
                            elif task["result"].get("report_row") and "details" in task["result"]["report_row"]: # Error might be in details from worker
                                err_msg_display = task["result"]["report_row"]["details"]
                        st.text_area("Error Details:", value=err_msg_display, height=75, disabled=True, key=f"error_display_{file_id}")
                    elif task["status"] == "queued": st.info("⏳ In queue for processing...")

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
                                       key="dl_consolidated_report", use_container_width=True)
                except Exception as e: st.error(f"Excel generation failed: {e}")
            
            if active_tasks_running:
                time.sleep(3) 
                logger.debug("Main App: Active tasks running, rerunning for UI poll.")
                st.rerun()

        st.write("---")
        if st.button("🧹 Clear All Data & Reset", key="clear_all_final_btn", use_container_width=True):
            st.session_state[tasks_session_key] = {}
            st.session_state[report_session_key] = []
            if "file_uploader_widget" in st.session_state: st.session_state.file_uploader_widget = [] # Attempt to reset uploader
            st.toast("Cleared all data.", icon="♻️"); st.rerun()

    elif page == "DB Scanner (Placeholder)":
        st.header("🗄️ Database Scanner")
        st.warning("This feature is under development.", icon="⚠️")


# CRITICAL: This guard is ESSENTIAL for ProcessPoolExecutor, especially on Windows.
if __name__ == "__main__":
    # No need to call multiprocessing.set_start_method('spawn') on Windows,
    # as 'spawn' is the default and often the only available method.
    # The `if __name__ == "__main__":` guard itself is the most important part.
    main()