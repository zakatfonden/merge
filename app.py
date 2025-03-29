import streamlit as st
import os
import io
import pandas as pd # Import pandas
from docx import Document
from docx.oxml import OxmlElement # More robust way to copy content
from natsort import natsorted, ns # Import natsorted and ns for algorithms

# --- Configuration ---
OUTPUT_FILENAME = 'merged_document.docx'
# --- End Configuration ---

def merge_word_documents_from_streams(uploaded_files):
    """
    Merges content from uploaded Word files (passed as Streamlit UploadedFile objects)
    into a single Document object returned as a BytesIO stream, using natural sort order.

    Args:
        uploaded_files (list): A list of Streamlit UploadedFile objects (should be pre-sorted).

    Returns:
        io.BytesIO or None: A BytesIO stream containing the merged Word document,
                            or None if no files were processed or an error occurred.
        int: Count of files successfully processed.
        list: List of filenames that failed to process.
    """
    if not uploaded_files:
        # This case should ideally be caught before calling the function if sorting happened
        return None, 0, []

    # Note: Files are assumed to be pre-sorted by the calling UI logic using natsort

    merged_document = Document()
    files_processed_count = 0
    failed_files = []
    first_file = True # To avoid leading page break

    st.write(f"Processing {len(uploaded_files)} files...") # Simple status message
    # Display the confirmed order again for safety check during processing
    st.caption(f"Confirmed merge order: {', '.join([f.name for f in uploaded_files])}")

    progress_bar = st.progress(0)
    status_text = st.empty()

    for index, uploaded_file in enumerate(uploaded_files):
        filename = uploaded_file.name
        status_text.text(f"Processing file {index + 1}/{len(uploaded_files)}: {filename}")
        print(f"Processing file {index + 1}/{len(uploaded_files)}: {filename}") # Debug print

        # Add a page break before appending the new document's content (except for the first file)
        if not first_file:
            if filename not in failed_files:
                try:
                    merged_document.add_page_break()
                except Exception as page_break_err:
                    st.warning(f"Could not add page break before {filename}: {page_break_err}")
        else:
            first_file = False

        try:
            file_stream = io.BytesIO(uploaded_file.getvalue())
            file_stream.seek(0)
            sub_doc = Document(file_stream)
            for element in sub_doc.element.body:
                if element is not None:
                    merged_document.element.body.append(element)
                else:
                    print(f"Warning: Found None element in body of {filename}")
            files_processed_count += 1
        except Exception as e:
            st.warning(f"  ⚠️ Error processing '{filename}': {e}. Skipping this file.")
            print(f"  Error processing {filename}: {e}")
            failed_files.append(filename)

        progress_bar.progress((index + 1) / len(uploaded_files))

    status_text.text("Merging complete.")

    if files_processed_count == 0:
        st.error("No files were successfully processed.")
        return None, 0, failed_files

    try:
        output_stream = io.BytesIO()
        merged_document.save(output_stream)
        output_stream.seek(0)
        return output_stream, files_processed_count, failed_files
    except Exception as e:
        st.error(f"Error saving the final merged document to memory: {e}")
        print(f"Error saving merged document: {e}")
        return None, files_processed_count, failed_files

# --- Streamlit App UI ---
st.set_page_config(layout="wide")
st.title("📄 Word Document Merger")
st.markdown("Upload multiple `.docx` files below. They will be merged using **natural sort order** (e.g., `file2.docx` before `file10.docx`) based on filename.")

uploaded_files_list = st.file_uploader( # Renamed variable for clarity
    "Choose Word files (.docx)",
    type="docx",
    accept_multiple_files=True,
    help="Select all the Word documents you want to merge.",
    key="file_uploader" # Added a key for stability
)

# Initialize session state (only needed if logic depends on previous runs, but good practice)
if 'merged_doc_stream' not in st.session_state:
    st.session_state.merged_doc_stream = None
if 'files_processed_count' not in st.session_state:
    st.session_state.files_processed_count = 0
if 'failed_files' not in st.session_state:
    st.session_state.failed_files = []
if 'merge_attempted' not in st.session_state:
    st.session_state.merge_attempted = False
if 'sorted_files_for_merge' not in st.session_state: # Store the sorted list
    st.session_state.sorted_files_for_merge = []

# --- Display Uploaded Files Section ---
if uploaded_files_list:
    # Sort the files immediately for display using natural sort order
    sorted_files_display = natsorted(uploaded_files_list, key=lambda x: x.name, alg=ns.IGNORECASE)
    st.session_state.sorted_files_for_merge = sorted_files_display # Store for the merge button

    st.subheader(f"Files Selected ({len(sorted_files_display)}):")
    st.caption("Files will be merged in the order shown below.")

    # Create a pandas DataFrame for better display
    df_files = pd.DataFrame({
        'Order': range(1, len(sorted_files_display) + 1),
        'Filename': [f.name for f in sorted_files_display]
    })

    # Display the DataFrame - make it scrollable by setting height
    st.dataframe(
        df_files,
        hide_index=True, # Hide the default pandas index
        use_container_width=True, # Use available width
        height=300 # Set a fixed height (pixels) - adjust as needed
        )

    # --- Merge Button ---
    if st.button("✨ Merge Selected Files", key="merge_button"):
        st.session_state.merge_attempted = True
        st.session_state.merged_doc_stream = None # Reset previous results
        st.session_state.files_processed_count = 0
        st.session_state.failed_files = []

        # Use the pre-sorted list stored in session state
        files_to_merge = st.session_state.sorted_files_for_merge

        if not files_to_merge:
             st.warning("No files found in the sorted list state. Please re-upload.")
        else:
            with st.spinner("Merging documents... This might take a while for many files."):
                # Pass the already sorted list to the function
                merged_stream, processed_count, failures = merge_word_documents_from_streams(files_to_merge)

            st.session_state.merged_doc_stream = merged_stream
            st.session_state.files_processed_count = processed_count
            st.session_state.failed_files = failures

            if merged_stream:
                st.success(f"Successfully processed {processed_count} out of {len(files_to_merge)} files.")
            else:
                st.error("Merging process failed or produced no output.")

            if failures:
                st.warning(f"Could not process the following files: {', '.join(failures)}")

# --- Download Button Section ---
# Show download button only if merging was successful and attempted
if st.session_state.merge_attempted and st.session_state.merged_doc_stream:
    st.download_button(
        label=f"📥 Download Merged File ({st.session_state.files_processed_count} files)",
        data=st.session_state.merged_doc_stream,
        file_name=OUTPUT_FILENAME,
        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        key="download_button"
    )
    st.caption(f"Filename: `{OUTPUT_FILENAME}`")
elif st.session_state.merge_attempted and not st.session_state.merged_doc_stream:
    st.error("Merging failed. Cannot provide download link.")

st.markdown("---")
st.info("Note: Formatting from the original documents is preserved as much as possible, but complex layouts might vary slightly. Files are merged in natural sort order based on filename.")
