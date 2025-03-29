import streamlit as st
import os
import io
from docx import Document
from docx.oxml import OxmlElement # More robust way to copy content

# --- Configuration ---
OUTPUT_FILENAME = 'merged_document.docx'
# --- End Configuration ---

def merge_word_documents_from_streams(uploaded_files):
    """
    Merges content from uploaded Word files (passed as Streamlit UploadedFile objects)
    into a single Document object returned as a BytesIO stream.

    Args:
        uploaded_files (list): A list of Streamlit UploadedFile objects.

    Returns:
        io.BytesIO or None: A BytesIO stream containing the merged Word document,
                             or None if no files were processed or an error occurred.
        int: Count of files successfully processed.
        list: List of filenames that failed to process.
    """
    if not uploaded_files:
        return None, 0, []

    # Sort files by name for consistent merging order
    # Access the 'name' attribute of the UploadedFile object
    uploaded_files.sort(key=lambda x: x.name)

    merged_document = Document()
    files_processed_count = 0
    failed_files = []
    first_file = True # To avoid leading page break

    st.write(f"Attempting to merge {len(uploaded_files)} files...")
    progress_bar = st.progress(0)
    status_text = st.empty()

    for index, uploaded_file in enumerate(uploaded_files):
        filename = uploaded_file.name
        status_text.text(f"Processing file {index + 1}/{len(uploaded_files)}: {filename}")
        print(f"Processing file {index + 1}/{len(uploaded_files)}: {filename}") # Debug print

        # Add a page break before appending the new document's content (except for the first file)
        if not first_file:
            # Ensure we don't add a break if the previous file failed
            if filename not in failed_files:
               try:
                   merged_document.add_page_break()
               except Exception as page_break_err:
                   st.warning(f"Could not add page break before {filename}: {page_break_err}")

        else:
            first_file = False

        try:
            # Read the uploaded file's content into a BytesIO stream
            file_stream = io.BytesIO(uploaded_file.getvalue())
            file_stream.seek(0) # Go to the start of the stream

            # Open the source document from the stream
            sub_doc = Document(file_stream)

            # Append the content using the underlying XML structure
            for element in sub_doc.element.body:
                # Check if element is not None before appending
                if element is not None:
                    merged_document.element.body.append(element)
                else:
                    print(f"Warning: Found None element in body of {filename}") # Debug print

            files_processed_count += 1

        except Exception as e:
            st.warning(f"  ⚠️ Error processing '{filename}': {e}. Skipping this file.")
            print(f"  Error processing {filename}: {e}") # Debug print
            failed_files.append(filename)
            # Optionally add a note about the skip in the merged doc itself if needed
            # merged_document.add_paragraph(f"[Error processing file: {filename} - Skipped]")

        # Update progress bar
        progress_bar.progress((index + 1) / len(uploaded_files))

    status_text.text("Merging complete.")

    if files_processed_count == 0:
        st.error("No files were successfully processed.")
        return None, 0, failed_files

    # Save the merged document to a BytesIO stream
    try:
        output_stream = io.BytesIO()
        merged_document.save(output_stream)
        output_stream.seek(0) # Rewind the stream
        return output_stream, files_processed_count, failed_files
    except Exception as e:
        st.error(f"Error saving the final merged document to memory: {e}")
        print(f"Error saving merged document: {e}") # Debug print
        return None, files_processed_count, failed_files # Return count even if save fails

# --- Streamlit App UI ---
st.set_page_config(layout="wide")
st.title("📄 Word Document Merger")
st.markdown("Upload multiple `.docx` files below. They will be merged in alphabetical order by filename.")

uploaded_files = st.file_uploader(
    "Choose Word files (.docx)",
    type="docx",
    accept_multiple_files=True,
    help="Select all the Word documents you want to merge."
)

# Initialize session state for storing the result
if 'merged_doc_stream' not in st.session_state:
    st.session_state.merged_doc_stream = None
if 'files_processed_count' not in st.session_state:
    st.session_state.files_processed_count = 0
if 'failed_files' not in st.session_state:
    st.session_state.failed_files = []
if 'merge_attempted' not in st.session_state:
    st.session_state.merge_attempted = False


if uploaded_files:
    st.write(f"**{len(uploaded_files)}** file(s) selected.")

    if st.button("✨ Merge Selected Files", key="merge_button"):
        st.session_state.merge_attempted = True
        st.session_state.merged_doc_stream = None # Reset previous results
        st.session_state.files_processed_count = 0
        st.session_state.failed_files = []

        with st.spinner("Merging documents... This might take a while for many files."):
            merged_stream, processed_count, failures = merge_word_documents_from_streams(uploaded_files)

        st.session_state.merged_doc_stream = merged_stream
        st.session_state.files_processed_count = processed_count
        st.session_state.failed_files = failures

        if merged_stream:
            st.success(f"Successfully processed {processed_count} out of {len(uploaded_files)} files.")
        else:
            st.error("Merging process failed or produced no output.")

        if failures:
             st.warning(f"Could not process the following files: {', '.join(failures)}")


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
st.info("Note: Formatting from the original documents is preserved as much as possible, but complex layouts might vary slightly. Files are merged in alphabetical order.")
