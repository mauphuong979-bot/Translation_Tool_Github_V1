import pandas as pd
from docx import Document
import io
import translation_lib as tl

def process_financial_report(file_stream, metadata=None, translation_map=None, case_threshold=30, target_col="E", process_settings=None, py_dict=None, review_items=None):
    """
    Processes the financial report by applying metadata and global translations.
    """
    try:
        # Load the document from the uploaded file stream
        doc = Document(file_stream)
        
        # 1. Apply Metadata (if provided)
        # Note: We can translate metadata values using the translation_map if needed,
        # but here we assume the user provides the final translated values in metadata dict.
        if metadata:
            # We add a summary info section at the beginning as before
            # Or we could just replace placeholders if any. 
            # For now, let's keep the metadata summary as a demonstration.
            # (In a real scenario, we might want to inject these into specific locations)
            pass

        # 2. Apply Global Translations from the (potentially edited) dictionary
        if translation_map or (process_settings or {}).get("review_annotations", False):
            tl.replace_text_in_document(
                doc, 
                translation_map or {}, 
                case_threshold, 
                target_col=target_col, 
                metadata=metadata,
                process_settings=process_settings,
                py_dict=py_dict,
                review_items=review_items
            )
        
        if review_items:
            from review_roundtrip import bind_review_manifest
            bind_review_manifest(doc, review_items, target_col)

        # Save to a BytesIO object
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        
        return output, "Report processed."
    except Exception as e:
        return None, f"❌ Error during processing: {str(e)}"
