import pandas as pd
from docx import Document
import io
import translation_lib as tl

def process_financial_report(file_stream, metadata=None, translation_map=None, case_threshold=30, target_col="E", process_settings=None):
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
        if translation_map:
            tl.replace_text_in_document(
                doc, 
                translation_map, 
                case_threshold, 
                target_col=target_col, 
                metadata=metadata,
                process_settings=process_settings
            )
        
        # Save to a BytesIO object
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        
        return output, "Report processed."
    except Exception as e:
        return None, f"❌ Error during processing: {str(e)}"
