
import docx
import io
import os
import translation_lib as tl
from processor import process_financial_report

def repro():
    docx_path = "number.docx"
    if not os.path.exists(docx_path):
        print(f"Error: {docx_path} not found.")
        return

    # Metadata for processing
    metadata = {
        "Name (not capitalized)": "TEST COMPANY",
        "Reporting date": "31/12/2025",
        "Translate into": "E",
        "Year-end date": "31/12/2025",
        "Translated Name": "Test Company",
        "Period (in table)": "From 01/01/2025 to 31/12/2025",
        "signer_1": "Signer One",
        "signer_2": "Signer Two",
        "signer_3": "Signer Three"
    }
    
    # Mock translation map (not empty to ensure replace_text_in_document is called)
    translation_map = {"DUMMY_KEY": "DUMMY_VAL"}
    
    # Process settings (only number swap and date format for focus)
    process_settings = {
        "unicode": True,
        "clean_v": True,
        "para_template": True,
        "dictionary": True,
        "dual_font": False,
        "number_swap": True,
        "table_size": False,
        "date_format": True,
        "textbox": False,
        "signer_accents": False,
        "highlight": False,
        "suggestion": False
    }

    with open(docx_path, "rb") as f:
        processed_file_stream, msg = process_financial_report(
            f, 
            metadata=metadata, 
            translation_map=translation_map,
            case_threshold=30,
            target_col="E",
            process_settings=process_settings
        )

    if processed_file_stream:
        output_path = "scratch/number_output.docx"
        with open(output_path, "wb") as f:
            f.write(processed_file_stream.getvalue())
        print(f"Processed file saved to {output_path}")
        
        # Now inspect output
        print("\n--- Inspecting Processed Output ---")
        doc = docx.Document(output_path)
        for i, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    text = cell.text
                    if any(char.isdigit() for char in text):
                        # Find if any dots or commas exist in numbers
                        # VN style: 1.234,56 -> should become 1,234.56
                        if '.' in text or ',' in text:
                            print(f"Cell({r_idx}, {c_idx}): '{text}'")
    else:
        print(f"Processing failed: {msg}")

if __name__ == "__main__":
    repro()
