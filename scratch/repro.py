import sys
import os
import traceback

# Reconfigure stdout to use UTF-8
sys.stdout.reconfigure(encoding='utf-8')

# Add project directory to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import translation_lib as tl
from processor import process_financial_report
import metadata_extractor as mex

def main():
    docx_path = "../fs_AHVN_310326_V.docx"
    if not os.path.exists(docx_path):
        docx_path = "fs_AHVN_310326_V.docx"
    print(f"Using docx path: {docx_path}")
    
    with open(docx_path, "rb") as f:
        # Extract metadata
        ext_meta = mex.extract_metadata(f)
        f.seek(0)
        
        metadata_for_tags = {
            "name_vn": ext_meta.get("name_vn", "ABC Company"),
            "name_trans": ext_meta.get("name_trans", "CustomerName"),
            "year_end": ext_meta.get("year_end", ""),
            "report_date": ext_meta.get("report_date", ""),
            "period_in": ext_meta.get("period_in", ""),
            "period_in_2": ext_meta.get("period_in_2", ""),
            "signer_1": ext_meta.get("signer_1", ""),
            "signer_2": ext_meta.get("signer_2", ""),
            "signer_3": ext_meta.get("signer_3", "")
        }
        print("Extracted metadata:", metadata_for_tags)
        
        target_col = "E"
        v3_df = tl.load_and_fill_v3_dictionary(metadata_for_tags)
        translation_map = dict(zip(v3_df['Vietnamese'], v3_df[target_col]))
        
        metadata = {
            "Name (not capitalized)": tl.clean_text(metadata_for_tags["name_vn"]),
            "Reporting date": tl.clean_text(metadata_for_tags["report_date"]),
            "Translate into": target_col,
            "Year-end date": tl.clean_text(metadata_for_tags["year_end"]),
            "Translated Name": tl.clean_text(metadata_for_tags["name_trans"]),
            "Period (in table)": tl.clean_text(metadata_for_tags["period_in"]),
            "signer_1": tl.clean_text(metadata_for_tags["signer_1"]),
            "signer_2": tl.clean_text(metadata_for_tags["signer_2"]),
            "signer_3": tl.clean_text(metadata_for_tags["signer_3"])
        }
        
        process_steps = {
            "unicode": True,
            "clean_v": True,
            "para_template": True,
            "dictionary": True,
            "dual_font": True,
            "number_swap": True,
            "table_size": True,
            "date_format": True,
            "textbox": True,
            "signer_accents": True,
            "highlight": True,
            "suggestion": True
        }
        
        try:
            # We call tl.replace_text_in_document directly to see the traceback
            from docx import Document
            doc = Document(f)
            tl.replace_text_in_document(
                doc,
                translation_map,
                case_threshold=30,
                target_col=target_col,
                metadata=metadata,
                process_settings=process_steps
            )
            print("Successfully processed document!")
        except Exception as e:
            print("Traceback:")
            traceback.print_exc()

if __name__ == "__main__":
    main()
