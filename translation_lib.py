import pandas as pd
import os
import re
import json
import unicodedata
import openpyxl
from datetime import datetime
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt, Cm
from docx.oxml.ns import qn

# Use absolute path for Streamlit Cloud compatibility
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DICTIONARY_FILE = os.path.join(BASE_DIR, "Dictionary.xlsx")
CLEANV_FILE = os.path.join(BASE_DIR, "clean_v.json")
PARA_TEMPLATE_FILE = os.path.join(BASE_DIR, "para_template.json")

def clean_text(text):
    """
    Ultra-robust text cleaning for Word documents.
    Normalizes to NFC, strips invisible controls/soft hyphens, and collapses all whitespace.
    This ensures matching works even with hidden formatting characters or different XML namespaces.
    """
    if not isinstance(text, str) or pd.isna(text):
        return ""
    # Normalize to NFC (Normalization Form C) to ensure consistent Vietnamese character encoding
    text = unicodedata.normalize('NFC', str(text))
    
    # Remove hidden control characters, soft hyphens (\u00ad), and zero-width spaces
    # These often appear in Word headers/footers and prevent text matching.
    text = re.sub(r'[\u200b\ufeff\u00ad\u0000-\u0008\u000e-\u001f]', '', text)
    
    # Replace all whitespace sequences (tabs, newlines, non-breaking spaces) with a standard space
    text = re.sub(r'\s+', ' ', text)
    
    return text.strip()

def load_cleanv_map():
    """
    Loads Vietnamese text normalization map from clean_v.json.
    """
    if os.path.exists(CLEANV_FILE):
        try:
            with open(CLEANV_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading clean_v.json: {e}")
            return {}
    return {}

def load_para_template_map():
    """
    Loads Paragraph Templates from para_template.json.
    """
    if os.path.exists(PARA_TEMPLATE_FILE):
        try:
            with open(PARA_TEMPLATE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading para_template.json: {e}")
            return {}
    return {}

def _apply_para_templates_to_container(container, para_map, target_col, replaced_paragraphs):
    """
    Helper to apply paragraph templates to a container (Document, Header, Footer).
    """
    count = 0
    # docx Paragraphs
    for para in container.paragraphs:
        # We clean the text for matching
        text = clean_text(para.text)
        if not text: continue
        
        # Sort keys by length descending
        sorted_keys = sorted(para_map.keys(), key=len, reverse=True)
        for vn_key in sorted_keys:
            if vn_key in text:
                # Found a template match
                template_data = para_map[vn_key]
                new_text = template_data.get(target_col, "")
                if new_text:
                    # Replace entire paragraph text
                    para.text = new_text
                    replaced_paragraphs.add(para) # Mark as replaced
                    count += 1
                    break # Only one template per paragraph
                    
    return count

def apply_paragraph_templates(doc, para_map, target_col):
    """
    Applies paragraph-level templates to the entire document.
    Returns the count of replacements and a set of modified paragraph objects.
    """
    if not para_map:
        return 0, set()
        
    total_count = 0
    replaced_paragraphs = set()

    # 1. Main body
    total_count += _apply_para_templates_to_container(doc, para_map, target_col, replaced_paragraphs)
    
    # 2. Tables (Cells are containers)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                total_count += _apply_para_templates_to_container(cell, para_map, target_col, replaced_paragraphs)
                
    # 3. Headers and Footers
    for section in doc.sections:
        total_count += _apply_para_templates_to_container(section.header, para_map, target_col, replaced_paragraphs)
        total_count += _apply_para_templates_to_container(section.footer, para_map, target_col, replaced_paragraphs)
        if section.different_first_page_header_footer:
            total_count += _apply_para_templates_to_container(section.first_page_header, para_map, target_col, replaced_paragraphs)
            total_count += _apply_para_templates_to_container(section.first_page_footer, para_map, target_col, replaced_paragraphs)
        try:
            total_count += _apply_para_templates_to_container(section.even_page_header, para_map, target_col, replaced_paragraphs)
            total_count += _apply_para_templates_to_container(section.even_page_footer, para_map, target_col, replaced_paragraphs)
        except: pass
        
    return total_count, replaced_paragraphs

def _apply_cleanv_to_container(container, cleanv_map):
    """
    Helper to apply CleanV normalization (substring replacement) to all text in a container.
    """
    count = 0
    if not cleanv_map: return 0
    
    # Sort keys by length descending
    sorted_keys = sorted(cleanv_map.keys(), key=len, reverse=True)
    
    for para in container.paragraphs:
        changed_para = False
        for run in para.runs:
            if not run.text: continue
            new_text, changed = apply_normalization_to_text(run.text, cleanv_map)
            if changed:
                run.text = new_text
                changed_para = True
        if changed_para:
            count += 1
            
    return count

def apply_cleanv_normalization(doc, cleanv_map):
    """
    Applies Vietnamese text normalization to the entire document.
    """
    if not cleanv_map:
        return 0
        
    total_count = 0
    # 1. Main body
    total_count += _apply_cleanv_to_container(doc, cleanv_map)
    
    # 2. Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                total_count += _apply_cleanv_to_container(cell, cleanv_map)
                
    # 3. Headers and Footers
    for section in doc.sections:
        total_count += _apply_cleanv_to_container(section.header, cleanv_map)
        total_count += _apply_cleanv_to_container(section.footer, cleanv_map)
        if section.different_first_page_header_footer:
            total_count += _apply_cleanv_to_container(section.first_page_header, cleanv_map)
            total_count += _apply_cleanv_to_container(section.first_page_footer, cleanv_map)
        try:
            total_count += _apply_cleanv_to_container(section.even_page_header, cleanv_map)
            total_count += _apply_cleanv_to_container(section.even_page_footer, cleanv_map)
        except: pass
        
    return total_count

def apply_special_textbox_formatting(doc, target_col):
    """
    Finds "BẢN DỰ THẢO" in textboxes and replaces it with language-specific text.
    Also applies special formatting (10.5pt, margins, autosize).
    """
    find_text = "BẢN DỰ THẢO"
    
    replace_data = {
        "E": "DRAFT FOR DISCUSSION ONLY",
        "Hs": "草 稿",
        "Ht": "草 稿"
    }
    
    target_text = replace_data.get(target_col, find_text)
    is_chinese = target_col in ["Hs", "Ht"]
    
    # 1. Find all textbox contents in body, headers, footers
    # We use a combined list of all XML elements for a thorough hunt
    elements = [doc._element]
    for section in doc.sections:
        elements.extend([section.header._element, section.footer._element])
        if section.different_first_page_header_footer:
            elements.extend([section.first_page_header._element, section.first_page_footer._element])
        try:
            elements.extend([section.even_page_header._element, section.even_page_footer._element])
        except: pass

    count = 0
    for root in elements:
        textboxes = root.xpath('.//*[local-name()="txbxContent"]')
        for txbx in textboxes:
            # Check if it contains the target draft phrase
            txt_nodes = txbx.xpath('.//*[local-name()="t"]')
            full_text = "".join(t.text for t in txt_nodes if t.text)
            
            if find_text in full_text:
                # 2. Perform Replacement and Font Styling
                p_nodes = txbx.xpath('.//*[local-name()="p"]')
                for p_node in p_nodes:
                    para = Paragraph(p_node, doc)
                    if find_text in para.text:
                        para.text = para.text.replace(find_text, target_text)
                    
                    # Force font styling to all runs
                    for run in para.runs:
                        if is_chinese:
                            _set_run_fonts_refined(run, "Times New Roman", "DFKai-SB", 10.5)
                        else:
                            run.font.name = "Times New Roman"
                            run.font.size = Pt(10.5)
                
                # 3. XML Level Formatting (Margins and AutoSize)
                try:
                    # VBA MarginLeft=0, MarginRight=0, AutoSize=True, WordWrap=True
                    # VML Style (older textboxes)
                    v_textbox = txbx.getparent()
                    if v_textbox is not None and v_textbox.tag.endswith('textbox'):
                        v_textbox.set('inset', '0,0,0,0') # 0 margins
                        
                        v_shape = v_textbox.getparent()
                        if v_shape is not None and v_shape.tag.endswith('shape'):
                            style = v_shape.get('style', '')
                            # Add mso-fit-shape-to-text:t for AutoSize
                            if 'mso-fit-shape-to-text' not in style:
                                style += ';mso-fit-shape-to-text:t'
                            v_shape.set('style', style)
                    
                    # DrawingML Style (modern textboxes)
                    # We look for a sister element <wps:bodyPr> which controls margins/wrap
                    # txbxContent -> txbx -> (sister) bodyPr
                    wps_txbx = txbx.getparent()
                    if wps_txbx is not None and wps_txbx.tag.endswith('txbx'):
                        wps_wsp = wps_txbx.getparent()
                        if wps_wsp is not None:
                            bodyPr = wps_wsp.find(qn('wps:bodyPr'))
                            if bodyPr is not None:
                                bodyPr.set('lIns', '0')
                                bodyPr.set('rIns', '0')
                                bodyPr.set('tIns', '0')
                                bodyPr.set('bIns', '0')
                                # AutoSize in modern Word is often complex, but 0 margins helps.
                except:
                    pass
                count += 1
    return count

def apply_sizing_and_layout(doc, target_col="E"):
    """
    Applies specialized table sizing and layout formatting.
    """
    if not doc.tables:
        return

    # 1. SPECIAL CASE: Table 1 (Cover)
    # VBA logic: tbl.Columns.Count = 1 And tbl.Rows.Count = 7
    try:
        t1 = doc.tables[0]
        if len(t1.columns) == 1 and len(t1.rows) == 7:
            if target_col == "E":
                # VBA indices are 1-based. Row 1 is index 0.
                # R1: 14, R2: 10.5, R3: 10.5, R6: 12, R7: 12
                for i, size in [(0, 14), (1, 10.5), (2, 10.5), (5, 12), (6, 12)]:
                    for para in t1.rows[i].cells[0].paragraphs:
                        if not para.runs: para.add_run("") # Ensure at least one run
                        for run in para.runs:
                            run.font.size = Pt(size)
            elif target_col in ["Hs", "Ht"]:
                # R1: 16, R2: 13, R3: 10, R6: 15, R7: 15
                for i, size in [(0, 16), (1, 13), (2, 10), (5, 15), (6, 15)]:
                    for para in t1.rows[i].cells[0].paragraphs:
                        if not para.runs: para.add_run("")
                        for run in para.runs:
                            run.font.size = Pt(size)
    except: pass

    # 2. GLOBAL TABLE LOGIC
    for tbl in doc.tables:
        try:
            row_count = len(tbl.rows)
            col_count = len(tbl.columns)
            
            # Row count 37-40 -> Row 1 height = 1.01 cm
            if 37 <= row_count <= 40:
                tbl.rows[0].height = Cm(1.01)
                
            # Cols > 10 -> Font size = 7pt for all cells
            if col_count > 10:
                for row in tbl.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.font.size = Pt(7)
                                
            # 3-column table with specific text in Cell(0, 2)
            if col_count == 3:
                # VBA: tbl.cell(1, 3).Range.text (1-indexed) -> Cell(0, 2)
                cell_text = tbl.cell(0, 2).text
                if "Percentage of interest (%)" in cell_text:
                    # Setting width at cell level per row for better compatibility
                    for row in tbl.rows:
                        row.cells[1].width = Cm(4.98)
                        row.cells[2].width = Cm(4.07)
        except: pass

def _set_run_fonts_refined(run, latin_font, east_asia_font, size_pt=None):
    """
    Sets dual font families and optional size.
    Strengthened for maximum compatibility by setting all 4 font slots.
    """
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    
    # Set all slots to ensure consistent rendering across locales
    rFonts.set(qn('w:ascii'), latin_font)
    rFonts.set(qn('w:hAnsi'), latin_font)
    rFonts.set(qn('w:eastAsia'), east_asia_font)
    rFonts.set(qn('w:cs'), latin_font) # Complex scripts slots
    
    # Always hint how to handle the run
    if contains_chinese(run.text):
        rFonts.set(qn('w:hint'), 'eastAsia')
    else:
        rFonts.set(qn('w:hint'), 'default')
    
    if size_pt:
        run.font.size = Pt(size_pt)

def _process_paragraph_font_dual(para, target_col):
    """
    Splits runs in a paragraph to handle mixed Chinese/Latin formatting.
    Chinese -> DFKai-SB, 10pt
    Latin -> Times New Roman, Original Size
    """
    if not para.text:
        return

    old_runs = list(para.runs)
    if not old_runs:
        return

    # Clear existing runs
    p_el = para._element
    for r in old_runs:
        p_el.remove(r._element)
        
    # Regex to split by Chinese characters while keeping them
    cjk_pattern = r'([\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]+)'
    
    LATIN_FONT = "Times New Roman"
    CJK_FONT = "DFKai-SB"
    
    for old_run in old_runs:
        text = old_run.text
        if not text: continue
        
        # Get original size to preserve for Latin text
        orig_size_pt = None
        try:
            if old_run.font and old_run.font.size:
                orig_size_pt = old_run.font.size.pt
        except: pass
        
        if not contains_chinese(text):
            # Pure Latin/Numbers -> Keep original size
            new_run = para.add_run(text)
            _copy_run_format(old_run, new_run)
            _set_run_fonts_refined(new_run, LATIN_FONT, CJK_FONT, orig_size_pt)
        else:
            # Mixed content -> Split into small chunks
            parts = re.split(cjk_pattern, text)
            for part in parts:
                if not part: continue
                new_run = para.add_run(part)
                _copy_run_format(old_run, new_run)
                
                if contains_chinese(part):
                    # Chinese segment -> Enforce 10pt and CJK font
                    _set_run_fonts_refined(new_run, LATIN_FONT, CJK_FONT, 10)
                else:
                    # Latin/Number segment -> Keep original size and TNR font
                    _set_run_fonts_refined(new_run, LATIN_FONT, CJK_FONT, orig_size_pt)

def apply_chinese_font_formatting(doc, target_col):
    """
    Applies Chinese report font standards with dual-font and size preservation.
    """
    if target_col not in ["Hs", "Ht"]:
        return

    # 1. Main body
    for para in doc.paragraphs:
        _process_paragraph_font_dual(para, target_col)
    
    # 2. Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _process_paragraph_font_dual(para, target_col)
                
    # 3. Headers and Footers
    for section in doc.sections:
        # Loop through all possible headers/footers
        headers = [section.header, section.footer]
        if section.different_first_page_header_footer:
            headers.extend([section.first_page_header, section.first_page_footer])
        try:
            headers.extend([section.even_page_header, section.even_page_footer])
        except: pass
        
        for h in headers:
            if h:
                for para in h.paragraphs:
                    _process_paragraph_font_dual(para, target_col)
                for table in h.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                _process_paragraph_font_dual(para, target_col)

def format_dates_in_tables(doc, target_col="E"):
    """
    Finds dates in DD/MM/YYYY format in table cells and reformats them.
    English (E): MMM DD, YYYY
    Chinese (Hs/Ht): YYYY/MM/DD日
    """
    # Regex for DD/MM/YYYY (supports 1 or 2 digits for day/month)
    date_pattern = re.compile(r'^(0?[1-9]|[12][0-9]|3[01])/(0?[1-9]|1[0-2])/([0-9]{4})$')
    
    count = 0
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # 1. Aggressive cleaning like VBA: remove control chars < 32 and trim
                raw_text = "".join(ch for ch in cell.text if ord(ch) >= 32).strip()
                
                # 2. Check if text matches EXACTLY a date
                match = date_pattern.match(raw_text)
                if match:
                    try:
                        # 3. Parse date
                        day, month, year = match.groups()
                        dt = datetime(int(year), int(month), int(day))
                        
                        # 4. Reformat based on target language
                        if target_col == "E":
                            # VBA 'MMM dd, yyyy' -> 'Jan 01, 2023'
                            new_text = dt.strftime("%b %d, %Y")
                        elif target_col in ["Hs", "Ht"]:
                            # VBA 'yyyy/mm/dd' & "日" -> '2023/12/31日'
                            new_text = dt.strftime("%Y/%m/%d") + "日"
                        else:
                            continue
                            
                        # 5. Overwrite cell text
                        cell.text = new_text
                        count += 1
                    except:
                        pass
    return count

def contains_vietnamese(text):
    """
    Checks if a string contains characters unique to the Vietnamese language.
    Includes NFC-normalized characters.
    """
    if not text:
        return False
        
    # Standard list of specific Vietnamese characters (NFC)
    # Includes lowercase and uppercase variants
    vn_chars = "àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ"
    vn_chars += vn_chars.upper()
    vn_set = set(vn_chars)
    
    for char in text:
        if char in vn_set:
            return True
    return False

def contains_chinese(text):
    """
    Checks if a string contains CJK characters (Chinese characters) 
    including common CJK symbols and punctuation.
    """
    if not text:
        return False
    # Regex range for common CJK characters + CJK punctuation and fullwidth symbols
    pattern = r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff\u3000-\u303f\uff00-\uffef]'
    return bool(re.search(pattern, text))

def _copy_run_format(src_run, dest_run):
    """
    Clones standard font formatting from src_run to dest_run.
    """
    dest_run.bold = src_run.bold
    dest_run.italic = src_run.italic
    dest_run.underline = src_run.underline
    if src_run.font.name:
        dest_run.font.name = src_run.font.name
    if src_run.font.size:
        dest_run.font.size = src_run.font.size
    try:
        if src_run.font.color and src_run.font.color.rgb:
            dest_run.font.color.rgb = src_run.font.color.rgb
    except: pass

def _process_item_for_word_highlight(para):
    """
    Splits runs in a paragraph to highlight only the specific words containing Vietnamese.
    """
    old_runs = list(para.runs)
    if not old_runs:
        return

    # 1. Check if the paragraph has ANY Vietnamese first to avoid unnecessary splitting
    if not contains_vietnamese(para.text):
        return

    # 2. Clear existing runs from paragraph element safely
    p_el = para._element
    for r in old_runs:
        p_el.remove(r._element)
    
    # 3. Define Vietnamese characters for regex
    vn_chars = "àáảãạâầấẩẫậăằắẳẵặèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ"
    vn_chars += vn_chars.upper()
    
    # Pattern to match "words": sequences of alphanumeric + VN chars
    # Splits by everything else (punctuation, space, etc.)
    pattern = r'([a-zA-Z0-9' + vn_chars + r']+)'
    
    for old_run in old_runs:
        text = old_run.text
        if not text: continue
        
        if not contains_vietnamese(text):
            # Simple run, just restore it
            new_run = para.add_run(text)
            _copy_run_format(old_run, new_run)
        else:
            # Run contains Vietnamese, need to split it
            parts = re.split(pattern, text)
            for part in parts:
                if not part: continue
                new_run = para.add_run(part)
                _copy_run_format(old_run, new_run)
                if contains_vietnamese(part):
                    new_run.font.highlight_color = WD_COLOR_INDEX.YELLOW

def highlight_vietnamese_text(doc):
    """
    Scans the entire document and highlights only the specific WORDS 
    containing Vietnamese text in yellow.
    """
    # 1. Main body paragraphs
    for para in doc.paragraphs:
        _process_item_for_word_highlight(para)
        
    # 2. Table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _process_item_for_word_highlight(para)
                    
    # 3. Headers and Footers
    for section in doc.sections:
        containers = [
            section.header, section.footer,
            section.first_page_header, section.first_page_footer
        ]
        try:
            containers.extend([section.even_page_header, section.even_page_footer])
        except: pass
        
        for container in containers:
            if container:
                for para in container.paragraphs:
                    _process_item_for_word_highlight(para)
                # Tables in headers/footers
                for table in container.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                _process_item_for_word_highlight(para)

def apply_chinese_currency_cleanup(doc):
    """
    Final post-processing to fix specific currency redundant translations in Chinese.
    """
    corrections = {
        "越南盾（越南盾）": "越南盾（VND）",
        "美元（美元）": "美元（USD）"
    }
    
    def _clean_container(container):
        for para in container.paragraphs:
            original_text = para.text
            new_text = original_text
            for key, val in corrections.items():
                if key in new_text:
                    new_text = new_text.replace(key, val)
            if new_text != original_text:
                para.text = new_text

    # 1. Main body
    _clean_container(doc)
    
    # 2. Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _clean_container(cell)
                
    # 3. Headers and Footers
    for section in doc.sections:
        _clean_container(section.header)
        _clean_container(section.footer)
        if section.different_first_page_header_footer:
            _clean_container(section.first_page_header)
            _clean_container(section.first_page_footer)
        try:
            _clean_container(section.even_page_header)
            _clean_container(section.even_page_footer)
        except: pass

def apply_normalization_to_text(text, cleanv_map):
    """
    Applies Vietnamese text normalization using the cleanv_map.
    This replaces "dirty" Vietnamese patterns with "cleaned" ones.
    """
    if not text or not cleanv_map:
        return text, False
    
    new_text = text
    changed = False
    
    # Sort keys by length descending to match longest phrases first
    sorted_keys = sorted(cleanv_map.keys(), key=len, reverse=True)
    
    for key in sorted_keys:
        if key in new_text:
            new_text = new_text.replace(key, cleanv_map[key])
            changed = True
            
    return new_text, changed

# --- Excel Formula Simulation Helpers ---

def parse_vn_period_dates(period_str):
    """
    Parses 'từ ngày 04 tháng 11 năm 2025 đến ngày 31 tháng 12 năm 2025'
    Returns (start_date, end_date) as datetime objects
    """
    if not isinstance(period_str, str) or "đến" not in period_str:
        return None, None
    
    parts = period_str.split("đến")
    dates = []
    for part in parts:
        # Match 'ngày X tháng Y năm Z'
        # Group 1: day, Group 2: month, Group 3: year
        match = re.search(r"ngày\s+(\d+)\s+tháng\s+(\d+)\s+năm\s+(\d+)", part)
        if match:
            day, month, year = match.groups()
            try:
                dates.append(datetime(int(year), int(month), int(day)))
            except:
                dates.append(None)
        else:
            dates.append(None)
    
    if len(dates) >= 2:
        return dates[0], dates[1]
    return None, None

def format_excel_date_logic(dt, lang="vn"):
    if not isinstance(dt, datetime):
        return ""
    if lang == "vn":
        # 'Ngày 31 tháng 12 năm 2025'
        return f"Ngày {dt.day:02d} tháng {dt.month:02d} năm {dt.year}"
    elif lang == "en":
        # 'December 31, 2025'
        return dt.strftime("%B %d, %Y")
    elif lang == "cn":
        # '2025年12月31日'
        return f"{dt.year}年{dt.month:02d}月{dt.day:02d}日"
    elif lang == "short_vn":
        # '31/12/2025'
        return dt.strftime("%d/%m/%Y")
    return ""

def recalculate_dictionary_formulas(metadata):
    """
    Reproduces the logic of rows 6-64 from Dictionary.xlsx in Python.
    Returns a list of rows, each row is [VN, E, Hs, Ht]
    """
    name_vn = metadata.get("name_vn", "")
    name_trans = metadata.get("name_trans", "")
    
    # 1. Parse Dates
    def parse_dt(d_str):
        if not d_str: return None
        try: return datetime.strptime(d_str, "%d/%m/%Y")
        except: return None

    y_end_dt = parse_dt(metadata.get("year_end", ""))
    rep_date_dt = parse_dt(metadata.get("report_date", ""))
    p_out_start, p_out_end = parse_vn_period_dates(metadata.get("period_out", ""))
    
    # 2. Prepare Formatted Strings
    date_vn = format_excel_date_logic(y_end_dt, "vn")
    date_en = format_excel_date_logic(y_end_dt, "en")
    date_cn = format_excel_date_logic(y_end_dt, "cn")
    
    p_out = metadata.get("period_out", "")
    p_in = metadata.get("period_in", "")
    p_out_en_start = format_excel_date_logic(p_out_start, "en")
    p_out_en_end = format_excel_date_logic(p_out_end, "en")
    p_out_cn_start = format_excel_date_logic(p_out_start, "cn")
    p_out_cn_end = format_excel_date_logic(p_out_end, "cn")
    
    rep_year = str(y_end_dt.year) if y_end_dt else ""
    rep_vn = format_excel_date_logic(rep_date_dt, "vn")
    rep_en = format_excel_date_logic(rep_date_dt, "en")
    rep_cn = format_excel_date_logic(rep_date_dt, "cn")

    rows = []
    
    # 6: UPPER(A1 / Name)
    rows.append([name_vn.upper(), name_trans.upper(), name_trans.upper(), name_trans.upper()])
    
    # 7: Statements for year ended ...
    rows.append([
        f"Báo cáo tài chính cho năm tài chính kết thúc {date_vn.lower()}",
        f"Financial statements for the year ended {date_en}",
        f"截至{date_cn}止财务年度的财务报表",
        f"截至{date_cn}止財務年度之財務報表"
    ])
    
    # 8-9: Periods
    rows.append([
        f"Báo cáo tài chính cho giai đoạn {p_out}" if p_out else "",
        f"Financial statements for the period from {p_out_en_start} to {p_out_en_end}" if p_out_en_start else "",
        f"{p_out_cn_start}至{p_out_cn_end}期间的财务报表" if p_out_cn_start else "",
        f"{p_out_cn_start}至{p_out_cn_end}期間的財務報表" if p_out_cn_start else ""
    ])
    rows.append([
        f"Báo cáo tài chính cho giai đoạn {p_in}" if p_in else "",
        f"Financial statements for the period from {p_out_en_start} to {p_out_en_end}" if p_out_en_start else "",
        f"{p_out_cn_start}至{p_out_cn_end}期间的财务报表" if p_out_cn_start else "",
        f"{p_out_cn_start}至{p_out_cn_end}期間的財務報表" if p_out_cn_start else ""
    ])
    
    # 10-11: Opinion
    # Simplified long sentences logic based on Excel formulas
    opinion_vn = "Theo ý kiến của chúng tôi, Báo cáo tài chính đã phản ánh trung thực và hợp lý, trên các khía cạnh trọng yếu tình hình tài chính của Công ty tại "
    rows.append([
        f"{opinion_vn}{date_vn.lower()}, cũng như kết quả hoạt động kinh doanh và tình hình lưu chuyển tiền tệ cho năm tài chính kết thúc cùng ngày, phù hợp với chuẩn mực kế toán, chế độ kế toán doanh nghiệp Việt Nam và các quy định pháp lý có liên quan đến việc lập và trình bày Báo cáo tài chính.",
        f"In our opinion, the financial statements present fairly, in all material respects, the financial position of the Company as at {date_en}, and its financial performance and its cash flows for the year then ended in accordance with Vietnamese Accounting Standards, Vietnamese enterprise accounting system and applicable regulations relevant to the preparation and presentation of financial statements in Vietnam.",
        f"依本会计师的意见，上开财务报表 in all material respects 已真实、公允地反映公司在{date_cn}财务状况以及同日结束财务年度的经营成果和现金流量状况，且符合《越南会计准则》、《越南企业会计制度》以及财务报表编制和列报的相关法规。",
        f"依本會計師之意見，上開財務報表 in all material respects 已真實、公允地反映公司於{date_cn}財務狀況以及同日結束財務年度之經營成果及現金流量狀況，且符合《越南會計準則》、《越南企業會計制度》及財務報表編製及列報之相關法規。"
    ])
    rows.append([
        f"{opinion_vn}{date_vn.lower()}, cũng như kết quả hoạt động kinh doanh và tình hình lưu chuyển tiền tệ cho giai đoạn {p_out}, phù hợp với chuẩn mực kế toán, chế độ kế toán doanh nghiệp Việt Nam và các quy định pháp lý có liên quan đến việc lập và trình bày Báo cáo tài chính." if p_out else "",
        f"In our opinion, the financial statements present fairly, in all material respects, the financial position of the Company as at {date_en}, and its financial performance and its cash flows for the period from {p_out_en_start} to {p_out_en_end} in accordance with Vietnamese Accounting Standards, Vietnamese enterprise accounting system and applicable regulations relevant to the preparation and presentation of financial statements in Vietnam." if p_out_en_start else "",
        f"依本会计师的意见，上开财务报表 in all material respects 已真实、公允地反映公司在{date_cn}财务状况以及{p_out_cn_start}至{p_out_cn_end}期间的经营成果和现金流量状况，且符合《越南会计准则》、《越南企业会计制度》以及财务报表编制和列报的相关法规。" if p_out_cn_start else "",
        f"依本會計師之意見，上開財務報表 in all material respects 已真實、公允地反映公司於{date_cn}財務狀況以及{p_out_cn_start}至{p_out_cn_end}期間之經營成果及現金流量狀況，且符合《越南會計準則》、《越南企業會計制度》及財務報表編製及列報之相關法規。" if p_out_cn_start else ""
    ])

    # 12: Name (Standard)
    rows.append([name_vn, name_trans, name_trans, name_trans])
    
    # 13, 14, 15: Dates
    rows.append([date_vn, date_en, date_cn, date_cn])
    rows.append([date_vn.lower(), date_en, date_cn, date_cn])
    rows.append([date_vn.upper(), date_en.upper(), date_cn, date_cn])
    
    # 16, 17, 18, 19: Helpers
    rows.append([p_out, f"from {p_out_en_start} to {p_out_en_end}" if p_out_en_start else "", f"自{p_out_cn_start}至{p_out_cn_end}" if p_out_cn_start else "", ""])
    rows.append([f"từ ngày {format_excel_date_logic(p_out_start, 'short_vn')} đến ngày {format_excel_date_logic(p_out_end, 'short_vn')}" if p_out_start else "", "", "", ""])
    rows.append([p_in, f"From {p_out_en_start} to {p_out_en_end}" if p_out_en_start else "", f"自{p_out_cn_start}日至{p_out_cn_end}日" if p_out_cn_start else "", ""])
    rows.append(["", f"From {p_out_en_start} to {p_out_en_end}" if p_out_en_start else "", f"自{p_out_cn_start}日至{p_out_cn_end}日" if p_out_cn_start else "", ""])
    
    # 20: As at
    rows.append([f"Tại {date_vn.lower()}", f"As at {date_en}", f"于{date_cn}", f"於{date_cn}"])
    
    # 21, 22: For the year ended
    rows.append([f"Cho năm tài chính kết thúc {date_vn.lower()}", f"For the year ended {date_en}", f"截至{date_cn}止财务年度", f"截至{date_cn}止財務年度"])
    rows.append([f"cho năm tài chính kết thúc {date_vn.lower()}", f"for the year ended {date_en}", f"截至{date_cn}止财务年度", f"截至{date_cn}止財務年度"])
    
    # 23-26: Period variations
    rows.append([f"Cho giai đoạn {p_out}", f"For the period from {p_out_en_start} to {p_out_en_end}", f"{p_out_cn_start}至{p_out_cn_end}期间", ""])
    rows.append([f"Cho giai đoạn {format_excel_date_logic(p_out_start, 'short_vn')} đến {format_excel_date_logic(p_out_end, 'short_vn')}", f"For the period from {p_out_en_start} to {p_out_en_end}", "", ""])
    rows.append([f"cho giai đoạn {p_out}", f"for the period from {p_out_en_start} to {p_out_en_end}", "", ""])
    rows.append([f"cho giai đoạn {format_excel_date_logic(p_out_start, 'short_vn')} đến {format_excel_date_logic(p_out_end, 'short_vn')}", f"for the period from {p_out_en_start} to {p_out_en_end}", "", ""])
    
    # 27, 28: Reporting Dates
    rows.append([rep_vn, rep_en, rep_cn, rep_cn])
    rows.append([rep_vn.lower(), rep_en, rep_cn, rep_cn])
    
    # 29, 30: Names/Dates again (Placeholder Keys restored)
    rows.append(["[têncôngty]", name_trans, name_trans, name_trans])
    rows.append(["[ngàybáocáo]", rep_en, rep_cn, rep_cn])
    
    # 31: Year
    rows.append(["[nămbáocáo]", rep_year, rep_year, rep_year])
    
    # 32: Rep date (Placeholder Key)
    rows.append(["[ngàykếtthúcnăm]", date_en, date_cn, date_cn])
    
    # 33, 34, 35: Caps
    rows.append([f"CHO NĂM TÀI CHÍNH KẾT THÚC {date_vn.upper()}", f"FOR THE YEAR ENDED {date_en.upper()}", f"截至{date_cn}止财务年度", f"截至{date_cn}止財務年度"])
    rows.append([f"CHO GIAI ĐOẠN {p_out.upper()}", f"FOR THE PERIOD FROM {p_out_en_start.upper()} TO {p_out_en_end.upper()}", "", ""])
    rows.append([f"CHO GIAI ĐOẠN {p_in.upper()}", f"FOR THE PERIOD FROM {p_out_en_start.upper()} TO {p_out_en_end.upper()}", "", ""])
    
    # 36-64 (Auditors, Tax etc. - simplified logic based on your Excel)
    rows.append([f"Công ty TNHH Kiểm toán U&I đã kiểm toán Báo cáo tài chính cho năm tài chính kết thúc {date_vn.lower()} và bày tỏ nguyện vọng tiếp tục được bổ nhiệm làm kiểm toán viên cho Công ty.", f"The auditors, U&I Auditing Company Limited, have performed audit on the Company’s financial statements for the year ended {date_en} and have expressed their willingness to accept reappointment.", "", ""])
    rows.append([f"Công ty TNHH Kiểm toán U&I đã kiểm toán Báo cáo tài chính cho giai đoạn {p_out} và bày tỏ nguyện vọng tiếp tục được bổ nhiệm làm kiểm toán viên cho Công ty." if p_out else "", f"The auditors, U&I Auditing Company Limited, have performed audit on the Company’s financial statements for the period from {p_out_en_start} to {p_out_en_end} and have expressed their willingness to accept reappointment.", "", ""])
    
    rows.append([f"Chi phí thuế thu nhập doanh nghiệp cho năm tài chính kết thúc {date_vn.lower()} được tính trên thu nhập tính thuế ước tính. Chi phí thuế thu nhập doanh nghiệp này sẽ được cơ quan thuế xác định lại thông qua các cuộc kiểm tra.", f"Corporate income tax expense for the year ended {date_en} is calculated on the estimated assessable income. Corporate income tax expense will be determined again by the tax authority through their tax reviews. ", "", ""])
    rows.append([f"Chi phí thuế thu nhập doanh nghiệp cho giai đoạn {p_out} được tính trên thu nhập tính thuế ước tính. Chi phí thuế thu nhập doanh nghiệp này sẽ được cơ quan thuế xác định lại thông qua các cuộc kiểm tra." if p_out else "", f"Corporate income tax expense for the period from {p_out_en_start} to {p_out_en_end} is calculated on the estimated assessable income. Corporate income tax expense will be determined again by the tax authority through their tax reviews. ", "", ""])
    
    rows.append([f"Do đó, không có chi phí thuế thu nhập doanh nghiệp hoãn lại được ghi nhận trong năm {rep_year}.", f"Therefore, no deferred corporate income tax expense is provided in the year {rep_year}.", f"据此，并无递延所得税费用在{rep_year}年度内得以认列。", f"據此，並無遞延所得稅費用於{rep_year}年度內得以認列。"])
    rows.append([f"Do đó, không có chi phí thuế thu nhập doanh nghiệp hoãn lại được ghi nhận cho giai đoạn {p_out}.", f"Therefore, no deferred corporate income tax expense is provided in the period from {p_out_en_start} to {p_out_en_end}", f"据此，并无递延所得税费用{p_out_cn_start}至{p_out_cn_end}期间内得以认列。", f"據此，並無遞延所得稅費用{p_out_cn_start}至{p_out_cn_end}期間內得以認列。"])
    rows.append([f"Do đó, không có chi phí thuế thu nhập doanh nghiệp hoãn lại được ghi nhận trong giai đoạn {p_out}.", f"Same as above", "", ""])
    
    rows.append([f"Trong năm tài chính kết thúc {date_vn.lower()}, Công ty có các nghiệp vụ kinh tế quan trọng với các bên liên quan được trình bày ở bảng sau:", f"In the year ended {date_en}, the Company entered into significant economic transactions with its related parties as shown in the following table:", f"在截止{date_cn}财务年度之内，公司与其关联方发生重大经济交易，其列示在下表：", f"在截止{date_cn}財務年度之內，公司與其關聯方發生重大經濟交易，其列示於下表："])
    rows.append([f"Trong giai đoạn {p_out}, Công ty có các nghiệp vụ kinh tế quan trọng với các bên liên quan được trình bày ở bảng sau:" if p_out else "", f"In the period from {p_out_en_start} to {p_out_en_end}, the Company entered into significant economic transactions with its related parties as shown in the following table:", "", ""])
    rows.append([f"Trong năm tài chính kết thúc {date_vn.lower()}, Công ty có các nghiệp vụ kinh tế quan trọng và các khoản phải thu, phải trả với các bên liên quan được trình bày ở bảng sau", f"In the year ended {date_en}, the Company entered into significant economic transactions, receivables and payables with its related parties as shown in the following table", "", ""])
    rows.append([f"Trong giai đoạn {p_out}, Công ty có các nghiệp vụ kinh tế quan trọng và các khoản phải thu, phải trả với các bên liên quan được trình bày ở bảng sau:" if p_out else "", f"In the period from {p_out_en_start} to {p_out_en_end}, the Company entered into significant economic transactions, receivables and payables with its related parties as shown in the following table", "", ""])
    rows.append([f"Trong năm tài chính kết thúc {date_vn.lower()}, Công ty không phát sinh các nghiệp vụ kinh tế quan trọng với các bên liên quan.", f"In the year ended {date_en}, the Company had no significant economic transactions with its related parties.", "", ""])
    rows.append([f"Trong giai đoạn {p_out}, Công ty không phát sinh các nghiệp vụ kinh tế quan trọng với các bên liên quan." if p_out else "", f"In the period from {p_out_en_start} to {p_out_en_end}, the Company had no significant economic transactions with its related parties.", "", ""])
    
    approvers = ["Tổng Giám đốc", "Ban Giám đốc", "Hội đồng Thành viên và Ban Giám đốc", "Giám đốc điều hành", "Ban Tổng Giám đốc", "Hội đồng Thành viên", "Giám đốc", "Chủ tịch", "Hội đồng Quản trị", "Chủ sở hữu"]
    for app in approvers:
        rows.append([f"Báo cáo tài chính được phê duyệt bởi {app} Công ty để phát hành vào {rep_vn.lower()}.", "", "", ""])
        
    rows.append([f"Công ty không lập dự phòng chi phí thuế thu nhập doanh nghiệp hiện hành vì Công ty không có thu nhập chịu thuế trong năm {rep_year}.", f"No provision for current corporate income tax expense has been provided as the Company has no taxable income arising in the year {rep_year}.", "", ""])
    rows.append([f"Công ty không lập dự phòng chi phí thuế thu nhập doanh nghiệp hiện hành vì Công ty không có thu nhập chịu thuế cho giai đoạn {p_out}.", f"No provision for current corporate income tax expense has been provided as the Company has no taxable income arising in the period from {p_out_en_start} to {p_out_en_end}", "", ""])
    rows.append([f"Công ty không lập dự phòng chi phí thuế thu nhập doanh nghiệp hiện hành vì Công ty không có thu nhập chịu thuế trong giai đoạn {p_out}.", "Same as above", "", ""])
    rows.append([f"Công ty không lập dự phòng chi phí thuế thu nhập doanh nghiệp hiện hành vì Công ty không có thu nhập tính thuế sau khi trừ chuyển lỗ trong năm {rep_year}.", f"No provision for current corporate income tax expense has been provided as the Company has no assessable income after deducting tax loss brought forward in the year {rep_year}.", "", ""])
    rows.append([f"Công ty không lập dự phòng chi phí thuế thu nhập doanh nghiệp hiện hành vì Công ty không có thu nhập tính thuế sau khi trừ chuyển lỗ cho giai đoạn {p_out}.", f"No provision for current corporate income tax expense has been provided as the Company has no assessable income after deducting tax loss brought forward in the period from {p_out_en_start} to {p_out_en_end}", "", ""])
    rows.append([f"Đến {date_vn.lower()}, Công ty vẫn chưa tiến hành hoạt động sản xuất kinh doanh.", f"As at {date_en}, the Company’s operation has yet to start.", "", ""])

    return rows

def load_excel_dictionary():
    """
    Loads metadata and dictionary from Dictionary.xlsx.
    Metadata: A1, B1, C1, D1, A2, A3
    Dictionary: Row 5 (headers), Row 6+ (data)
    """
    metadata = {}
    df = None
    
    if os.path.exists(DICTIONARY_FILE):
        try:
            # 1. Load Metadata using openpyxl
            wb = openpyxl.load_workbook(DICTIONARY_FILE, data_only=True)
            ws = wb.active
            
            metadata = {
                "name_vn": clean_text(ws['A1'].value),
                "name_trans": clean_text(ws['B1'].value),
                "year_end": clean_text(ws['C1'].value),
                "report_date": clean_text(ws['D1'].value),
                "period_out": clean_text(ws['A2'].value),
                "period_in": clean_text(ws['A3'].value)
            }
            wb.close()
            
            # 2. Load Dictionary Data using pandas (Start from Row 5)
            # Row 5 is header, so header=4 (0-indexed)
            df = pd.read_excel(DICTIONARY_FILE, header=4)
            
            if df is not None and not df.empty:
                # Normalize columns
                for col in df.columns:
                    if df[col].dtype == object:
                        df[col] = df[col].apply(clean_text)
                # Ensure 'Vietnamese' column exists
                if 'Vietnamese' in df.columns:
                    df = df.dropna(subset=['Vietnamese'])
            
            return metadata, df
        except Exception as e:
            print(f"Error loading excel dictionary: {e}")
            return {}, None
    return {}, None

def save_excel_metadata(metadata, df=None):
    """
    Saves metadata to A1:D1, A2, A3 and optionally updates dictionary data.
    Uses openpyxl to preserve formulas in other cells.
    """
    try:
        if not os.path.exists(DICTIONARY_FILE):
            # Create a new workbook if it doesn't exist
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Dictionary"
        else:
            wb = openpyxl.load_workbook(DICTIONARY_FILE)
            ws = wb.active

        # Write Metadata
        ws['A1'] = metadata.get("name_vn", "")
        ws['B1'] = metadata.get("name_trans", "")
        ws['C1'] = metadata.get("year_end", "")
        ws['D1'] = metadata.get("report_date", "")
        ws['A2'] = metadata.get("period_out", "")
        ws['A3'] = metadata.get("period_in", "")

        # Recalculate Rows 6-64 based on new metadata
        # (Simulating Excel Formulas in Python)
        calculated_rows = recalculate_dictionary_formulas(metadata)
        
        # Write calculated values starting at row 6
        for i, row_data in enumerate(calculated_rows):
            target_row = 6 + i
            if target_row > 64: break
            for j, val in enumerate(row_data):
                if val: # Only overwrite if we have a calculated value
                    ws.cell(row=target_row, column=j+1, value=val)

        # Write Dictionary Data if provided (starting from Row 5)
        if df is not None:
            # Write headers at Row 5
            headers = df.columns.tolist()
            for col_idx, header in enumerate(headers, start=1):
                ws.cell(row=5, column=col_idx, value=header)
            
            # Write data from Row 6
            for row_idx, row in enumerate(df.values, start=6):
                for col_idx, value in enumerate(row, start=1):
                    # We only overwrite if value is not a formula (very basic check)
                    ws.cell(row=row_idx, column=col_idx, value=value)

        wb.save(DICTIONARY_FILE)
        return True
    except Exception as e:
        print(f"Error saving excel: {e}")
        return False

def load_dictionary():
    """
    Backwards compatibility wrapper for load_excel_dictionary.
    Returns only the DataFrame.
    """
    _, df = load_excel_dictionary()
    return df

def save_dictionary(df):
    """
    Backwards compatibility wrapper for save_excel_metadata.
    """
    # Try to load existing metadata first to preserve it
    metadata, _ = load_excel_dictionary()
    return save_excel_metadata(metadata, df)

def get_translation_map(df, target_lang):
    """
    Creates a dictionary of Vietnamese -> Target Language.
    """
    if df is None or target_lang not in df.columns:
        return {}
    
    # Create map, dropping rows with null translations in the target column
    subset = df[['Vietnamese', target_lang]].dropna()
    return dict(zip(subset['Vietnamese'], subset[target_lang]))

def find_missing_terms(text, translation_map):
    """
    Optional: Find terms that might be missing (not used currently based on user request).
    """
    pass

def prepare_translation_list(translation_map, case_threshold=100):
    """
    Pre-processes and sorts translation terms for efficiency.
    """
    # Sort keys by length descending to match longest phrases first
    sorted_keys = sorted(translation_map.keys(), key=len, reverse=True)
    prepared = []
    for key in sorted_keys:
        val = translation_map[key]
        if len(key) >= case_threshold:
            # Pre-compile regex for long phrases
            pattern = re.compile(re.escape(key), re.IGNORECASE)
            prepared.append((True, pattern, val))
        else:
            # Keep as string for fast replacement
            prepared.append((False, key, val))
    return prepared

def apply_translations_to_text(text, prepared_list):
    """
    Core string replacement logic using the pre-compiled translation list.
    """
    if not text:
        return text, False
        
    new_text = text
    changed = False
    
    for is_regex, key_or_pattern, val in prepared_list:
        if is_regex:
            # Case-insensitive for long phrases (pre-compiled regex)
            if key_or_pattern.search(new_text):
                new_text = key_or_pattern.sub(val, new_text)
                changed = True
        else:
            # Case-sensitive for short terms (literal replacement)
            if key_or_pattern in new_text:
                new_text = new_text.replace(key_or_pattern, val)
                changed = True
    return new_text, changed

def apply_translations_to_paragraph(paragraph, prepared_list):
    """
    Applies a pre-processed list of translations to a paragraph.
    Works for single-paragraph units.
    """
    inline = paragraph.runs
    if not inline and not paragraph.text:
        return False

    full_text = "".join(run.text for run in inline)
    
    # Normalize document text to ensure matching with dictionary
    full_text = clean_text(full_text)
    
    new_text, changed = apply_translations_to_text(full_text, prepared_list)
    
    if changed:
        # Update runs while preserving paragraph-level formatting
        if len(inline) > 0:
            for i, run in enumerate(inline):
                run.text = new_text if i == 0 else ""
        else:
            paragraph.add_run(new_text)
    return changed

def replace_text_in_paragraph(paragraph, translation_map, case_threshold=100):
    """
    Backwards compatibility wrapper for apply_translations_to_paragraph.
    """
    prepared = prepare_translation_list(translation_map, case_threshold)
    return apply_translations_to_paragraph(paragraph, prepared)

def _process_container(container, prepared_list, replaced_paragraphs=None):
    """
    Enhanced processing of containers (Document, Header, Footer).
    Matches multi-paragraph terms inside table cells and text boxes by 
    joining them into a single logical unit.
    """
    count = 0
    element = container._element
    processed_p_elements = set()
    
    if replaced_paragraphs is None:
        replaced_paragraphs = set()

    # 1. Process Multi-paragraph Containers (Cells, Textboxes)
    # tc: Table Cell, txbxContent: Textbox Content
    containers = element.xpath('.//*[local-name()="tc"] | .//*[local-name()="txbxContent"]')
    
    for c_el in containers:
        # Find paragraphs directly within this container (or descendants)
        p_elements = c_el.xpath('.//*[local-name()="p"]')
        if not p_elements: continue
        
        # Build full text by joining paragraph contents
        p_texts = []
        for p_el in p_elements:
            t_nodes = p_el.xpath('.//*[local-name()="t"]')
            p_texts.append("".join(t.text for t in t_nodes if t.text))
        
        full_text = " ".join(p_texts)
        cleaned_body_text = clean_text(full_text)
        
        new_text, changed = apply_translations_to_text(cleaned_body_text, prepared_list)
        
        if changed:
            # Found a match across the entire container unit (e.g. "Mã\nsố")
            # We consolidate the translation into the first paragraph and clear others
            first_p = Paragraph(p_elements[0], container)
            inline = first_p.runs
            if len(inline) > 0:
                for i, run in enumerate(inline):
                    run.text = new_text if i == 0 else ""
            else:
                first_p.add_run(new_text)
                
            # Remove all other paragraphs in this container to prevent duplicates/layout issues
            for other_p_el in p_elements[1:]:
                parent = other_p_el.getparent()
                if parent is not None:
                    parent.remove(other_p_el)
            count += 1
            
        # Mark these paragraphs as handled
        for p_el in p_elements:
            processed_p_elements.add(p_el)
            
    # 2. Process lone paragraphs (not inside cells or textboxes)
    all_p_elements = element.xpath('.//*[local-name()="p"]')
    for p_el in all_p_elements:
        if p_el not in processed_p_elements:
            p = Paragraph(p_el, container)
            
            # SKIP if this paragraph was modified by ParaTemplate
            is_replaced = False
            if replaced_paragraphs:
                for rp in replaced_paragraphs:
                    if rp._element == p_el:
                        is_replaced = True
                        break
            
            if not is_replaced:
                if apply_translations_to_paragraph(p, prepared_list):
                    count += 1
                
    return count

def apply_metadata_placeholders(doc, metadata, target_col):
    """
    Directly replaces tag-style placeholders using the provided metadata.
    Acts as a fail-safe for tags like [têncôngty].
    """
    if not metadata: return 0
    
    # Map tag variations to metadata values
    # Note: Values are already calculated fragment strings
    y_end_dt = datetime.strptime(metadata.get("year_end", "31/12/2025"), "%d/%m/%Y") if metadata.get("year_end") else None
    rep_date_dt = datetime.strptime(metadata.get("report_date", "01/01/2026"), "%d/%m/%Y") if metadata.get("report_date") else None
    
    date_formatted = format_excel_date_logic(y_end_dt, "en" if target_col == "E" else "cn")
    rep_date_formatted = format_excel_date_logic(rep_date_dt, "en" if target_col == "E" else "cn")
    rep_year = str(y_end_dt.year) if y_end_dt else ""

    tag_map = {
        "[têncôngty]": metadata.get("name_trans", ""),
        "[tên công ty]": metadata.get("name_trans", ""),
        "[ngàykếtthúcnăm]": date_formatted + ("日" if target_col != "E" else ""),
        "[ngàybáocáo]": rep_date_formatted + ("日" if target_col != "E" else ""),
        "[nămbáocáo]": rep_year,
        "[giaiđoạn]": metadata.get("period_out", "") # Simplified, usually handled by dictionary
    }
    
    count = 0
    # Process all paragraphs, tables, etc using a simple prepared list
    prepared = []
    for k, v in tag_map.items():
        if v: prepared.append((False, k, v))

    count += _process_container_for_metadata(doc, prepared)
    
    for section in doc.sections:
        count += _process_container_for_metadata(section.header, prepared)
        count += _process_container_for_metadata(section.footer, prepared)
        # Check for text boxes specifically for metadata
        # (Already handled by _process_container_for_metadata calling deep traversal)
        
    return count

def _process_container_for_metadata(container, prepared_list):
    """Helper to run direct replacements in a container."""
    count = 0
    element = container._element
    # We target paragraphs directly for speed in metadata pass
    for p_el in element.xpath('.//*[local-name()="p"]'):
        p = Paragraph(p_el, container)
        if apply_translations_to_paragraph(p, prepared_list):
            count += 1
    return count

def replace_text_in_document(doc, translation_map, case_threshold=100, cleanv_map=None, para_map=None, target_col="E", metadata=None):
    """
    Performs global search and replace in paragraphs, tables, headers and footers.
    Follows exact order: 
    0. Metadata Placeholder Replacement (Safety Pass)
    1. CleanV Normalization (Unicode & Corrections)
    2. ParaTemplate Swaps (Full paragraph)
    3. Dictionary Translation (Sub-paragraph)
    """
    # 0. Load maps if not provided but exist
    if cleanv_map is None:
        cleanv_map = load_cleanv_map()
    if para_map is None:
        para_map = load_para_template_map()

    total_count = 0

    # Step 0: Metadata Placeholder Replacement
    if metadata:
        apply_metadata_placeholders(doc, metadata, target_col)

    # Pass 1: Global Normalization (CleanV)
    # This standardizes Vietnamese text everywhere first.
    if cleanv_map:
        apply_cleanv_normalization(doc, cleanv_map)

    # Pass 2: Paragraph Template Replacements (Full paragraph swaps)
    # Replaces boilerplates with translated templates.
    replaced_paras = set()
    if para_map:
        _, replaced_paras = apply_paragraph_templates(doc, para_map, target_col)

    # Pass 3: Dictionary-based replacements (Final translation)
    # Skip paragraphs that were already swapped by templates.
    prepared_list = prepare_translation_list(translation_map, case_threshold)
    
    # 3. Add normalization list (CleanV) to the start of prepared_list
    if cleanv_map:
        norm_list = []
        for key in sorted(cleanv_map.keys(), key=len, reverse=True):
            norm_list.append((False, key, cleanv_map[key])) # Using same format as prepared_list
        
        # We can prepend the normalization list to our processing, 
        # BUT it's better to do a distinct pass or just handle it as part of the loop.
        # Actually, adding them to the start of prepared_list is efficient IF we want to do it in one pass.
        # HOWEVER, the user specifically mentioned "sau khi coding xong... xử lý đồng bộ unicode", 
        # suggesting it should be a deliberate step.
        # We'll prepend them to prepared_list so they run FIRST.
        prepared_list = norm_list + prepared_list

    # 3. Process the main document body
    total_count += _process_container(doc, prepared_list, replaced_paragraphs=replaced_paras)
    
    # 4. Process all headers and footers in all sections
    for section in doc.sections:
        # Primary header and footer
        total_count += _process_container(section.header, prepared_list, replaced_paragraphs=replaced_paras)
        total_count += _process_container(section.footer, prepared_list, replaced_paragraphs=replaced_paras)
        
        # First page header and footer
        if section.different_first_page_header_footer:
            total_count += _process_container(section.first_page_header, prepared_list, replaced_paragraphs=replaced_paras)
            total_count += _process_container(section.first_page_footer, prepared_list, replaced_paragraphs=replaced_paras)
            
        # Even page header and footer
        # Note: python-docx handles this via odd_and_even_pages_header_footer (but we check headers directly)
        try:
            total_count += _process_container(section.even_page_header, prepared_list, replaced_paragraphs=replaced_paras)
            total_count += _process_container(section.even_page_footer, prepared_list, replaced_paragraphs=replaced_paras)
        except:
            # Not all versions of docx or documents have these defined
            pass
            
    # Step 4: Final Chinese Currency Cleanup (Hs/Ht only)
    if target_col in ["Hs", "Ht"]:
        apply_chinese_currency_cleanup(doc)
        
    # Step 5: Format dates in tables
    format_dates_in_tables(doc, target_col)
    
    # Step 7: Dual-font formatting (Chinese only)
    if target_col in ["Hs", "Ht"]:
        apply_chinese_font_formatting(doc, target_col)
        
    # Step 8: Specialized Table Sizing and Layout
    apply_sizing_and_layout(doc, target_col)
    
    # Step 9: Specialized TextBox/Draft Handling
    apply_special_textbox_formatting(doc, target_col)
    
    # Step 10: Highlight remaining Vietnamese text
    # This must be the absolute final step to ensure all text is scanned.
    highlight_vietnamese_text(doc)

    return total_count
