
import unicodedata

# Decomposed 'ạ' (a + dot below)
decomposed_a = "a" + "\u0323"
# Precomposed 'ạ'
precomposed_a = "\u1ea1"

with open('scratch/unicode_test_results.txt', 'w', encoding='utf-8') as f:
    f.write(f"Decomposed: {repr(decomposed_a)}\n")
    f.write(f"Precomposed: {repr(precomposed_a)}\n")

    normalized = unicodedata.normalize('NFC', decomposed_a)
    f.write(f"Normalized: {repr(normalized)}\n")
    f.write(f"Equal? {normalized == precomposed_a}\n")

    # 'Tổng Giám đốc' from the doc
    doc_text = "Tổng Gia" + "\u0301" + "m đ" + "\u00f4" + "\u0301" + "c"
    f.write(f"Doc text: {repr(doc_text)}\n")
    norm_doc = unicodedata.normalize('NFC', doc_text)
    f.write(f"Norm doc: {repr(norm_doc)}\n")

    dict_text = "Tổng Giám đốc" # NFC
    norm_dict = unicodedata.normalize('NFC', dict_text)
    f.write(f"Norm dict: {repr(norm_dict)}\n")

    f.write(f"Match? {norm_dict in norm_doc}\n")
    f.write(f"Exact Match? {norm_dict == norm_doc}\n")
