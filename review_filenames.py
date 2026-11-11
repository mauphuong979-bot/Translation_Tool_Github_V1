import re
from pathlib import Path, PureWindowsPath


OUTPUT_SUFFIXES = {
    "original": "_original",
    "clean": "_clean",
    "review": "_TC",
    "ai_review": "_AI_Form",
    "ai_result": "_AI_TC",
}
WORKFLOW_SUFFIX = re.compile(r"(?:_AI_translation_worksheet|_AI_reviewed|_AI_Form|_AI_TC|_review|_TC|_original|_clean)$", re.IGNORECASE)
LEGACY_TIMESTAMP = re.compile(r"_(E|Hs|Ht)_tool_(\d{6})[ _](\d{4})$", re.IGNORECASE)


def review_output_filename(filename, output_kind):
    stem = PureWindowsPath(Path(filename).name).stem
    while match := WORKFLOW_SUFFIX.search(stem):
        stem = stem[:match.start()]
    stem = LEGACY_TIMESTAMP.sub(r"_\1_\2_\3", stem)
    return f"{stem or 'Report'}{OUTPUT_SUFFIXES[output_kind]}.docx"
