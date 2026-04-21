
import docx
from docx.oxml.ns import qn

def inspect_structure(docx_path):
    doc = docx.Document(docx_path)
    body = doc._body._body
    print(f"Body children count: {len(body)}")
    for i, child in enumerate(body):
        tag = child.tag.split('}')[-1]
        print(f"Child {i}: {tag}")
        if tag == 'tbl':
            rows = child.xpath('.//w:tr')
            print(f"  Table rows: {len(rows)}")
            # Check for gridCol
            grid = child.xpath('.//w:tblGrid/w:gridCol')
            print(f"  Table grid columns: {len(grid)}")
        elif tag == 'p':
            print(f"  Paragraph text: '{child.text if hasattr(child, 'text') else '???'}'")

if __name__ == "__main__":
    inspect_structure("number.docx")
