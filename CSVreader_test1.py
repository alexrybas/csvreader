import csv
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def read_csv(csv_path):
    """Reads a CSV file and returns the data as a list of lists."""
    with open(csv_path, "r", newline="", encoding="utf-8") as csvfile:
        reader = csv.reader(csvfile, delimiter=';', quoting=csv.QUOTE_MINIMAL)
        data = list(reader)
    
    max_cols = max(len(row) for row in data)  # Find max columns
    data = [row + [""] * (max_cols - len(row)) for row in data]  # Fill missing columns
    
    return data

def set_table_borders(table):
    """Applies GOST-style borders to a Word table."""
    tbl = table._element
    tbl_pr = tbl.find(qn('w:tblPr'))
    if tbl_pr is None:
        tbl_pr = OxmlElement('w:tblPr')
        tbl.append(tbl_pr)
    
    borders = OxmlElement('w:tblBorders')

    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')  # 'single' is a thin line, use 'thick' for a bold line
        border.set(qn('w:sz'), '8')  # Border thickness (8 = 0.5pt)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')  # Black border
        borders.append(border)
    
    tbl_pr.append(borders)

def create_gost_word(csv_path, output_path):
    """Creates a Word document with a table formatted according to GOST standards."""
    data = read_csv(csv_path)
    doc = Document()
    table = doc.add_table(rows=len(data), cols=len(data[0]))  # Ensure enough columns
    
    for row_idx, row in enumerate(data):
        for col_idx, value in enumerate(row):
            table.cell(row_idx, col_idx).text = value
    
    set_table_borders(table)
    doc.save(output_path)
    print(f"Document saved to {output_path}")

if __name__ == "__main__":
    csv_file = "CSV_Test.csv"  # Change to the actual path if needed
    word_file = "output.docx"
    create_gost_word(csv_file, word_file)
