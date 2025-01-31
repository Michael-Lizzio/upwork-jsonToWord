import json
import yaml
from docx import Document

# Load static text from YAML
with open("static_text.yaml", "r") as yaml_file:
    static_text = yaml.safe_load(yaml_file)

# Load JSON data
with open("example.json", "r") as json_file:
    data = json.load(json_file)

# Create a new Word document
doc = Document()


# Function to create a table with data
def create_table(doc, title, headers, rows):
    doc.add_paragraph(title, style="Heading 2")
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.style = "Table Grid"

    # Add headers
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header

    # Add row data
    for row_idx, row_data in enumerate(rows, start=1):
        row_cells = table.rows[row_idx].cells
        for col_idx, value in enumerate(row_data):
            row_cells[col_idx].text = str(value)

    doc.add_paragraph("\n")


# Define the structure of each segment
segment_definitions = {
    "segment1": ["name", "value", "description"],
    "segment2": ["name", "value", "description"],
    "segment3": ["ID", "Score", "Category"],
    "segment4": ["Product", "Price", "Stock"],
    "segment5": ["Location", "Status", "Last Updated"]
}

# Generate document content
for segment, headers in segment_definitions.items():
    if segment in data:
        doc.add_heading(static_text["segments"][segment], level=1)

        for item in data[segment]:
            for table_title in static_text["tables"].values():
                print(item, headers)
                rows = [[item[key] for key in headers]]  # Extract row values
                create_table(doc, table_title, headers, rows)

# Save the document
output_file = "output.docx"
doc.save(output_file)
print(f"Word document '{output_file}' created successfully!")
