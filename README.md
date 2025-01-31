# JSON to Word Document Transformer

## 📌 Overview  
This Python script converts structured **JSON data** into a formatted **Word document (`.docx`)**, generating five distinct **segments** based on object types. Each segment contains **three tables**, dynamically populated with data from the JSON. **Static text is stored in YAML**, ensuring easy modifications without changing the script.

## 🔧 Features  
✅ **Dynamic JSON Parsing** – Automatically adapts to input data without hardcoding.  
✅ **Fully Generated Word Document** – No templates required, ensuring flexibility.  
✅ **YAML Configuration** – Static text (segment headers, table titles) is externalized.  
✅ **Scalable & Modular Design** – Easily extendable for new object types and formats.  
✅ **Unit Testing Included** – Ensures data is correctly processed and output matches expectations.  

---

## 📂 Project Structure  
```
/json_to_docx
│── script.py         # Main script to generate the Word document
│── static_text.yaml  # Stores static text (segment names, table titles)
│── example.json      # Sample JSON input data
│── test_script.py    # Unit tests to validate output
│── README.md         # Project documentation
```

---

## ⚙️ Installation & Setup  

### 1️⃣ Clone the Repository  
```sh
git clone https://github.com/yourusername/json-to-docx.git
cd json-to-docx
```

### 2️⃣ Install Dependencies  
```sh
pip install python-docx pyaml
```

### 3️⃣ Run the Script  
```sh
python script.py
```
This will generate `output.docx` with structured content based on `example.json`.

---

## 💑 Input & Configuration  

### 🏋️ JSON Input (`example.json`)
```json
{
    "segment1": [
        {"name": "Object A", "value": 100, "description": "Test object"},
        {"name": "Object B", "value": 200, "description": "Another test"}
    ]
}
```

### 🗒️ YAML Static Text (`static_text.yaml`)
```yaml
segments:
  segment1: "Segment 1 Header"
  segment2: "Segment 2 Header"

tables:
  table1: "Table 1 Title"
  table2: "Table 2 Title"
  table3: "Table 3 Title"
```

---

## 🚀 Script Workflow (`script.py`)  

1️⃣ **Reads JSON data** – Loads structured data from `example.json`.  
2️⃣ **Loads static text** – Extracts segment headers and table titles from `static_text.yaml`.  
3️⃣ **Generates document** – Iterates through five segment types, creating three tables per segment.  
4️⃣ **Saves output** – Exports a `.docx` file with structured content.  

---

## 🧠 Running Unit Tests  
```sh
python -m unittest test_script.py
```
✅ **Checks if JSON data appears in the output document.**  
✅ **Ensures segment structures match expected formatting.**  

---

## 📌 Next Steps  
- 🏰️ **Enhance formatting** (custom fonts, table styles).  
- 📄 **Support images & charts in Word output.**  
- 🔄 **Allow CLI customization for file inputs/outputs.**  

🌟 **Looking forward to contributions and feedback!** 🚀
