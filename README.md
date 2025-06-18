# 📊 INRAE Excel Data Reader – Freelance Project

This project was developed for INRAE (French National Research Institute for Agriculture, Food and Environment) to streamline the reading, processing, and analysis of complex internal Excel spreadsheets containing financial and analytical data.

## 🧠 Objective

Automate the extraction, transformation, and querying of data from Excel files used by INRAE researchers.  
The tool enables:

- Identification of scientists linked to budget lines.
- Segmentation of data by researcher.
- Filtering and sorting by date or amount.
- Cross-referencing budgetary entries.
- Output via a simple interface (Streamlit or command-line).

---

## ⚙️ Key Features

### 🔍 Data Loading and Parsing
- Automatic loading of `Cible.xlsx`.
- Reads multiple sheets:
  - Project deadline tracking.
  - Budgetary and analytical data.

### 🧑‍🔬 Scientist Identification
- Extracts and sorts names from column `H`.
- Filters for unique scientist identifiers.

### 🗂️ Analytical Line Management
- Extracts unique analytical lines from column `L`.

### 🧾 Data Processing
- Cleans and replaces text values (`Terminé`, `PROLONGE`, etc.) with `datetime` objects or placeholders.
- Transforms blocks of raw data into structured Python dictionaries.

### 🔄 Querying System
- Enables complex queries, such as:
  - Search all entries with value "120".
  - Sort by project end date or maximum amount.

### 💻 User Interface
- Can be used as a CLI Python script or via a visual interface using Streamlit.
- Allows basic interactive input for dynamic filtering.

---

## 🧰 Tech Stack

- **Python 3**
- `pandas`, `numpy`
- `openpyxl`
- `datetime`
- `glob2`
- `streamlit` (optional)

---

## 🚀 How to Run

1. **Install dependencies**:
   ```bash
   pip install pandas numpy openpyxl streamlit glob2
Project structure:

css
Copy
Edit
.
├── Cible.xlsx
├── main.py
└── README.md
Run via Python CLI:

bash
Copy
Edit
python main.py
Or with Streamlit (optional):
Uncomment streamlit lines in the script, then run:

bash
Copy
Edit
streamlit run main.py
📌 Example Use Case
"Retrieve all rows with a budget value of 120, then sort them by project deadline and filter by highest monetary amount."

📁 Input File Requirement
The source file Cible.xlsx must include two main sheets:

Sheet 1: Project end dates (column F)

Sheet 2: Analytical and budgetary data (columns H, L, etc.)

🧼 Developer Notes
Splitting logic relies on INRAE’s internal data structure markers (e.g., "BPP").

Each block is parsed and transformed into a manipulable list or dictionary.

Streamlit compatibility allows for broader accessibility and testing.

👤 Author
Damien Lemoine
Freelance developer – Data automation & tools
Developed in collaboration with the INRAE team in Dijon, France.

📄 License
This project is proprietary and was developed under a freelance contract for INRAE.
Distribution, reproduction, or commercial use without prior written consent is prohibited.
