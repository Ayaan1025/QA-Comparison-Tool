## ✅ Excel Comparison Tool

This script compares **5 specific columns** between two Excel files:

- `report.xlsx` (with scores like `"5. Fully Compliant (100%)"`)
- `validated.xlsx` (with plain numeric scores like `100`)

It generates a report showing **differences** between the two, based on `Baseline Unique ID`.

---

### 📂 File Structure

```
project-folder/
├── input/
│   ├── report.xlsx
│   ├── validated.xlsx
├── main.py
├── banner.py (optional)
├── comparison_results.xlsx (output)
├──requirements.txt
```

---

### 📋 Columns Compared

| Column in `report.xlsx`       | Column in `validated.xlsx`        | Label         |
|-------------------------------|------------------------------------|----------------|
| Maturity - Policy             | Adjusted Score - Policy            | Policy         |
| Maturity - Procedure          | Adjusted Score - Procedure         | Procedure      |
| Maturity - Implementation     | Adjusted Score - Implementation    | Implementation |
| Maturity - Measured           | Adjusted Score - Measured          | Measured       |
| Maturity - Managed            | Adjusted Score - Managed           | Managed        |

---

### 🛠 How It Works

1. Reads `report.xlsx`, one sheet at a time.
2. Extracts numbers from strings like `"5. Fully Compliant (100%)"` → `100`.
3. Reads `validated.xlsx` and grabs raw numeric values.
4. Compares values **only if `Baseline Unique ID` exists in both files.**
5. If values differ, logs them.
6. Outputs results to `comparison_results.xlsx`.

---

### 📦 Output: `comparison_results.xlsx`

Contains only rows with mismatched values:

| Sheet Name | Unique ID | Column        | Report Value | Validated Value |
|------------|-----------|---------------|---------------|------------------|
| Sheet1     | abc-123   | Policy        | 80            | 100              |
| Sheet2     | xyz-456   | Implementation| 90            | 70               |

---

### ⚠️ Important Notes

- Always double-check the output of the script.  
  - Scores marked as **`N/A`**, **`False`**, **`Inheritance`**, or values that simply **don’t exist** may be converted to **`0`** or **`100`** in the results.
  - 🔍 These should be reviewed manually to confirm accuracy.
- The script **only compares rows that exist in both files** based on `"Baseline Unique ID"`.
- If values are **not shown** in the output, it means they **matched** between the two files.

---

### 🧰 How to Get Set Up

1. **Install [Python](https://www.python.org/downloads/)**  
   - Use version **3.9.5 or later**  
   - ✅ Be sure to check **“Add Python to PATH”** during installation

2. **Install [Visual Studio Code](https://code.visualstudio.com/)**

3. **Open VSCode and install the following extensions**:  
   - `ms-python.python`  
   - *(Optional)* `better-comments`, `vscode-icons` for enhanced visuals

4. **Download this repository from Bitbucket to your computer**  
   - Clone or download the repo  
   - If downloaded as a ZIP file, **unzip** the folder

5. **Create a `requirements.txt` file in the project folder**  
   Add the following lines to it:

6. **Install the required Python packages**  
   Open a terminal in the project folder and run:

   ```bash
   pip install -r requirements.txt

---

### 🚀 How to Run

2. **Get the required Excel files:**
   - Download the **Validated Workbook** from **SharePoint**.
   - Download the **Report** from **MyCSF** under the **Analytics** tab (choose the **Report (Column)** option).
   - Place both files in the `input/` folder and rename accordingly:
     - `input/report.xlsx`
     - `input/validated.xlsx`

3. Run the script:

```bash
python main.py


3. Check the output in:

```
comparison_results.xlsx
```

---

### 🧩 Optional

- Add a `banner.py` file if you'd like to show a custom ASCII header/logo.
- Modify the column pairs in `main.py` if your workbook structure changes.

