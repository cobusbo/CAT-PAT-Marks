# 📊 CAT-PAT-Marks

### 🔍 Automatically extract learner marks for CAT PAT assessment

This Python script helps streamline the process of compiling **Phase 1, 2, and 3 marks** from individual learner Excel marking grids into a single **summary Excel file**.

---

## ✅ Features

- 🔄 **Auto-detect completed files**  
  Learner Excel files are renamed with a `.` at the beginning (e.g. `.Boekwurm, Bennie.xlsx`) to mark them as finished. The script automatically detects these files only.

- 🧠 **Extract learner names**  
  Strips the leading dot `.` and the `.xlsx` extension to extract clean learner names.

- 📄 **Reads marks from the `Opsomming` sheet**  
  For each completed file, it retrieves:
  - `Fase 1` → Cell **E4**
  - `Fase 2` → Cell **E5**
  - `Fase 3` → Sum of cells **E6 + E7 + E8**
  - `Total` → Cell **E10**

- 📊 **Creates a new Excel file (`summary.xlsx`)**  
  - Column 1: Learner name  
  - Column 2–4: Fase 1, 2, 3 marks  
  - Column 5: Total  
  - Entries are automatically **sorted alphabetically by learner name**
  - Header row is **bold** and columns are **auto-sized**

---

## 📁 Example Workflow

1. You mark a learner’s PAT and rename their file from:
   ```
   Boekwurm, Bennie.xlsx ➜ .Boekwurm, Bennie.xlsx
   ```

2. Run the script in the same folder.

3. A file called `summary.xlsx` is generated with:

   | Name              | Fase 1 | Fase 2 | Fase 3 | Total |
   |-------------------|--------|--------|--------|--------|
   | Boekwurm, Bennie | 25     | 23     | 47     | 95     |
   | ...               | ...    | ...    | ...    | ...    |

---

## 🛠️ Requirements

- Python 3.x
- `openpyxl` library  
  Install with:
  ```bash
  pip install openpyxl
  ```

---

## 🚀 Usage

1. Place the script in the same folder as the marked PAT Excel files.
2. Run the script:
   ```bash
   python cat_pat_marks.py
   ```
3. Open `summary.xlsx` to view all compiled marks.

---
