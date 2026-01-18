# 📘 CDMP Quiz Application – Clean & Optimized

A streamlined desktop quiz application built with **Python + Tkinter** for practicing the  
**Certified Data Management Professional (CDMP)** exam.

The application runs fully offline and loads questions directly from an Excel file.

---

## 👤 Author

**Mostafa Maher**  
GitHub: https://github.com/FINAL094  
LinkedIn: https://www.linkedin.com/in/eng-mostafa-maher  

---

## ✨ Features

- 📚 Chapter-based question selection
- 🔀 Optional randomization of questions and answers
- ⏱️ Exam-style countdown timer
- ✅ Automatic scoring (single & multiple correct answers)
- 🔍 Review mode with correct answers and references
- ⌨️ Keyboard shortcuts for answer selection (1–9)
- 🕌 Islamic greeting and respectful user interface
- 📄 Excel-based question source
- 🖥️ Fully offline desktop application

---

## 🛠️ Requirements

- Python **3.9+** (tested up to Python 3.13)
- Required Python packages:

pip install pandas openpyxl

---

## 📂 Project Structure (IMPORTANT)

This application expects **both files in the SAME directory**.

cdmp-quiz/

├── cdmp_quiz.py

└── CDMP Practice Exam.xlsx

✔ Do NOT place files in subfolders  
✔ File name must match **exactly**

---

## 🚀 How to Run

### 1️⃣ Clone or Download the Repository

git clone https://github.com/FINAL094/cdmp-quiz.git

Or download the ZIP file and extract it.

---

### 2️⃣ Navigate to the Application Folder

cd cdmp-quiz

Make sure both `cdmp_quiz.py` and `CDMP Practice Exam.xlsx` are in this folder.

---

### 3️⃣ Install Dependencies (One Time Only)

pip install pandas openpyxl

---

### 4️⃣ Run the Application

python cdmp_quiz.py

The quiz window will open immediately.

ℹ️ Note  
The application automatically sets its working directory to the script location to ensure the Excel file is always found correctly, even when launched from an IDE or a different terminal location.

---

## 📄 Excel File Format

The application supports **TWO Excel formats**.

---

### ✅ Option 1: Two-Sheet Format

Your Excel file may contain two sheets:

- Sheet `ques` → Questions
- Sheet `ans` → Answers

This format allows advanced customization and scoring.

---

### ✅ Option 2: Single-Sheet CDMP Format (Most Common)

A single sheet with the following columns:

- Question Number
- Knowledge Area
- Question
- A
- B
- C
- D
- E
- Correct

Example:

Question Number | Knowledge Area | Question | A | B | C | D | Correct  
Q1 | Data Governance | What is data stewardship? | Option A | Option B | Option C | Option D | B  

✔ Matches the standard CDMP mock exam layout  
✔ Multiple correct answers supported (e.g. A,C)

---

## 🔍 Review Mode

Review Mode becomes available when:

- All questions are completed  
OR  
- The exam timer expires  

In Review Mode, you can see:

- Your selected answers
- Correct answers
- References (if available)

Click **“Review Exam”** when it becomes enabled.

---

## 🆘 Troubleshooting

### ❌ Excel file not found

Make sure that:
- The file name is exactly: CDMP Practice Exam.xlsx
- The file is in the same folder as cdmp_quiz.py
- You are running the script from that folder

---

### ❌ Excel format not supported

This error means the Excel file does not match a supported structure.

Ensure that:
- The file follows one of the two formats described above
- Column names are spelled correctly
- At least one question and answer exist

---

## 🤲 Acknowledgment

بِسْمِ اللهِ الرَّحْمنِ الرَّحِيمِ

اللهم صل على محمد وعلى آل محمد كما صليت على إبراهيم وعلى آل إبراهيم، إنك حميد مجيد، اللهم بارك على محمد وعلى آل محمد كما باركت على إبراهيم وعلى آل إبراهيم، إنك حميد مجيد ﷺ  

If this application benefits you, please remember the author in your Prayer.

---

## 📜 License

This project is provided for **educational and personal practice use only**.  
All CDMP-related content remains the property of its respective owners.
