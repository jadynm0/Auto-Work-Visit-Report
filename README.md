# 📊 Monthly Market Visit Performance Report Generator

An automated utility built with **Python**, **Streamlit**, and **pandas** that instantly converts raw field-worker survey datasets (`.xlsx`) into structured, executive-ready market performance reports.

This tool eliminates manual data auditing, tracks monthly KPI targets dynamically across retail channels, calculates SKU on-shelf availability, and operates completely locally for maximum data privacy and speed.

---

## 🚀 Key Benefits

* **⚡ Speed:** Cuts processing time from hours of manual spreadsheet auditing down to under 2 seconds.
* **🔒 Security & Privacy:** Operates 100% locally or on your internal network. Internal store data, photos, and sales metrics never leave your machine.
* **🎯 Accuracy:** Automatically parses text responses (e.g., *In Stock*, *Out of Stock*, *Label missing*) to dynamically compute exact coverage rates without human tallying errors.
* **⚙️ Decoupled Configuration:** Uses a dual-sheet architecture to pull target goals dynamically from a `Targets` tab—allowing non-technical users to adjust targets without modifying Python code.

---

## 📂 Repository & Folder Structure

To run the report generator, keep your project folder structured as follows:

    AUTO WORK VISIT REPORT/
    │
    ├── app.py                # Main Streamlit web application interface
    ├── generate_report.py    # Standalone command-line / backend automation script
    ├── .gitignore            # Version control exclusions
    ├── README.md             # Complete project documentation
    └── 行場計劃202606.xlsx    # [Input] Monthly market visit workbook

---

## 📊 Excel File Format Setup

The input workbook (`.xlsx`) must contain two specific sheets:

### Sheet 1: `表格回應 1`
The raw survey response tab populated directly by field employees visiting retail stores.

### Sheet 2: `Targets`
A simple two-column table defining the monthly store audit goals for each retail channel:

| Channel | Target |
| :--- | :--- |
| 7-11 | 100 |
| Circle K | 40 |
| SMKT | 40 |
| Min.Chain | 19 |
| 佳寶 | 20 |

> 🍎 **Note for macOS / Apple Numbers Users:** If you edit or create the `Targets` sheet using Apple Numbers, you **must** export the workbook back to Excel format:
> **File → Export To → Excel... → Check "One per sheet" → Save as .xlsx**.

---

## ⚙️ One-Time Setup & Installation Instructions

Follow the instructions below for your operating system before running the tool for the first time.

### 🪟 For Windows Users

#### Step 1: Install Python 3
1. Download the official installer from https://www.python.org/downloads/windows/
2. Run the installer file.
3. **🔥 CRUCIAL STEP:** Check the box at the bottom that says **"Add Python.exe to PATH"** (or **"Add Python to PATH"**) before clicking Install!
4. Click **Install Now**.

#### Step 2: Install Required Libraries
1. Open **Command Prompt** (search `cmd` in Start Menu).
2. Copy and run the following command:
    pip install streamlit pandas openpyxl

---

### 🍏 For macOS Users

#### Step 1: Install Python 3
1. Download the installer from https://www.python.org/downloads/macos/
2. Open the `.pkg` file and follow the standard installation prompts.

#### Step 2: Install Required Libraries
1. Open your Mac's **Terminal** (`Cmd + Space`, type `Terminal`).
2. Copy and run the following command:
    python3 -m pip install streamlit pandas openpyxl

---

## 🔄 How to Run the Tool

You can run this application either through an interactive web browser interface or directly via terminal commands.

### Method 1: Interactive Web App (Recommended for Chrome / Safari)

1. Open **Terminal** (macOS) or **Command Prompt** (Windows).
2. Navigate to your project folder:
    cd path/to/AUTO WORK VISIT REPORT
3. Launch Streamlit:
    streamlit run app.py
4. A web browser window will automatically open at `http://localhost:8501`.
5. Drag and drop your monthly `.xlsx` file into the upload box.
6. The app will immediately render:
   * District coverage totals and total store visit counts.
   * Target vs. Actual performance flags (🟢 Target Met / ❌ MISSED TARGET).
   * Detailed SKU on-shelf availability tables.
7. Click **"Download Report (.txt)"** or copy the output text directly into your email update or slide deck!

---

### Method 2: Command Line Interface (CLI Script)

If you prefer running a terminal script that automatically scans for the latest `.xlsx` file in the directory:

    python generate_report.py
    (On Mac, use: python3 generate_report.py)

The script will automatically detect the newest `.xlsx` file, calculate metrics, and output the report text straight to your terminal screen.

---

## 🛠️ Tech Stack & Dependencies

* **Language:** Python 3.9+
* **Web UI Framework:** Streamlit
* **Data Processing Engine:** pandas
* **Excel Reader Engine:** openpyxl
