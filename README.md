# AI-Powered Microsoft Excel Workshop — Daily Student Report & Portal

An interactive, responsive, and editable web portal & institutional document management system built for **N.B.K.R. Institute of Science & Technology (NBKRIST)** (Autonomous • Department of IT and AI&DS).

This website digitizes the **5-Day AI-Powered Microsoft Excel Workshop: From Basics to Advanced Analytics** practical lab record and certificate, allowing full customization, interactive learning simulations, and institutional print / PDF generation.

---

## 🌟 Key Features

### 1. 🔄 Dual-Mode Architecture
- **Interactive Studio & Hub**:
  - Daily Technical Learning Logs with live inline editing.
  - Interactive Excel & AI Simulators for each workshop day.
  - Evaluation & Feedback rating selector and confidence pill toggles.
  - Expandable Reflection Q&A (all 25 questions across 5 days).
- **Institutional Record Mode (Print / PDF Ready)**:
  - Exact 12-page high-fidelity reproduction of the official document.
  - **Page 1**: Bonafide Certificate of Practical Work with NBKRIST Seal, custom academic year, and faculty signatures.
  - **Page 2**: Index Table with day-wise topics, session dates, page numbers, and verification stamps.
  - **Pages 3–12**: Days 1 through 5 Daily Student Reports with structured 5-section evaluation tables.
  - `@media print` optimized A4 pagination for direct browser printing or saving as PDF.

---

### 2. ⚡ Day-by-Day Interactive Simulators
- **Day 1: Excel Basics, Formulas & Copilot Integration**
  - Live interactive table with real-time `SUM`, `AVERAGE`, and `IF` logic calculation.
  - Microsoft Copilot AI Formula Prompt Assistant with dynamic formula generation.
- **Day 2: Power Query & Data Cleaning ETL Pipeline**
  - Interactive step-by-step pipeline cleaner (Raw Messy Data &rarr; Trim/Clean &rarr; Impute Nulls &rarr; XLOOKUP Department join).
  - Power Query M-Code syntax inspector.
- **Day 3: Pivot Tables, Pivot Charts & AI Insights**
  - Dynamic grouping by Department, Academic Year, or Grade Band with live Chart.js visualizations.
  - Automated Copilot descriptive insights box.
- **Day 4: Executive KPI & Power Pivot Dashboard**
  - Metric summary cards (Total Cohort, Pass Rate, AI Adoption, Satisfaction Score).
  - Interactive Analytics Competency vs. Completion Speed Trend chart.
- **Day 5: Office Scripts Automation Playground**
  - TypeScript code editor for Excel Online macros.
  - 1-click script execution with animated progress and execution log terminal.

---

### 3. 🛠️ Live Customization & Data Management
- **Global Profile Synchronizer**: Update Student Name, Roll Number, Branch, College, Session Dates, and Examiner names in one place; updates propagate instantly to the certificate, index, and all 5 day sheets.
- **Local Storage Persistence**: Edits are automatically saved in the browser so nothing is lost on refresh.
- **JSON Backup & Import**: Export the entire dataset as JSON, download a `.json` backup file, or upload student profiles.
- **Reset to Scanned Defaults**: Restore the original default OCR values at any time.
- **Dark / Light Theme**: Smooth theme switcher with sleek modern glassmorphism.

---

## 🚀 How to Run Locally

Simply double-click `index.html` to open in any web browser (Chrome, Edge, Firefox, Safari), or serve with any static web server:

```bash
# Optional: Using Python built-in server
python -m http.server 3000

# Open in browser:
# http://localhost:3000
```

---

## 🖨️ How to Print / Export PDF

1. Click the **"Institutional Record (Print/PDF)"** button in the top navigation bar (or press `Ctrl + P` / `Cmd + P`).
2. Set **Destination** to *"Save as PDF"* or select your printer.
3. Set **Paper Size** to `A4` and **Margins** to `Default` or `None`.
4. Ensure **Background Graphics** is checked for optimal styling of borders and seals.
5. Click **Save / Print**.
