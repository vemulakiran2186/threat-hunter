# 🛡️ Threat Hunter: CVE Intelligence Collector

Threat Hunter is a lightweight, local web dashboard designed for cybersecurity analysts. It automates scanning trending vulnerability news, threat feeds, and security advisories, searching for **CVE (Common Vulnerabilities and Exposures) IDs**, and outputting a professionally structured and styled Excel spreadsheet.

## Key Features
- 🔍 **Regex-Based CVE Detection**: Scans the text and html of provided URLs for any `CVE-YYYY-NNNN+` identifiers.
- 💡 **Smart Description Context**: Rather than just grabbing general site description headers, the tool attempts to find the exact paragraph/text block containing the specific CVE ID and uses it as the short description.
- 📂 **Multi-Input Formats**: Paste lists of URLs directly or upload a text/CSV list of targets.
- 📊 **Interactive Data Table**: View and review threat hunting findings in real-time inside the web interface.
- 📥 **Custom-Styled Excel Report**: Exports an Excel report using the openpyxl engine with proper column widths, cell wrapping, bold text, and a clean cybersecurity color palette.

---

## Getting Started

### 📋 Prerequisites
- Python 3.8 or higher installed.

### ⚙️ Setup Instructions
1. Open your terminal or Command Prompt and navigate to the application folder:
   ```powershell
   cd f:\Antigravity\threat-hunter
   ```

2. (Optional but Recommended) Create and activate a Python virtual environment:
   ```powershell
   python -m venv .venv
   .venv\Scripts\Activate.ps1
   ```

3. Install the required libraries:
   ```powershell
   pip install -r requirements.txt
   ```

4. Launch the Streamlit dashboard:
   ```powershell
   streamlit run app.py
   ```

---

## Using the Dashboard
1. Paste security URLs into the input box (e.g. CISA Advisories or security news links).
2. Click **Start Threat Hunt**.
3. Watch the progress bar and real-time logs fetch the data.
4. Preview the results in the web dashboard.
5. Click **Download Excel Threat Report** to export the structured spreadsheet report!
