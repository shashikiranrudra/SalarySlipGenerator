# 💼 Salary Slip Generator
### Professional PDF Salary Slip Generator from Excel Sheets

---

## 🚀 Quick Start

### Windows
1. Double-click **`install_and_run.bat`** (first time only — installs dependencies)
2. After that, use **`run.bat`** to launch
3. Login with default password: **`admin123`**

### Mac / Linux
```bash
chmod +x install_and_run.sh
./install_and_run.sh
```

---

## 📋 Requirements
- Python 3.8 or newer (https://python.org)
- Internet connection for first install (to download packages)
- Packages (auto-installed): `openpyxl`, `pandas`, `Pillow`, `reportlab`

---

## 🔐 Password Protection
- **Default password:** `admin123`
- Change it in **Settings** tab after login
- Password is stored as a SHA-256 hash (secure, never plain text)

---

## 📊 Excel Sheet Format

Your Excel salary sheet must have **column headers in Row 1**.  
See **`SAMPLE_Salary_Sheet_Template.xlsx`** for an example.

### Default Column Names (you can remap these in the app)

| Internal Field | Default Excel Column Name |
|----------------|--------------------------|
| Employee Name | Employee Name |
| Employee Code | Employee Code |
| Department | Department |
| Designation | Designation |
| Location | Location |
| PF No | PF No |
| UAN No | UAN No |
| PAN No | PAN No |
| Aadhar No | Aadhar No |
| Bank Name | Bank Name |
| Bank A/c No | Bank A/c No |
| Monthly Salary | Monthly Salary |
| Month | Month |
| Year | Year |
| Basic | Basic |
| Conveyance | Conveyance |
| House Rent Allowance | House Rent Allowance |
| Post Allowance | Post Allowance |
| Other Allowance | Other Allowance |
| Profession Tax | Profession Tax |
| Provident Fund | Provident Fund |
| ESIC Deduction | ESIC Deduction |
| Other Deduction | Other Deduction |
| Month Days | Month Days |
| Weekly Off | Weekly Off |
| Holidays | Holidays |
| Present Days | Present Days |
| Pay Days | Pay Days |
| Leave Without Pay | Leave Without Pay |

---

## 🏢 Company Customization

In the **Company Info** tab, you can:
- ✅ Set your **Company Name**
- ✅ Set your **Company Address**
- ✅ Upload your **Company Logo** (PNG/JPG)
- ✅ Pick a custom **Accent Color** for PDF headers

---

## 📋 Column Mapping

If your Excel uses different column names:
1. Go to **Column Mapping** tab
2. Type the exact header name from your Excel file
3. Click **Save Column Mapping**

---

## 🖨️ Generating Slips

1. Go to **Generate Slips** tab
2. Click **Browse Excel** → select your salary sheet
3. Click **Load / Refresh**
4. Select employees (double-click or Space bar to toggle)
5. Set output folder
6. Click **Generate Selected PDFs** or **Generate All**

PDFs are saved as: `{EmpCode}_{Name}_{Month-Year}.pdf`

---

## 📁 Data Storage

App settings are saved at: `~/.salary_slip_app/config.json`  
This includes company info, column mappings, and hashed password.

---

## ❓ Troubleshooting

**"No module named tkinter"**  
→ Install tkinter: `sudo apt install python3-tk` (Linux)  
→ On Windows/Mac, tkinter comes with Python automatically.

**PDF not generating?**  
→ Check the output folder has write permission  
→ Ensure earnings columns have numeric values (not text)

---

*Generated slips match the RAJ PATH INFRACON format and can be customized for any company.*
