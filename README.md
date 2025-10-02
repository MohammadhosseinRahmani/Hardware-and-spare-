# 🖥️ Hardware & Spare Pending Cases Report Automation

## 📋 Overview

This Python script automates the generation of a **Hardware & Spare Pending Cases Report** by:

1. 📥 Downloading the latest Excel attachment from a specific Exchange email folder.
2. 🧹 Cleaning and processing the ticket data.
3. 🔗 Merging with network data from a CSV link.
4. 📊 Creating a formatted Excel report with tables and a pivot table dashboard.
5. ✉️ Preparing an Outlook email with the dashboard included as an HTML table and attaching the Excel file.

This tool helps teams monitor and manage pending hardware and spare part cases efficiently.

---

## ⚙️ Prerequisites

* **Python 3.8+**
* **Libraries**: Install via pip

  ```bash
  pip install exchangelib pandas openpyxl numpy pywin32
  ```
* **Microsoft Outlook** installed (for sending emails via `win32com`).

---

## 📂 Folder Structure

* **Download/** – Stores Excel attachments downloaded from Exchange.
* **Attached/** – Stores the final report to be attached to the Outlook email.
* **Script File/** – The Python script itself.

Example:

```
project_folder/
│
├─ Download/        📥
├─ Attached/        📎
├─ hardware_report.py
```

---

## 🔧 Configuration

1. **✉️ Exchange Email**:
   Edit the script and replace:

   ```python
   username='Email'
   password='Password'
   primary_smtp_address='Email'
   ```

2. **📁 Email Folder**:
   Set the Exchange folder containing the Excel attachments:

   ```python
   TDS_Folder = account.inbox / 'Folder Name'
   ```

3. **🌐 CSV Link**:
   Replace with the actual network data CSV URL:

   ```python
   url = 'Network data CSV Link'
   ```

4. **📧 Email Recipients**:
   Replace `To` and `CC` emails in the `Email_preparation()` function:

   ```python
   mail.To = 'Email TO'
   mail.CC = 'Email CC'
   ```

---

## 🏃 Usage

1. **Run the Script**:

   ```bash
   python hardware_report.py
   ```

2. **Check Output**:

   * 📊 Excel report will be saved in the `Attached/` folder.
   * 📈 Pivot table dashboard will be included in the report.
   * ✉️ A new Outlook email will open with the HTML table and attached Excel file for review before sending.

---

## ⭐ Features

* ✅ Automatically filters tickets related to **hardware, spare parts, or vandalism**.
* ⏱️ Calculates MTTR (Mean Time to Repair) in `H:M` format.
* 🔧 Determines site vendor automatically based on network technology.
* 📊 Generates a pivot table dashboard grouped by **Region**, **Owner**, and **Ticket Title**.
* ✉️ Prepares an HTML email with the pivot table included for easy review.

---

## ⚠️ Notes

* Ensure **Outlook** is running; otherwise, email creation may fail.
* The script does **not** send the email automatically; it opens it for review.
* Compatible with both **local Exchange server** and **Office 365**.
* All paths are relative to the script folder for portability.

---

## 🖊️ Author

Mohammad Rahmani
📧 Mhrs1995@gmail.com
