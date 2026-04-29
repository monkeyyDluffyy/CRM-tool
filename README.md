# CKInternal CRM

CKInternal CRM is a lightweight, incredibly fast Customer Relationship Management application that uses a robust **Excel Spreadsheet** as its underlying database. This allows for unparalleled flexibility without the complexity of SQL or cloud database management! 

The application utilizes a **Python Flask API** to seamlessly parse, interact with, and save your CRM records natively back to the Excel sheets, feeding the data to a responsive **Vanilla JS and dark-theme CSS Dashboard**.

---

## ⚡ Key Features
* **Zero Database Overhead**: Simply manage an Excel spreadsheet `CRM_Sample_Data_Template.xlsx` directly. 
* **Dynamic Analytics**: Live tracking of Monthly Recurring Revenue (MRR), Annual Recurring Revenue (ARR), Pipeline Deals, Health Scores, and CSAT metrics.
* **Auto-Sync functionality**: Edits from the Dashboard automatically alter the `xlsx` data natively.
* **Live Polling Updates**: Editing the underlying Excel file while the app is running forces visual sync loops natively into the dashboard UI!
* **Easy Cross-Platform Export**: Setup scripts provided to natively host out of both Linux OS and Microsoft Windows.

---

## 🚀 How to Run the Application

### 1. The Normal Way (For Developers)
Make sure you have Python 3.8+ installed.
1. Create and authenticate to your virtual environment (if using one).
2. Install the lightweight Python requirements via `pip install -r requirements.txt`.
3. Open your terminal natively into the folder and type:
   ```bash
   python3 app.py
   ```
4. Open a browser and visit **http://localhost:5000** to use the CRM.

> **Optional**: Provide custom Excel databases by loading `python3 app.py "/path/to/data.xlsx"`

---

### 2. Fast Launch (For Users)
Don't want to type terminal commands? We created convenience wrappers:

* **Windows Users**: Just double-click `start.bat`. It will dynamically set up your virtual Python environment, install everything invisibly, and run the server immediately!
* **Mac/Linux Users**: Run `./start.sh` from your terminal to fast-launch your backend.

---

### 3. Professional "Always-On" Setup
Looking to keep this natively active in the background forever on your PC without annoying popups?

#### Linux Setup (`systemd`)
For native Linux environments, you successfully bypass terminals by embedding the engine into background services.
1. From the primary repository folder, run:
   ```bash
   sudo ./install_linux_service.sh
   ```
2. The UI is completely mapped to your system startup. Your Linux machine will quietly boot `localhost:5000` internally every time your computer turns on moving forward!

#### Windows Setup 
1. Avoid the terminal by double-clicking `run_hidden.vbs`. It runs your internal architecture invisibly allowing `localhost:5000` to be free of clunky command interfaces!
2. You can drag a shortcut of `run_hidden.vbs` directly into your `shell:startup` folder so that Windows boots it up invisibly the identical moment you power up your computer!

---

## 🗂 Database Schema Guide (Excel Instructions)
The native architectural flow uses four key relational sheets inside the `CRM_Sample_Data_Template.xlsx` database. Please **always leave row 1 as Headers**, and row 2 natively incorporates "AUTO" instructional data variables that the system inherently ignores.

1. **Customers**: The master core of operations. You assign `Customer Tier` levels, Location, IDs, and financial tracking metrics (MRR/ARR/etc.).
2. **Interactions**: Allows you to natively link outbound interactions, such as Zoom Meetings, Phone Calls, and Emails against precise Customer ID mapping!
3. **Deals**: The core Sales Pipeline tracker mapping probabilities against Expected Close Dates!
4. **Support Tickets**: CSAT resolution tracking, assigned to categories to gauge native product health and API downtime issues!

*Note: The frontend architecture natively features "Safe-Float" constraints, meaning if your user edits an Excel document row maliciously by typing characters instead of numbers for pricing tables, the server will intentionally parse that bad data to $0 instead of crashing!*
