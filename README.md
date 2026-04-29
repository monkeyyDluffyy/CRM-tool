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


🪟 For Windows Customers
Windows makes it very easy to run the app in the background silently. Your project has a run_hidden.vbs script that avoids the black terminal box and boots the server invisibly.

Setup Steps for the Customer:

Send them the complete folder of your application.
Tell them to double-click start.bat at least once just to make sure it runs (it will install Python dependencies automatically).
Once confirmed it works, tell them to press Win + R on their keyboard, type shell:startup, and press Enter. This will open their Windows Startup folder.
Go to the CRM application folder, right-click on run_hidden.vbs, and select Create Shortcut.
Drag and drop that newly created shortcut into the shell:startup folder that was opened in Step 3.
Result: Every time the customer turns on their computer, Windows will automatically boot up the CRM server invisibly in the background. They can simply bookmark http://localhost:5000 in their browser and it will always work.

🐧 For Linux Customers
For Linux, the best paradigm is using systemd to run it as a robust background service that automatically revives itself and starts on boot. Your install_linux_service.sh script handles this perfectly.

Setup Steps for the Customer:

Have them open a terminal and navigate to the application folder.
First, they should make sure the script is executable (if it isn't already):
bash
chmod +x install_linux_service.sh start.sh
Run the installer script with root permissions:
bash
sudo ./install_linux_service.sh
Result: The script automatically generates a nexuscrm.service, enables it on boot, and starts it. The backend is now fully embedded in their OS. They can check it anytime by going to http://localhost:5000.

Note: They can view logs using sudo journalctl -u nexuscrm -f or restart it utilizing sudo systemctl restart nexuscrm if they ever need to.

🍎 For Mac Customers
Mac users can use launchd to achieve a similar permanent background setup. Since you don't have a script for Mac yet, here is what they can do:

Setup Steps for the Customer:

Open a terminal and run nano ~/Library/LaunchAgents/com.ckinternal.crm.plist
Add the following XML (replacing TARGET_PATH with the actual path to the folder):
xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple Computer//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.ckinternal.crm</string>
    <key>ProgramArguments</key>
    <array>
        <string>/bin/bash</string>
        <string>TARGET_PATH/start.sh</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>WorkingDirectory</key>
    <string>TARGET_PATH</string>
</dict>
</plist>
Save the file (Ctrl + O, Enter, Ctrl + X).
Load the script so it turns on: launchctl load ~/Library/LaunchAgents/com.ckinternal.crm.plist
Important Advice for Distribution
Make sure your customers have Python 3.8+ installed before they run these steps because your application relies on it to build the virtual environment (venv). It's highly recommended to instruct them to check the "Add Python to PATH" box if they install it freshly on Windows.


*Note: The frontend architecture natively features "Safe-Float" constraints, meaning if your user edits an Excel document row maliciously by typing characters instead of numbers for pricing tables, the server will intentionally parse that bad data to $0 instead of crashing!*
