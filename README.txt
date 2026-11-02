SAARIC ÉLAN LOCAL BILLING SYSTEM

Requirements: Python 3.10+.

Windows setup:
1. Extract this ZIP.
2. Open the extracted folder in File Explorer.
3. In the folder address bar type cmd and press Enter.
4. Run:
   py -m pip install -r requirements.txt
   py app.py
5. Open http://127.0.0.1:5000 in your browser.

Admin login:
Username: admin
Password: saaric123

Change the default password in app.py before using it for anything beyond a local demo.
The app stores data in saaric.db in this folder. Keep a backup of that file.

Features: admin login, product catalog, customer billing, automatic bill numbers,
bill history, print / save as PDF. This is a simple local-only starter system.
