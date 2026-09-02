COMPANY DATABASE V1

This is a mobile-style web app prototype.
Fields:
- Company Name
- Company Location
- Contact Person
- Designation
- Mobile Number
- Email ID
- Domain
- Industry: IT / Electrical & Electronics / Manufacturing / Construction / Fashion / Others

Features:
- Add / Edit / Delete
- Search
- Excel export (.xlsx)
- Excel report is A4 landscape, with filters, frozen header, and print-friendly widths.

RUN:
1. Install Python 3.
2. Open a terminal in this folder.
3. Run: pip install -r requirements.txt
4. Run: python app.py
5. Open: http://127.0.0.1:5000

For phone testing on the same Wi-Fi:
- Find the computer's local IP (e.g. 192.168.1.10)
- Start the app with: python app.py
- On the phone open: http://COMPUTER-IP:5000
- Allow Python through Windows Firewall if prompted.

Data is stored in companies.json in this folder.
