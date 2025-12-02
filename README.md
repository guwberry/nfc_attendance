# NFC Attendance System

## Introduction
NFC Attendance System is a web-based platform for educational institutions to manage teacher attendance using NFC card scans. It provides real-time tracking, analytics, reporting, and a modern dark-themed interface. Features include secure login, teacher registration, attendance scanning, data export, and Telegram reporting.

## Features
- NFC card-based attendance (Clock In/Out)
- Time-based attendance rules
- Dashboard with stats and recent records
- Attendance analytics and charts
- Teacher management (add/edit/delete)
- Export attendance to Excel
- Telegram integration for daily reports
- Auto clock-out after 6 PM
- Secure login for admin
- Responsive, mobile-friendly UI

## Installation
### Prerequisites
- Python 3.8+
- pip
- Node.js and npm (for frontend, if needed)

### Steps
1. **Clone the Repository**
   ```sh git clone https://github.com/guwberry/nfc_attendance.git
   cd nfc_attendance ```
```
2. **Set Up Python Environment**
   ```sh
python -m venv venv
venv\Scripts\activate  # Windows
# Or: source venv/bin/activate  # Linux/Mac
```
3. **Install Python Dependencies**
   ```sh
pip install -r requirements.txt
```
   *If missing, install manually:*
   ```sh
pip install flask pandas openpyxl aiohttp asgiref
```
4. **Configure Telegram Bot**
   - Set your Telegram Bot Token and Chat ID in the config section of the code.
5. **Initialize the Database**
   - The database (`attendance.db`) is auto-created on first run.
6. **Run the Application**
   ```sh
python app.py
```
   Or for ASGI:
   ```sh
uvicorn app:asgi_app --host 0.0.0.0 --port 5000 --reload
```

## Usage
1. Login at `http://localhost:5000` (default user: `tatbeng`, password: `aee6060`)
2. Use dashboard for stats and attendance
3. Scan NFC cards for attendance
4. Register teachers via web UI
5. Export attendance to Excel
6. Send Telegram reports
7. Auto clock-out after 6 PM

## Dependencies
- Flask, pandas, openpyxl, aiohttp, asgiref
- Bootstrap 5, Chart.js, Material Design Icons
- SQLite
- Uvicorn (for ASGI)

## License
MIT License
