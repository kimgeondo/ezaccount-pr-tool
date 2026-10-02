# EZAccount access guide

## Local access

Run the app on the PC:

```powershell
cd C:\ezaccount
.\.venv\Scripts\Activate.ps1
python app.py
```

Open on the same PC:

```text
http://127.0.0.1:5000/
```

## Private device access

Do not expose this prototype directly to the public internet. It can create purchasing requests and currently has no user authentication. For phone access, use an approved company VPN/private network and add authentication before making the app reachable to other devices.

## One-click startup

Double-click:

```text
start_app.bat
```

This will activate the venv and run the app.
