# Doc Sign

Doc Sign is a local Windows document-signing companion. A Windows laptop opens a PDF or Word document and a phone signs over the same local Wi-Fi network using a QR code, session code, and PIN.

## Features

- Open and preview PDF documents.
- Navigate PDF pages and click the exact signature position.
- Resize the signature before saving.
- Save a signed PDF copy without overwriting the original.
- Open Word `.docx` documents.
- Insert a received signature at the current Microsoft Word cursor position when desktop Word is open.
- Save a separate signed Word copy.
- QR pairing, session code, and PIN.
- Signature transfer stays on the local network.
- Signature is held in memory for the active session.
- Uses TCP port 8765.

## Windows builds

GitHub Actions produces two artifacts:

- `Doc-Sign-Portable` containing `Doc Sign.exe`
- `Doc-Sign-Installer` containing `Doc Sign-Setup.exe`

The installer creates Desktop and Start Menu shortcuts and adds a Windows Firewall rule for private networks on TCP port 8765.

## Phone connection

The laptop and phone must be on the same Wi-Fi/LAN. Start Doc Sign, scan the QR code, confirm the session code/PIN, sign on the phone, and send the signature to the laptop.
