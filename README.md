# Doc Sign

**By Dave Wilson**

Doc Sign is a Windows document-signing application that lets a person sign on a phone while the document is displayed on a Windows PC. The phone and computer communicate directly over the same local Wi-Fi or LAN.

## Quick start

1. Install and start **Doc Sign** on the Windows computer.
2. Make sure the computer and phone are connected to the same Wi-Fi/LAN.
3. Scan the QR code shown in Doc Sign with the phone.
4. Confirm the session code and PIN if requested.
5. Sign in the signature box on the phone and send the signature.
6. Open a PDF or Microsoft Word document in Doc Sign.
7. Click the document page where the signature must be placed.
8. Use the **-** and **+** controls to resize the signature.
9. Press **Confirm on this page & Next** to keep that placement and continue to another page.
10. Repeat on any other pages that need a signature.
11. Press **Save signed document** and choose a new file name/location.

Doc Sign does not overwrite the original document unless you deliberately choose the original file as the save destination.

## PDF documents

- Click **Open PDF**.
- Navigate pages with the left/right page buttons.
- Click anywhere on the page to position the current signature.
- Resize with the **- / +** controls.
- Confirm each required placement.
- One signature can be placed on several pages before the final document is saved.
- Save creates a signed PDF copy.

## Microsoft Word documents

Doc Sign supports common desktop Microsoft Word formats including `.doc`, `.docx`, and `.docm` when Microsoft Word is installed on the computer.

- Click **Open Word**.
- Doc Sign creates a local page preview using Microsoft Word in the background.
- Place and resize the signature directly on the page preview in Doc Sign.
- Confirm each required page placement.
- When saved, signatures are inserted as floating images positioned relative to the page and set **Behind Text**, preventing the signature from pushing document text out of position.
- The original Word document is kept unchanged when a new signed copy is saved.

## Close Document

Use **Close Document** before opening a different document when you no longer need the current one. If there are confirmed placements that have not yet been saved, Doc Sign warns before discarding them.

## Save Signature PNG

After a signature is received, click **Save Signature PNG**. The signature is saved as a transparent `.png` image which can be reused in programs such as Adobe Acrobat, Microsoft Word, email, or other document software.

## Phone connection

Doc Sign uses TCP port **8765** on the local network. The phone does not need a separate app installed; it opens the local signing page in its browser.

The installer requests administrator permission and creates a Windows Firewall inbound rule for **Doc Sign.exe** on TCP port 8765 for **Private networks**. Some third-party security products, including endpoint firewalls, may apply their own rules independently of Windows Firewall.

## Security and privacy

- Signing traffic stays on the local network.
- Pairing uses a session code and PIN.
- The signature is transferred directly from the phone to the computer.
- No cloud service is required for the signing transfer.
- The original document is not automatically overwritten.

## Troubleshooting phone connection

If the QR page will not open on the phone:

1. Confirm the phone and PC are on the same Wi-Fi/LAN.
2. Confirm the network is set as **Private** in Windows where appropriate.
3. Restart Doc Sign and scan the new QR code.
4. Check Windows Firewall or a third-party firewall/security suite if the connection is still blocked.
5. TCP port 8765 must be allowed for Doc Sign on the local network.

## Windows builds

GitHub Actions produces:

- `Doc-Sign-Portable` containing `Doc Sign.exe`
- `Doc-Sign-Installer` containing `Doc Sign-Setup.exe`

For normal use, the installer build is recommended because it creates the shortcuts and Windows Firewall rule automatically.
