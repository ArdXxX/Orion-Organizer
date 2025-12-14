# Orion Organizer

A Windows desktop utility that **monitors Orion clients**, **auto-restarts them on crash/closure**, and **keeps your window layout organized**. It also supports **per-client scheduling** (startup/shutdown at specific hours), profile-based configs, theming, and localization.

> This project is designed for Windows environments where Orion clients run as `OrionUO.exe` and can be launched via `OrionLauncher.exe`.

---

## Key Features

- **Client watchdog / auto-restart**: monitors configured clients and relaunches them if they disappear or crash.
- **Double-check relaunch logic**: optional confirmation timer before relaunch to avoid false positives.
- **Per-client schedule (HH:MM)**: configure **Startup** and **Shutdown** windows individually.
- **Schedule enforcement**: closes clients during the “off” window and restarts them when allowed.
- **Window layout manager**: save and restore each client’s **X, Y, W, H**.
- **Bring-to-front + re-layout**: focuses all saved clients and then reapplies the saved layout.
- **Auto-reposition after launch**: restores layout automatically after a configurable delay.
- **Launch queue with throttling**: relaunches are queued with a delay between starts to reduce load spikes.
- **Disconnect handling**: configurable list of disconnect titles; matching clients are terminated.
- **Error-window killer**: detects windows containing “Error” in the title and terminates the owning process.
- **Multi-profile support**: isolated profiles with independent settings (`perfiles/`).
- **Localization-ready**: JSON language packs in `languages/`, English baseline included.
- **Dark/Light themes** (CustomTkinter).

---

## Screenshots

<img width="994" height="651" alt="image" src="https://github.com/user-attachments/assets/4f7452c5-cb7e-4cad-81f0-a043345a95e0" />

<img width="417" height="449" alt="image" src="https://github.com/user-attachments/assets/c91b0c52-59f8-4927-bfcc-d96d05a1d34b" />

<img width="512" height="483" alt="{4246FF33-C572-4191-984C-B56A5FEB20AA}" src="https://github.com/user-attachments/assets/cbd73fb3-2391-4785-989a-860900800156" />

<img width="513" height="487" alt="{7902D2C9-8971-4E8C-9FCC-3D675D2ECD4A}" src="https://github.com/user-attachments/assets/394560fb-0a73-4d42-9307-9b6487636054" />

<img width="516" height="487" alt="{0DF33946-F0B1-45DD-98C6-45328E997683}" src="https://github.com/user-attachments/assets/df50090a-e61f-4121-9755-e06c1d2d2ce6" />



## Requirements

- **Windows 10/11**
- **Python 3.10+** (recommended)
- Orion installation folder containing:
  - `OrionLauncher.exe`
  - Orion clients running as `OrionUO.exe`

---

## Dependencies

Core Python packages:

- `customtkinter`
- `pygetwindow`
- `psutil`
- `pywin32`

---

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/ArdXxX/Orion-Organizer.git
   cd Orion-Monitor

2. **Install dependencies**
   ```bash
   pip install -U pip
   pip install customtkinter pygetwindow psutil pywin32


## Run

### Option A: Double-click (recommended)

If you have Python installed, simply **double-click the `.pyw` file** to run the app.

- The `.pyw` extension runs the script **without opening a console window**.

### Option B: Run from terminal

```bash
python orion_organizer.pyw

Replace `orion_organizer.pyw` with the actual filename.

---

## Notes

- **Windows-only**: relies on Win32 APIs (`pywin32`) and window management libraries.
- For auto-launch via Orion profiles, configure the Orion folder so the tool can find `OrionLauncher.exe`.



