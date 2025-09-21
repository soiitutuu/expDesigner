# expDesigner

<p align="center">
  <img src="https://i.ibb.co/392Wp49N/image.png" alt="expDesigner logo" width="128" height="128">
</p>

<p align="center">
  <b>Windows 10/11 customization utility — fast, safe(ish), and pretty.</b><br>
  Self-elevating, registry-backed, dark/light themes, and a one-click “Apply & Restart Explorer.”
</p>

<p align="center">
  <a href="https://github.com/voidwither/expDesigner/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/voidwither/expDesigner?style=for-the-badge&color=FFD700"></a>
  <a href="https://github.com/voidwither/expDesigner/issues"><img alt="GitHub issues" src="https://img.shields.io/github/issues/voidwither/expDesigner?style=for-the-badge"></a>
  <a href="https://github.com/voidwither/expDesigner/network/members"><img alt="GitHub forks" src="https://img.shields.io/github/forks/voidwither/expDesigner?style=for-the-badge"></a>
  <img alt="Python" src="https://img.shields.io/badge/Python-3.11%2B-3776AB?style=for-the-badge&logo=python&logoColor=white">
  <img alt="PyQt6" src="https://img.shields.io/badge/GUI-PyQt6-41cd52?style=for-the-badge">
  <img alt="Platform" src="https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D4?style=for-the-badge&logo=windows&logoColor=white">
  <img alt="UAC" src="https://img.shields.io/badge/Admin-UAC%20Elevated-cc0000?style=for-the-badge">
  <a href="#license"><img alt="License" src="https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge"></a>
</p>


## ✨ What is this?

expDesigner is a single-file Windows customization tool built with PyQt6. It reads your current settings from the registry, lets you tweak a ton of personalization, taskbar, Explorer, privacy, and performance options, and applies changes safely with logs and backups. It also restarts Explorer for instant shell updates.

Made to be easy to fork, easy to ship, and easy to undo.


## 🚀 Features

- Self-elevating (UAC) — prompts for admin and relaunches automatically
- Crash-safe logging to expDesigner.log with full tracebacks
- Dark/Light theme with live toggle and auto-detects Windows theme on first run
- Data-driven settings across 5 pages:
  - Personalization
  - Taskbar & Start
  - File Explorer & UI
  - Privacy
  - System & Performance
- Global search for settings
- Pending-changes model: Preview, Apply, or Revert before writing
- Per-setting reset, page reset, and “Favorites” to pin your go-tos
- Profiles: Save/Load JSON configs
- .reg export of pending changes + automatic .reg backup of previous values on Apply
- Optional create System Restore Point before Apply (PowerShell)
- Live-ish registry sync (polls and updates UI when external changes happen)
- One-click “Apply & Restart Explorer”


## 🧩 Settings buffet (highlights)

- Theme: App/System dark mode, transparency, accent color usage, taskbar animations
- Taskbar: Search style, size, alignment, combine behavior, clock seconds, Widgets/Task View
- Explorer: Hidden files, extensions, protected OS files, compact view, status bar, info tips, launch target
- Privacy: Ads ID, Activity History/Upload, telemetry level, tips/suggestions, clipboard history, location, Cortana
- Performance: Fast Startup, Hibernation, Game DVR, GPU scheduling, Visual Effects preset, NTFS tweaks, Responsiveness, menu/mouse delays

Note: Some options are version/build-dependent and may be ignored by certain Windows editions. That’s fine; they’re safe to toggle and fully logged.


## 📦 Install & Run

- Python 3.11+ recommended
- Windows 10/11

```bash
pip install PyQt6
python expDesigner.py
```

The app will prompt for admin via UAC and restart itself with elevated rights.


## 🖥️ Build a portable .exe

Using PyInstaller:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name expDesigner expDesigner.py
```

You’ll get dist/expDesigner.exe — ship it! Running it still prompts for admin automatically.


## 🛟 Safety, backups, and rollbacks

- Every Apply writes a backup .reg of the previous values to expDesigner-backup.reg
- Optional “Create restore point before Apply” (Tools menu)
- You can also Preview pending changes and export them as a .reg file
- If anything looks off, Revert All to discard pending changes before writing


## 🔎 Troubleshooting

- UAC loop: If elevation fails (you clicked “No”), the app exits. Launch again and allow.
- “Some toggles don’t stick”: Your Windows build may enforce policies. Sign out or reboot if the badge says so.
- “Explorer didn’t update”: Click “Apply & Restart Explorer” again or sign out/in.
- “I broke a thing”: Double-click expDesigner-backup.reg to restore previous values, or use your restore point.


## 🧪 Dev quickstart

- Single-file by design — fork and edit expDesigner.py directly
- UI is PyQt6 only; no external theme libs
- The app creates QApplication immediately after PyQt import to avoid widget-before-app crashes
- All registry operations are centralized in RegistryManager
- Data-driven settings make it easy to add more toggles

PRs welcome! Add new settings with a clear tooltip, registry path/value, defaults, and any “requires Explorer/logoff/reboot” notes.


## 📸 Screenshots
<img width="1282" height="852" alt="NxcGCPo" src="https://github.com/user-attachments/assets/c8d8b1e1-2dae-4ae4-b9e3-c8fb9f660c5a" />


## 🤝 Contributing

- Fork
- Make your change
- Open a PR with a clear description and screenshots if UI-related

Please keep settings safe, documented, and clearly scoped.


## ⭐ About

expDesigner | Made by VoidWither on GitHub | Feel free to use this code on your own projects!


## 📜 License

MIT. See LICENSE for details.

