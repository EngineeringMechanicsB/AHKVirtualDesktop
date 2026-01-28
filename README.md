# Windows Virtual Desktop
This is a small AutoHotkey project that provides lightweight virtual desktop and window management features for Windows. It was originally built for my own daily work needs and later shared publicly in case it may be useful to others.
The script keeps things simple — no heavy UI, no installation, and no system-level modifications — just keyboard-driven utilities to manage windows more efficiently.
## ✨ Features
Virtual desktop switching (up to 9 desktops)
Move windows between desktops
Pie menu for directional actions
Status bar with time and work progress display
Auto-generated configuration file
Clipboard recovery output
Optional Vim floating window integration
Optional terminal launcher
Minimal footprint, no background services
## Hotkeys Overview
Action
Hotkey
Switch to Desktop
Alt + 1..9
Move Window to Desktop
Alt + Shift + 1..9
Move + Switch
Ctrl + Alt + 1..9
Tile Current Desktop
Alt + D
Gather All Windows
Alt + Shift + G
Toggle Pin (Always on Top)
Ctrl + Alt + B
Toggle Status Bar
Ctrl + Alt + B
Close Window Under Cursor
Alt + Q or Alt + Middle Click
Toggle Maximize
Alt + F
Hide Window
Alt + W
Open Terminal
Alt + Enter
Open With Vim
Alt + V
Reload Script
Alt + R
Pie Menu
Space + Right Click
Help Menu
Alt + /
Transparency Adjust
Alt + Mouse Wheel
And more (see code for full list).
## 🛠 Requirements
Turn off Windows animation

AutoHotkey v2

This script was created because I often needed to keep many windows open for work, but the machine was shared with other people. I couldn't install extra software or modify the system much, so I wrote a simple tool that improved my workflow without changing the environment.
Since it solved my own problem, I thought it might be helpful to share it with others as well.
