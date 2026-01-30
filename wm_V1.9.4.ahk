#Requires AutoHotkey v2.0
#SingleInstance Force
#WinActivateForce
; ==============================================================================
; Configuration
; ==============================================================================

SetWorkingDir(A_ScriptDir)
CoordMode("Mouse", "Screen")
SetTitleMatchMode(2)
SetWinDelay(0)
SetControlDelay(0)

global Color_Bg, Color_Text, Color_Active, BarHeight, BarTransparent, BarFontSize
global MenuSize, Radius, CenterZone, FontSize, FontSizeActive, MenuTransparent
global ButtonDir, OutputDir, OutputFile, VimPath, TerminalExe
global VimWinX, VimWinY, VimWinWidth, VimWinHeight
global OSDHeight, OSDTransparent
global WorkStart, WorkEnd, WDayBar, WorkTime
global PieConfig
global ConfigFile := A_ScriptDir . "\wm_config.ini"
global CurrentDesktop := 1
global DesktopCount   := 9
global Desktops       := Map()
global AlwaysVisible  := Map()
global BarVisible     := true
global CurrentVimPID  := 0
global LastClipContent := ""
global BarGui := "", BarLeftText := "", BarRightText := "", BarProgress := ""

PieConfig := Map(
    "Top", "↑", "TopRight", "↗", "Right", "→", "DownRight", "↘", 
    "Down", "↓", "DownLeft", "↙", "Left", "←", "TopLeft", "↖", "Center", "●"
)

; ==============================================================================
; Initialization
; ==============================================================================

LoadOrInitConfig()
Loop DesktopCount {
    Desktops[A_Index] := []
}

if !DirExist(OutputDir)
    DirCreate(OutputDir)
if !DirExist(ButtonDir)
    DirCreate(ButtonDir)

if InitializeButtons() {
    Reload()
}

CreateStatusBar()
UpdateStatusBar()
UpdateClockAndProgress()
SetTimer(UpdateClockAndProgress, 1000)
SetupTrayIcon()
if !DirExist(OutputDir)
DirCreate(OutputDir)
RecordClipboard()

; ==============================================================================
; Key Bindings
; ==============================================================================

Hotkey("!/", ShowHelpGui)
Loop 9 {
    i := A_Index
    Hotkey("!" . i, SwitchDesktop.Bind(i))
    Hotkey("!+" . i, MoveWindowToDesktop.Bind(i))
    Hotkey("^!" . i, MoveAndSwitch.Bind(i))
}

Hotkey("!d", TileCurrentDesktop)
Hotkey("!+g", GatherAllToCurrent)
Hotkey("^!t", TogglePin)
Hotkey("^!b", ToggleBar)
Hotkey("!F12", RestoreAndExit)

Hotkey("!q", CloseWindowUnderMouse)
Hotkey("!MButton", CloseWindowUnderMouse)
Hotkey("!f", ToggleMaximizeUnderMouse)
Hotkey("!t", ToggleTopUnderMouse)
Hotkey("!w", HideUnderMouse)

Hotkey("!WheelUp", AdjustTransparency.Bind(20))
Hotkey("!WheelDown", AdjustTransparency.Bind(-20))
~LButton & RButton::Send("^c")
~RButton & LButton::Send("^c")

Hotkey("!Enter", LaunchTerminal)
Hotkey("!s", (*) => Run("devmgmt.msc"))
Hotkey("!n", (*) => Run("ncpa.cpl"))
Hotkey("!v", OpenWithVim)
Hotkey("!x", ShowPowerMenu)

!r::{
    Reload
    }

^`:: ToggleVimWindow()
~^c:: (Sleep(100), RecordClipboard())

~Space & RButton:: PieMenu.Start()

Space Up:: 
RButton Up:: 
{
    if PieMenu.IsActive
        PieMenu.Execute()
}

; ==============================================================================
; Function Modules
; ==============================================================================
; ------------------------------------------------------------------------------
; [Module] Config File
; ------------------------------------------------------------------------------
LoadOrInitConfig() {
    global
    
    if !FileExist(ConfigFile) {
        DefaultIni := "
        (
        [WM_Config]

        [Visual]
        ; Background Color
        Color_Bg=181818
        ; Text Color
        Color_Text=CCCCCC
        ; Active/Highlight Color
        Color_Active=A020F0
        ; Status Bar Height
        BarHeight=35
        ; Status Bar Transparency (0-255)
        BarTransparent=200
        ; Status Bar FontSize
        BarFontSize=10
        
        [PieMenu]
        ; Menu Diameter
        MenuSize=300
        ; Center Dead Zone Range
        CenterZone=40
        ; Menu Font Size
        FontSize=14
        ; Active Item Font Size
        FontSizeActive=22
        ; Menu Transparency (0-255)
        MenuTransparent=200
        
        [Paths]
        ; Button Script Directory
        ButtonDir=Buttons
        ; Clipboard Recovery Output Directory
        OutputDir=C:\Users\Administrator\Documents
        ; Editor Path(Vim recommended)
        VimPath=C:\Windows\system32\notepad.exe
        ; Terminal Path
        TerminalExe=C:\Windows\system32\cmd.exe
        
        [Layout]
        ; Vim Window X
        VimWinX=400
        ; Vim Window Y
        VimWinY=0
        ; Vim Window Width
        VimWinWidth=1000
        ; Vim Window Height
        VimWinHeight=800

        [OSD]
        ; OSD Height Position
        OSDHeight=850
        ; OSD Transparency (0-255)
        OSDTransparent=200

        [WorkTime]
        ; Work Time Format HHmm
        WorkStart=0900
        WorkEnd=1745
        ; Set weekends full bar or normal
        WDayBar=off
        ; Set Work or full-day progress
        WorkTime=on
        )"
        
        try {
            FileAppend(DefaultIni, ConfigFile, "UTF-8")
            MsgBox("First run detected. Config file created at:`n" . ConfigFile . "`n`nPlease press 'Alt + /' to view the Help Menu.", "WM Config", "Iconi")
        } catch as e {
            MsgBox("Failed to create config file, check permissions!`n" . e.Message)
        }
    }

    ; --- Visual ---
    Color_Bg        := IniRead(ConfigFile, "Visual", "Color_Bg", "181818")
    Color_Text      := IniRead(ConfigFile, "Visual", "Color_Text", "CCCCCC")
    Color_Active    := IniRead(ConfigFile, "Visual", "Color_Active", "A020F0")
    BarHeight       := Integer(IniRead(ConfigFile, "Visual", "BarHeight", "28"))
    BarTransparent  := Integer(IniRead(ConfigFile, "Visual", "BarTransparent", "200"))
    BarFontSize     := Integer(IniRead(ConfigFile, "Visual", "BarFontSize", "10"))

    ; --- PieMenu ---
    MenuSize        := Integer(IniRead(ConfigFile, "PieMenu", "MenuSize", "300"))
    Radius          := MenuSize / 2 
    CenterZone      := Integer(IniRead(ConfigFile, "PieMenu", "CenterZone", "40"))
    FontSize        := Integer(IniRead(ConfigFile, "PieMenu", "FontSize", "14"))
    FontSizeActive  := Integer(IniRead(ConfigFile, "PieMenu", "FontSizeActive", "22"))
    MenuTransparent := Integer(IniRead(ConfigFile, "PieMenu", "MenuTransparent", "200"))

    ; --- Paths ---
    bDirTemp        := IniRead(ConfigFile, "Paths", "ButtonDir", "Buttons")
    ButtonDir       := (bDirTemp ~= "^[a-zA-Z]:") ? bDirTemp : (A_ScriptDir . "\" . bDirTemp)
    OutputDir       := IniRead(ConfigFile, "Paths", "OutputDir", "C:\Users\Administrator\Documents")
    OutputFile      := OutputDir . "\CB.txt"
    VimPath         := IniRead(ConfigFile, "Paths", "VimPath", "C:\Windows\system32\notepad.exe")
    TerminalExe     := IniRead(ConfigFile, "Paths", "TerminalExe", "C:\Windows\system32\cmd.exe")

    ; --- Layout ---
    VimWinX         := Integer(IniRead(ConfigFile, "Layout", "VimWinX", "400"))
    VimWinY         := Integer(IniRead(ConfigFile, "Layout", "VimWinY", "0"))
    VimWinWidth     := Integer(IniRead(ConfigFile, "Layout", "VimWinWidth", "1000"))
    VimWinHeight    := Integer(IniRead(ConfigFile, "Layout", "VimWinHeight", "800"))

    ; --- OSD ---
    OSDHeight       := Integer(IniRead(ConfigFile, "OSD", "OSDHeight", "850"))
    OSDTransparent  := Integer(IniRead(ConfigFile, "OSD", "OSDTransparent", "200"))

    ; --- WorkTime ---
    WorkStart       := IniRead(ConfigFile, "WorkTime", "WorkStart", "0900")
    WorkEnd         := IniRead(ConfigFile, "WorkTime", "WorkEnd", "1745")
    WDayBar         := IniRead(ConfigFile, "WorkTime", "WDayBar", "off")
    WorkTime        := IniRead(ConfigFile, "WorkTime", "WorkTime", "on")
}


; ------------------------------------------------------------------------------
; [Module] Help GUI
; ------------------------------------------------------------------------------
ShowHelpGui(*) {
    static helpGui := ""
    
    CloseWatcher() {
        if !IsObject(helpGui) {
            SetTimer CloseWatcher, 0
            return
        }

        if GetKeyState("Escape", "P") || GetKeyState("LButton", "P") {
            try helpGui.Destroy()
            helpGui := ""
            SetTimer CloseWatcher, 0
        }
    }

    if IsObject(helpGui) {
        helpGui.Destroy()
        helpGui := ""
        return
    }

    helpGui := Gui("-Caption +AlwaysOnTop +ToolWindow -DPIScale +Owner")
    helpGui.BackColor := Color_Bg
    helpGui.SetFont("s16 w700 c" . Color_Active, "Segoe UI")
    helpGui.Add("Text", "x0 y25 w600 Center", "HELP")
    
    helpGui.SetFont("s10 w600 c" . Color_Active)
    helpGui.Add("Text", "x50 y65 w500 h2 0x10")

    shortcuts := [
        ["Alt + /", "Show/Hide Help Menu"],
        ["Space + RClick", "Pie Menu (Mouse)"],
        ["Alt + 1-9", "Switch Desktop (Shift to Move)"],
        ["Alt + LButton", "Move Window"],
        ["Alt + RButton", "Resize Window"],
        ["Alt + Wheel", "Window Transparency"],
        ["Alt + Shift + G", "Gather All Windows"],
        ["Alt + Q", "Close Window"],
        ["Alt + D", "Smart Tiling"],
        ["Alt + W", "Minimize Window"],
        ["Alt + F", "Maximize/Restore"],
        ["Alt + R", "Reload Script"],
        ["Alt + T", "Toggle Pin/OnTop"],
        ["Ctrl + ``", "Clipboard History"],
        ["Ctrl + Alt + B", "Toggle Top Bar"],
        ["Alt + V", "Edit Selectd File"],
        ["Alt + F12", "Safely Exit"],
        ["Alt + X", "Power Menu"]
    ]

    helpGui.SetFont("s11 w400 c" . Color_Text)
    for i, item in shortcuts {
        yPos := 80 + (i-1)*30
        helpGui.Add("Text", "x60 y" . yPos . " w160 c" . Color_Active, item[1])
        helpGui.Add("Text", "x220 y" . yPos . " w320", item[2])
    }
    
    helpGui.Show("Center")
    SetTimer CloseWatcher, 50
}


; ------------------------------------------------------------------------------
; [Module] Round Menu
; ------------------------------------------------------------------------------
InitializeButtons() {
    dirs := ["Top", "TopRight", "Right", "DownRight", "Down", "DownLeft", "Left", "TopLeft"]
    created := false
    
    for d in dirs {
        fPath := ButtonDir . "\" . d . ".ahk"
        if !FileExist(fPath) {
            template := '
            (
            SetWorkingDir(A_ScriptDir)
            CoordMode("Mouse", "Screen")
            SetTitleMatchMode(2)
            SetWinDelay(0)
            SetControlDelay(0)

            %dir%(){
                ToolTip "%dir%"
                Sleep 200
                ToolTip()
            }
            )'
            FileAppend(StrReplace(template, "%dir%", d), fPath, "UTF-8")
            created := true
        }
    }
    return created 
}

; ------------------------------------------------------------------------------
; [Module] OSD
; ------------------------------------------------------------------------------
ShowOSD(text) {
    try{
    global OSDHeight, OSDTransparent
    global Color_Active
    static OsdGui := ""
    if IsObject(OsdGui)
        OsdGui.Destroy()
    OsdGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Disabled +Owner -DPIScale")
    OsdGui.BackColor := Color_Bg
    OsdGui.SetFont("s20 w600 c" . Color_Active, "Segoe UI")
    OsdGui.Add("Text", "Center", text)
    OsdGui.Show(Format("NoActivate AutoSize y{}", OSDHeight))
    WinSetTransparent(OSDTransparent, OsdGui.Hwnd)
    SetTimer(() => (IsObject(OsdGui) ? OsdGui.Destroy() : ""), -1000)
    }
}
; ------------------------------------------------------------------------------
; [Module] Windows Function
; ------------------------------------------------------------------------------
CloseWindowUnderMouse(*) {
    MouseGetPos(,, &hwnd)
    try {
        WinClose(hwnd)
        ShowOSD("Closing Windows...")
    }
}

HideUnderMouse(*) {
    MouseGetPos(,, &hwnd)
    try {
        WinMinimize(hwnd)
        ShowOSD("WinMinimized")
    }
}

ToggleMaximizeUnderMouse(*) {
    MouseGetPos(,, &hwnd)
    try {
        if WinGetMinMax(hwnd) {
            WinRestore(hwnd)
            ShowOSD("WinRestore")
        } else {
            WinMaximize(hwnd)
            ShowOSD("WinMaximized")
        }
    }
}

ToggleTopUnderMouse(*) {
    MouseGetPos(,, &hwnd)
    try {
        WinSetAlwaysOnTop(-1, hwnd)
        state := WinGetExStyle(hwnd)
        ShowOSD((state & 0x8) ? "WinOnTop" : "WinOffTop")
    }
}

AdjustTransparency(amount, *) {
    MouseGetPos(,, &hwnd)
    try {
        cur := WinGetTransparent(hwnd)
        if !IsNumber(cur) 
            cur := 255
        newVal := cur + amount
        newVal := Max(2.55, Min(255, newVal))
        WinSetTransparent(newVal, hwnd)
        ShowOSD("WinTransparent: " . Integer(newVal/2.55) . "%")
    }
}

; ------------------------------------------------------------------------------
; [Module] Pie Menu
; ------------------------------------------------------------------------------
class PieMenu {
    
    static IsActive := false, GuiObj := "", Labels := Map(), TimerFn := ObjBindMethod(PieMenu, "CheckMouse")
    static StartX := 0, StartY := 0, CurrentSector := "", LastSector := ""

    static Start() {
        if this.IsActive || GetKeyState("Alt", "P")
            return
        this.IsActive := true
        MouseGetPos(&x, &y)
        this.StartX := x, this.StartY := y, this.CurrentSector := "Center", this.LastSector := ""
        this.CreateGui()
        SetTimer(this.TimerFn, 10)
    }

    static CreateGui() {
        global MenuTransparent
        this.GuiObj := Gui("-Caption +AlwaysOnTop +ToolWindow +Owner +E0x20 -DPIScale") 
        this.GuiObj.BackColor := Color_Bg
        WinSetTransparent(MenuTransparent, this.GuiObj)
        WinSetRegion("0-0 w" . MenuSize . " h" . MenuSize . " E", this.GuiObj)
        
        this.Labels["Center"] := this.GuiObj.Add("Text", "x" Radius-20 " y" Radius-20 " w40 h40 Center +0x200 c" Color_Text, PieConfig["Center"])
        
        dirs := ["Right", "DownRight", "Down", "DownLeft", "Left", "TopLeft", "Top", "TopRight"]
        loop 8 {
            dir := dirs[A_Index], angle := (A_Index-1)*45, rad := angle*0.01745329
            fX := Radius + Cos(rad)*Radius*0.75 - 30
            fY := Radius + Sin(rad)*Radius*0.75 - 20
            this.Labels[dir] := this.GuiObj.Add("Text", "x" fX " y" fY " w60 h40 Center +0x200 c" Color_Text, PieConfig[dir])
        }
        this.GuiObj.Show("x" this.StartX-Radius " y" this.StartY-Radius " w" MenuSize " h" MenuSize " NoActivate")
    }

    static CheckMouse() {
        if !this.IsActive
            return
        MouseGetPos(&mx, &my)
        dx := mx - this.StartX
        dy := my - this.StartY
        dist := Sqrt(dx*dx + dy*dy)

        if (dist < CenterZone) {
            this.CurrentSector := "Center"
        } else {
            angle := DllCall("msvcrt\atan2", "Double", dy, "Double", dx, "Cdecl Double") * 180 / 3.1415926
            if (angle < 0)
                angle += 360
            sectorIdx := Round(angle / 45)
            if (sectorIdx == 8)
                sectorIdx := 0
            static dirMap := ["Right", "DownRight", "Down", "DownLeft", "Left", "TopLeft", "Top", "TopRight"]
            this.CurrentSector := dirMap[sectorIdx + 1]
        }

        if (this.CurrentSector != this.LastSector) {
            this.UpdateUI()
            this.LastSector := this.CurrentSector
        }
    }

    static UpdateUI() {
        if !IsObject(this.GuiObj)
            return
        for dir, ctrl in this.Labels {
            try {
                ctrl.SetFont("s" . FontSize . " c" . Color_Text . " w600")
                ctrl.Opt("c" . Color_Text)
            }
        }
        if this.Labels.Has(this.CurrentSector) {
            try {
                curr := this.Labels[this.CurrentSector]
                curr.SetFont("s" . FontSizeActive . " c" . Color_Active . " w700")
                curr.Opt("c" . Color_Active)
            }
        }
    }

    static Execute() {
        this.IsActive := false
        SetTimer(this.TimerFn, 0)
        if IsObject(this.GuiObj)
            this.GuiObj.Destroy()
        if (this.CurrentSector != "Center" && this.CurrentSector != "") {
            try %this.CurrentSector%()
            catch
                ShowOSD("Function Loust: " . this.CurrentSector)
        }
    }
}

; ------------------------------------------------------------------------------
; [Module] WM & Bar
; ------------------------------------------------------------------------------
SwitchDesktop(target, *) {
    global CurrentDesktop, Desktops, AlwaysVisible
    if (target == CurrentDesktop) {
        ShowOSD("Desktop " . target)
        return
    }
    Desktops[CurrentDesktop] := GetVisibleWindows()
    for hwnd in Desktops[CurrentDesktop]
        if (!AlwaysVisible.Has(hwnd))
            try WinMinimize(hwnd)
    for hwnd in Desktops[target]
        try WinRestore(hwnd)
    for hwnd, _ in AlwaysVisible
        try WinRestore(hwnd)
    
    CurrentDesktop := target
    UpdateStatusBar()
    ShowOSD("DeskTop " . CurrentDesktop)
}

MoveWindowToDesktop(target, *) {
    global CurrentDesktop, Desktops, AlwaysVisible
    
    hwnd := 0
    try {
        hwnd := WinExist("A")
    } catch {
        return
    }
    
    if (!hwnd || (BarGui && hwnd == BarGui.Hwnd)) 
        return

    if (AlwaysVisible.Has(hwnd)) 
        AlwaysVisible.Delete(hwnd)

    Loop DesktopCount {
        d := A_Index
        if (Desktops.Has(d)) {
            nl := []
            for h in Desktops[d] {
                if (h != hwnd) 
                    nl.Push(h)
            }
            Desktops[d] := nl
        }
    }
    
    Desktops[target].Push(hwnd)
    
    if (target != CurrentDesktop) {
        try WinMinimize(hwnd)
        ShowOSD("Window -> Desktop " . target)
    }
}

MoveAndSwitch(target, *) {
    MoveWindowToDesktop(target)
    SwitchDesktop(target)
    ShowOSD("Move And Switch -> " . target)
}

CreateStatusBar() {
    global BarTransparent, BarGui, BarLeftText, BarRightText, BarProgress, BarHeight, Color_Active, BarFontSize

    try {
        if IsSet(BarGui) && IsObject(BarGui)
            BarGui.Destroy()
    }

    local minNeededHeight := Round(BarFontSize * 2.5)
    if !IsSet(BarHeight) || (BarHeight < minNeededHeight)
        BarHeight := minNeededHeight
    local padding := 15
    local progressWidth := Round(A_ScreenWidth * 0.30)
    local textBoxWidth := Round((A_ScreenWidth - progressWidth - (padding * 4)) / 2)
    local textControlH := Round(BarFontSize * 2)
    local textY := (BarHeight - textControlH) / 2
    local progH := 6
    local progY := (BarHeight - progH) / 2
    local progressX := (A_ScreenWidth / 2) - (progressWidth / 2)
    local rightTextX := A_ScreenWidth - textBoxWidth - padding

    BarGui := Gui("-Caption +AlwaysOnTop +ToolWindow +Owner +E0x08000000 -DPIScale")
    BarGui.BackColor := "181818"
    BarGui.SetFont("s" . BarFontSize . " w600 c" . Color_Active, "Segoe UI")
    BarLeftText := BarGui.Add("Text", "x" . padding . " y" . textY . " w" . textBoxWidth . " h" . textControlH . " BackgroundTrans", "")
    BarGui.Add("Text", "x" . progressX . " y" . progY . " w" . progressWidth . " h" . progH . " Background333333", "") 
    ProgressOptions := Format("x{1} y{2} w{3} h{4} c{5} Background333333 +Smooth", progressX, progY, progressWidth, progH, Color_Active)
    BarProgress := BarGui.Add("Progress", ProgressOptions, 0)
    BarRightText := BarGui.Add("Text", "x" . rightTextX . " y" . textY . " w" . textBoxWidth . " h" . textControlH . " BackgroundTrans Right", "")
    BarGui.Show("x0 y0 w" . A_ScreenWidth . " h" . BarHeight . " NoActivate")

    WinSetTransparent(BarTransparent, BarGui.Hwnd)
}UpdateStatusBar() {
    global CurrentDesktop, DesktopCount, BarLeftText
    if !IsObject(BarLeftText)
        return
    str := ""
    Loop DesktopCount
        str .= (A_Index == CurrentDesktop) ? " [" A_Index "] " : "  " A_Index "  "
    try BarLeftText.Value := str
}

UpdateClockAndProgress() {
    global BarRightText, BarProgress
    global WorkStart, WorkEnd, WDayBar, WorkTime
    
    if !IsObject(BarRightText)
        return

    try BarRightText.Value := FormatTime(, "yyyy-MM-dd   HH:mm")

    NowTime := A_Now
    TodayDate := FormatTime(NowTime, "yyyyMMdd")
    WDay := A_WDay
    StartTS := ""
    EndTS := ""
    ForceFull := false
    if (WorkTime = "off") {
        StartTS := TodayDate . "000000"
        EndTS   := TodayDate . "235959"
    } 
    else {
        StartTS := TodayDate . WorkStart . "00"
        EndTS   := TodayDate . WorkEnd . "00"

        if (WDay == 1 || WDay == 7) {
            if (WDayBar = "off") {
                ForceFull := true
            }
        }
    }
    
    pct := 0
    
    if (ForceFull) {
        pct := 100
    } else {
        TotalSec := DateDiff(EndTS, StartTS, "Seconds")
        ElapsedSec := DateDiff(NowTime, StartTS, "Seconds")
        
        if (TotalSec <= 0) {
            pct := 100 
        } else if (ElapsedSec < 0) {
            pct := 0
        } else if (ElapsedSec > TotalSec) {
            pct := 100
        } else {
            pct := (ElapsedSec / TotalSec) * 100
        }
    }
    
    try {
        if IsObject(BarProgress)
            BarProgress.Value := Integer(pct)
    }
}

ToggleBar(*) {
    global BarVisible, BarGui
    if (BarVisible := !BarVisible)
        BarGui.Show("NoActivate")
    else
        BarGui.Hide()
}


TogglePin(*) {
    global AlwaysVisible
    hwnd := 0
    try { 
        hwnd := WinExist("A") 
    } catch { 
        return 
    }
    
    if (!hwnd || (BarGui && hwnd == BarGui.Hwnd)) 
        return
        
    if (AlwaysVisible.Has(hwnd)) {
        AlwaysVisible.Delete(hwnd)
        ShowOSD("Unpinned")
    } else {
        AlwaysVisible[hwnd] := true
        ShowOSD("Pinned (Always Visible)")
    }
}

GatherAllToCurrent(*) {
    global Desktops, CurrentDesktop, AlwaysVisible
    ShowOSD("Gathering All Windows...")
    
    fullList := WinGetList()
    Loop DesktopCount
        Desktops[A_Index] := []
    
    AlwaysVisible.Clear()
    count := 0
    
    for hwnd in fullList {
        try {
            if (BarGui && hwnd == BarGui.Hwnd) 
                continue
            
            class := WinGetClass(hwnd)
            if (class == "Progman" || class == "Shell_TrayWnd") 
                continue
            
            WinRestore(hwnd)
            Desktops[CurrentDesktop].Push(hwnd)
            count++
        }
    }
    ShowOSD("Gathered " . count . " Windows")
}

; Tile Rule
TileCurrentDesktop(*) {
    global BarHeight, BarVisible
    
    windows := GetVisibleWindow()
    count := windows.Length
    
    if (count == 0) {
        ShowOSD("No Windows To Tile")
        return
    }
    ShowOSD("Tile: " . count)
    
    activeWin := WinExist("A")
    targetMon := GetMonitorIndex(activeWin)
    MonitorGetWorkArea(targetMon, &WL, &WT, &WR, &WB)
    if (IsSet(BarVisible) && BarVisible && IsSet(BarHeight))
        WT += BarHeight
    
    W := WR - WL
    H := WB - WT 
    
    if (count == 1) {
        try WinRestore(windows[1]), WinMove(WL, WT, W, H, windows[1])
    } 
    else if (count == 2) {
        try WinRestore(windows[1]), WinMove(WL, WT, W/2, H, windows[1])
        try WinRestore(windows[2]), WinMove(WL + W/2, WT, W/2, H, windows[2])
    }
    else if (count == 3) {
        try WinRestore(windows[1]), WinMove(WL, WT, W/2, H, windows[1])
        try WinRestore(windows[2]), WinMove(WL + W/2, WT, W/2, H/2, windows[2])
        try WinRestore(windows[3]), WinMove(WL + W/2, WT + H/2, W/2, H/2, windows[3])
    }
    else if (count == 5) {
        colW := W / 3, halfH := H / 2
        try WinRestore(windows[1]), WinMove(WL + colW, WT, colW, H, windows[1])
        try WinRestore(windows[2]), WinMove(WL, WT, colW, halfH, windows[2])
        try WinRestore(windows[3]), WinMove(WL, WT + halfH, colW, halfH, windows[3])
        try WinRestore(windows[4]), WinMove(WL + 2*colW, WT, colW, halfH, windows[4])
        try WinRestore(windows[5]), WinMove(WL + 2*colW, WT + halfH, colW, halfH, windows[5])
    }
    else if (Mod(count, 2) != 0) {
        try {
            itemWidth := W / count
            for i, hwnd in windows
                WinRestore(hwnd), WinMove(WL + (i-1)*itemWidth, WT, itemWidth, H, hwnd)
        }
    } else {
        try {
            cols := count / 2, itemW := W / cols, itemH := H / 2
            for i, hwnd in windows {
                idx := i - 1, r := Floor(idx/cols), c := Mod(idx, cols)
                WinRestore(hwnd), WinMove(WL + c*itemW, WT + r*itemH, itemW, itemH, hwnd)
            }
        }
    }
}

GetVisibleWindow() {
    windows := []
    ids := WinGetList(,, "Program Manager")
    for this_id in ids {
        try {
            style := WinGetStyle(this_id)
        } catch {
            continue
        }
        if !(style & 0x10000000) 
            continue
            
        exStyle := WinGetExStyle(this_id)
        if (exStyle & 0x00000080)
            continue
            
        isCloaked := 0
        try {
            DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", this_id, "Int", 14, "Int*", &isCloaked, "Int", 4)
            if (isCloaked)
                continue
        }
        
        title := WinGetTitle(this_id)
        if (title == "")
            continue
            
        WinGetPos(,, &w, &h, this_id)
        if (w < 100 || h < 100)
            continue
            
        windows.Push(this_id)
    }
    return windows
}

GetMonitorIndex(hwnd := 0) {
    if !hwnd || !WinExist(hwnd) {
        MouseGetPos(&mx, &my)
        loop MonitorGetCount() {
            MonitorGet(A_Index, &mL, &mT, &mR, &mB)
            if (mx >= mL && mx < mR && my >= mT && my < mB)
                return A_Index
        }
        return 1
    }
    
    WinGetPos(&wx, &wy, &ww, &wh, hwnd)
    cx := wx + (ww / 2), cy := wy + (wh / 2)
    loop MonitorGetCount() {
        MonitorGet(A_Index, &mL, &mT, &mR, &mB)
        if (cx >= mL && cx < mR && cy >= mT && cy < mB)
            return A_Index
    }
    return 1 
}

; ------------------------------------------------------------------------------
; [Module] Clipboard Tool and Others
; ------------------------------------------------------------------------------
RecordClipboard() {
    global LastClipContent
    try txt := A_Clipboard
    catch 
        return
    if (Type(txt) != "String" || txt == "" || txt == LastClipContent)
        return
    LastClipContent := txt
    Content := "------------------------------------------------------------------------------------------------`r`n"
             . FormatTime(, "yyyy-MM-dd HH:mm:ss") . "`r`n" . txt . "`r`n`r`n"
    try FileAppend(Content, OutputFile, "UTF-8")
}

ToggleVimWindow() {
    global CurrentVimPID, VimPath, OutputFile, VimWinX, VimWinY
    if (CurrentVimPID && WinExist("ahk_pid " . CurrentVimPID)) {
        WinClose("ahk_pid " . CurrentVimPID)
        CurrentVimPID := 0
    } else {
        if InStr(VimPath, "vim") {
            RunCmd := Format('"{1}" "+$" "{2}"', VimPath, OutputFile)
        } else {
            RunCmd := Format('"{1}" "{2}"', VimPath, OutputFile)
        }

        try {
            Run(RunCmd, , , &pid)
            CurrentVimPID := pid
            if WinWait("ahk_pid " . pid, , 3) {
                WinSetAlwaysOnTop(1, "ahk_pid " . pid)
                WinMove(VimWinX, VimWinY, , , "ahk_pid " . pid)
                WinActivate("ahk_pid " . pid)
            }
        } catch {
            ShowOSD("Vim Boot Filed")
        }
    }
}

LaunchTerminal(*) {
    path := Explorer_GetPath()
    try Run('"' . TerminalExe . '"' . (path ? ' -d "' . path . '"' : ""))
}

OpenWithVim(*) {
    targetPath := Explorer_GetSelection()
    if (targetPath == "") {
        ShowOSD("No file selected")
        return
    }
    try {
        Run('"' . VimPath . '" "' . targetPath . '"')
    } catch as e {
        ShowOSD("Vim Launch Failed")
    }
}

ShowPowerMenu(*) {
    static pGui := ""
    if IsObject(pGui) {
        pGui.Destroy(), pGui := ""
        return
    }
    pGui := Gui("+AlwaysOnTop -Caption +ToolWindow +Owner")
    pGui.BackColor := "2e3440"
    pGui.SetFont("s12", "Arial")
    pGui.Add("Text", "x0 y15 w500 Center c" . Color_Active, "System Power Menu")
    pGui.Add("Text", "x50 y45 w400 h2 0x10")
    
    AddBtn(x, y, txt, fn, col) {
        btn := pGui.Add("Text", "x" x " y" y " w120 h60 Center 0x200 +Border cWhite Background" col, txt)
        btn.OnEvent("Click", fn)
    }
    AddBtn(50, 70, "Shutdown", (*) => Shutdown(1), "b48ead")
    AddBtn(190, 70, "Sleep", (*) => DllCall("PowrProf\SetSuspendState", "Int", 0, "Int", 0, "Int", 0), "5e81ac")
    AddBtn(330, 70, "Reboot", (*) => Shutdown(2), "bf616a")
    pGui.OnEvent("Escape", (*) => (pGui.Destroy(), pGui := ""))
    pGui.Show("w500 h160")
}

; ------------------------------------------------------------------------------
; [Module] Others 
; ------------------------------------------------------------------------------
RestoreAndExit(*) {
    global BarGui
    ShowOSD("Script Shutting Down ...")
    Sleep(500)
    if IsObject(BarGui)
        BarGui.Destroy()
    for hwnd in WinGetList() {
        try {
            class := WinGetClass(hwnd)
            if (class != "Progman" && class != "Shell_TrayWnd")
                WinRestore(hwnd)
        }
    }
    ExitApp
}

GetVisibleWindows() {
    global BarGui
    list := WinGetList()
    windows := []
    for hwnd in list {
        try {
            if (BarGui && hwnd == BarGui.Hwnd) 
                continue
            class := WinGetClass(hwnd)
            if (class == "Progman" || class == "Shell_TrayWnd") 
                continue
            if (WinGetMinMax(hwnd) != -1) 
                windows.Push(hwnd)
        }
    }
    return windows
}

Explorer_GetSelection() {
    hwnd := WinExist("A")
    if !hwnd
        return ""

    WinClass := WinGetClass(hwnd)
    
    if (WinClass ~= "Progman|WorkerW") {
        try {
            oDesktop := ComObject("Shell.Application").Windows.Item(ComValue(19, 8))
            
            sel := oDesktop.Document.SelectedItems
            if (sel.Count > 0)
                return sel.Item(0).Path
        } catch {
            return ""
        }
    }
    else if (WinClass ~= "(Cabinet|Explore)WClass") {
        try {
            for window in ComObject("Shell.Application").Windows {
                if (window.HWND == hwnd) {
                    sel := window.Document.SelectedItems
                    if (sel.Count > 0)
                        return sel.Item(0).Path
                }
            }
        }
    }
    
    return ""
}

Explorer_GetPath() {
    hwnd := WinExist("A")
    if !hwnd
        return ""
    
    WinClass := WinGetClass(hwnd)
    
    if (WinClass ~= "Progman|WorkerW")
        return A_Desktop
        
    if (WinClass ~= "(Cabinet|Explore)WClass") {
        try {
            for window in ComObject("Shell.Application").Windows {
                if (window.HWND == hwnd)
                    return window.Document.Folder.Self.Path
            }
        }
    }
    return ""
}

SetupTrayIcon() {
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Restore & Exit", RestoreAndExit)
}

; ------------------------------------------------------------------------------
; [Module] Mouse Interaction
; ------------------------------------------------------------------------------

!LButton:: {
    MouseGetPos(,, &hwnd)
    
    try {
        WinActivate(hwnd)
    } catch {
        return
    }

    if (WinGetMinMax(hwnd) == 1) {
        try {
            WinRestore(hwnd)
            WinGetPos(,, &rw, &rh, hwnd)
            MouseGetPos(&mx, &my)
            WinMove(mx - rw/2, my - rh/2,,, hwnd)
        } catch {
            return
        }
    }
    
    MouseGetPos(&startX, &startY)
    try {
        WinGetPos(&winX, &winY,,, hwnd)
    } catch {
        return
    }
    
    while GetKeyState("LButton", "P") {
        MouseGetPos(&curX, &curY)
        try WinMove(winX + (curX - startX), winY + (curY - startY),,, hwnd)
    }
}

!RButton:: {
    MouseGetPos(,, &hwnd)
    if (WinGetMinMax(hwnd) == 1) 
        return
        
    try {
        WinGetPos(&winX, &winY, &winW, &winH, hwnd)
        MouseGetPos(&startX, &startY)
        isLeft := (startX - winX) / winW < 0.5
        isUp   := (startY - winY) / winH < 0.5
        
        while GetKeyState("RButton", "P") {
            MouseGetPos(&curX, &curY)
            dX := curX - startX, dY := curY - startY
            nX := isLeft ? (winX+dX) : winX, nW := isLeft ? (winW-dX) : (winW+dX)
            nY := isUp ? (winY+dY) : winY, nH := isUp ? (winH-dY) : (winH+dY)
            
            if (nW > 50 && nH > 50)
                try WinMove(nX, nY, nW, nH, hwnd)
        }
    }
}
; ------------------------------------------------------------------------------
; [Module] External Buttons Include
; ------------------------------------------------------------------------------

#Include "*i %A_ScriptDir%\Buttons\Top.ahk"
#Include "*i %A_ScriptDir%\Buttons\TopRight.ahk"
#Include "*i %A_ScriptDir%\Buttons\Right.ahk"
#Include "*i %A_ScriptDir%\Buttons\DownRight.ahk"
#Include "*i %A_ScriptDir%\Buttons\Down.ahk"
#Include "*i %A_ScriptDir%\Buttons\DownLeft.ahk"
#Include "*i %A_ScriptDir%\Buttons\Left.ahk"
#Include "*i %A_ScriptDir%\Buttons\TopLeft.ahk"
