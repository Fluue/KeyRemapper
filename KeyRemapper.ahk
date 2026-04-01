#Requires AutoHotkey v2.0
#SingleInstance Force

global SaveDir := A_MyDocuments "\RemapperV2.0"
global SelectedSource := "", SelectedDest := ""
global ActiveRemaps := []
global HeldKeys := Map()
global IsCapturing := false
global CurrentCaptureTarget := ""

if !DirExist(SaveDir)
    DirCreate(SaveDir)

; Enables dark title bar on Win10 1809+
SetDarkMode(guiObj) {
    if (VerCompare(A_OSVersion, "10.0.17763") >= 0)
        try DllCall("dwmapi\DwmSetWindowAttribute", "ptr", guiObj.Hwnd, "int", 20, "int*", 1, "int", 4)
    guiObj.BackColor := "0x121212"
}

MyGui := Gui("+AlwaysOnTop", "Pro Key Remapper")
SetDarkMode(MyGui)
MyGui.SetFont("s10 cWhite", "Segoe UI")

MyGui.Add("Text", "Center w280 c808080", "1. Click box -> 2. Press Key Combo")

MyGui.Add("Text", "y+15 c00FFFF", "WHEN I PRESS (Source):")
SourceBtn := MyGui.Add("Button", "vSourceBtn w280 h40", "[ Click to set ]")
SourceBtn.OnEvent("Click", StartCapture)

MyGui.Add("Text", "y+10 c00FF00", "ACTION TO PERFORM (Destination):")
DestBtn := MyGui.Add("Button", "vDestBtn w280 h40", "[ Click to set ]")
DestBtn.OnEvent("Click", StartCapture)

MyGui.Add("Text", "y+10 c808080", "TRIGGER TYPE:")
TriggerMode := MyGui.Add("DropDownList", "vTriggerMode w280 Choose1 Background222222 cWhite",
    ["Full Remap", "On Press", "On Release"])

AddBtn := MyGui.Add("Button", "y+20 w280 h35 Default", "Add Mapping")
AddBtn.OnEvent("Click", AddRemap)

RemapList := MyGui.Add("ListBox", "r5 w280 vActiveList Background222222 cWhite")

RemoveBtn := MyGui.Add("Button", "w135 x10", "Remove Selected")
RemoveBtn.OnEvent("Click", RemoveMapping)

ResetBtn := MyGui.Add("Button", "w135 x155 yp", "Reset All")
ResetBtn.OnEvent("Click", (*) => Reload())

MyGui.Add("Text", "x10 y+15 w280 c808080", "PROFILES:")
global ProfileCombo := MyGui.Add("ComboBox", "x10 y+5 w145 Section Background222222 cWhite", GetProfileList())

SaveBtn := MyGui.Add("Button", "ys w60 h26", "Save")
SaveBtn.OnEvent("Click", SaveProfile)
LoadBtn := MyGui.Add("Button", "ys w60 h26", "Load")
LoadBtn.OnEvent("Click", LoadProfile)

if (ProfileCombo.Text == "")
    ProfileCombo.Text := "Default"

MyGui.Show()


StartCapture(GuiCtrl, *) {
    global IsCapturing, CurrentCaptureTarget
    if IsCapturing
        return

    CurrentCaptureTarget := GuiCtrl
    GuiCtrl.Text := ">>> RECORDING... <<<"
    GuiCtrl.SetFont("cFF0000")
    IsCapturing := true

    ih := InputHook("L1 M")
    ih.KeyOpt("{All}", "E")
    ih.OnEnd := (ihObj) => FinishCapture(ihObj.EndKey)
    ih.Start()
}

FinishCapture(KeyName) {
    global IsCapturing, SelectedSource, SelectedDest, CurrentCaptureTarget
    if (!IsCapturing || CurrentCaptureTarget == "")
        return

    mods := ""
    if GetKeyState("Ctrl",  "P") && !InStr(KeyName, "Ctrl")    && !InStr(KeyName, "Control")
        mods .= "^"
    if GetKeyState("Alt",   "P") && !InStr(KeyName, "Alt")
        mods .= "!"
    if GetKeyState("Shift", "P") && !InStr(KeyName, "Shift")
        mods .= "+"
    if (GetKeyState("LWin", "P") || GetKeyState("RWin", "P")) && !InStr(KeyName, "Win")
        mods .= "#"

    finalKey := mods . KeyName
    if (finalKey != "") {
        if (CurrentCaptureTarget.Name == "SourceBtn") {
            SelectedSource := finalKey
            SourceBtn.Text := finalKey
        } else {
            SelectedDest := finalKey
            DestBtn.Text := finalKey
        }
    }

    CurrentCaptureTarget.SetFont("cFFFFFF")
    IsCapturing := false
    CurrentCaptureTarget := ""
}

#HotIf IsCapturing
LButton::FinishCapture("LButton")
RButton::FinishCapture("RButton")
MButton::FinishCapture("MButton")
XButton1::FinishCapture("XButton1")
XButton2::FinishCapture("XButton2")
#HotIf


AddRemap(*) {
    global SelectedSource, SelectedDest
    if (SelectedSource == "" || SelectedDest == "") {
        MsgBox("Please set both keys first.", "Missing Keys")
        return
    }
    CreateBind(SelectedSource, SelectedDest, TriggerMode.Text)
    SourceBtn.Text := "[ Click to set ]"
    DestBtn.Text  := "[ Click to set ]"
    SelectedSource := SelectedDest := ""
}

CreateBind(src, dest, mode) {
    global ActiveRemaps
    try {
        if (mode == "Full Remap") {
            Hotkey("*" src,        DoDown.Bind(src, dest), "On")
            Hotkey("*" src " Up",  DoUp.Bind(src, dest),   "On")
        } else if (mode == "On Press") {
            Hotkey("*" src,       DoSinglePress.Bind(src, dest), "On")
        } else if (mode == "On Release") {
            Hotkey("*" src " Up", DoSinglePress.Bind(src, dest), "On")
        }
        ActiveRemaps.Push({Src: src, Dest: dest, Mode: mode})
        RemapList.Add([src " -> " dest " [" mode "]"])
    } catch Error as e {
        MsgBox("Couldn't create bind for '" src "': " e.Message)
    }
}

RemoveMapping(*) {
    global ActiveRemaps, HeldKeys
    idx := RemapList.Value
    if (idx == 0)
        return

    item := ActiveRemaps[idx]
    try {
        if (item.Mode == "Full Remap") {
            Hotkey("*" item.Src,       "Off")
            Hotkey("*" item.Src " Up", "Off")
        } else if (item.Mode == "On Press") {
            Hotkey("*" item.Src, "Off")
        } else if (item.Mode == "On Release") {
            Hotkey("*" item.Src " Up", "Off")
        }
        if HeldKeys.Has(item.Src)
            HeldKeys.Delete(item.Src)
    }
    ActiveRemaps.RemoveAt(idx)
    RemapList.Delete(idx)
}

ClearAllRemaps() {
    global ActiveRemaps, HeldKeys
    for item in ActiveRemaps {
        try {
            if (item.Mode == "Full Remap") {
                Hotkey("*" item.Src,       "Off")
                Hotkey("*" item.Src " Up", "Off")
            } else if (item.Mode == "On Press") {
                Hotkey("*" item.Src, "Off")
            } else if (item.Mode == "On Release") {
                Hotkey("*" item.Src " Up", "Off")
            }
        }
    }
    ActiveRemaps := []
    HeldKeys := Map()
    RemapList.Delete()
}


GetProfileList() {
    list := []
    Loop Files, SaveDir "\*.rmp"
        list.Push(RegExReplace(A_LoopFileName, "\.rmp$"))
    return list
}

SaveProfile(*) {
    profName := RegExReplace(ProfileCombo.Text, '[\\/:\*\?"<>\|]', "")
    if (profName == "") {
        MsgBox("Please enter a profile name.", "Error")
        return
    }
    filePath := SaveDir "\" profName ".rmp"
    try {
        if FileExist(filePath)
            FileDelete(filePath)
        f := FileOpen(filePath, "w")
        for item in ActiveRemaps
            f.WriteLine(item.Src "|" item.Dest "|" item.Mode)
        f.Close()
        ProfileCombo.Delete()
        ProfileCombo.Add(GetProfileList())
        ProfileCombo.Text := profName
        MsgBox("Profile '" profName "' saved!", "Saved", "T1.5")
    } catch Error as e {
        MsgBox("Failed to save: " e.Message)
    }
}

LoadProfile(*) {
    profName := ProfileCombo.Text
    filePath := SaveDir "\" profName ".rmp"
    if !FileExist(filePath) {
        MsgBox("Couldn't find '" profName "' in Documents\RemapperV2.0", "Error")
        return
    }
    ClearAllRemaps()
    try {
        Loop Read filePath {
            if (A_LoopReadLine == "")
                continue
            parts := StrSplit(A_LoopReadLine, "|")
            if (parts.Length == 3)
                CreateBind(parts[1], parts[2], parts[3])
        }
        MsgBox("Loaded '" profName "'!", "Loaded", "T1.5")
    } catch Error as e {
        MsgBox("Error loading profile: " e.Message)
    }
}


DoSinglePress(srcKey, destKey, *) {
    RegExMatch(destKey, "^([\^\!\+\#]*)(.*)", &m)
    Send("{Blind}" m[1] "{" m[2] "}")
}

DoDown(srcKey, destKey, *) {
    global HeldKeys
    if (HeldKeys.Has(srcKey) && HeldKeys[srcKey])
        return
    HeldKeys[srcKey] := true
    RegExMatch(destKey, "^([\^\!\+\#]*)(.*)", &m)
    if (m[1] != "")
        SendModifiersDown(m[1])
    Send("{Blind}{" m[2] " Down}")
}

DoUp(srcKey, destKey, *) {
    global HeldKeys
    HeldKeys[srcKey] := false
    RegExMatch(destKey, "^([\^\!\+\#]*)(.*)", &m)
    Send("{Blind}{" m[2] " Up}")
    if (m[1] != "")
        SendModifiersUp(m[1])
}

SendModifiersDown(mods) {
    if InStr(mods, "^")
        Send("{Blind}{LCtrl Down}")
    if InStr(mods, "!")
        Send("{Blind}{LAlt Down}")
    if InStr(mods, "+")
        Send("{Blind}{LShift Down}")
    if InStr(mods, "#")
        Send("{Blind}{LWin Down}")
}

SendModifiersUp(mods) {
    if InStr(mods, "^")
        Send("{Blind}{LCtrl Up}")
    if InStr(mods, "!")
        Send("{Blind}{LAlt Up}")
    if InStr(mods, "+")
        Send("{Blind}{LShift Up}")
    if InStr(mods, "#")
        Send("{Blind}{LWin Up}")
}

GuiClose(*) => ExitApp()