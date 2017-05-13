Attribute VB_Name = "modSystray"
Option Explicit

'copied from http://cuinl.tripod.com/Tips/systray.htm

Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uID As Long
uFlags As Long
uCallbackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Global TaskIcon As NOTIFYICONDATA
Global Const NIM_ADD = &H0
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_MOUSEMOVE = &H200

Public Sub Systray()
    blnSystray = True
    TaskIcon.cbSize = Len(TaskIcon)
    TaskIcon.hwnd = frmGH.picTray.hwnd
    TaskIcon.uID = 1&
    TaskIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    TaskIcon.uCallbackMessage = WM_MOUSEMOVE
    TaskIcon.hIcon = frmGH.picTray
    TaskIcon.szTip = "Game Hunter" & vbNullChar
    Shell_NotifyIcon NIM_ADD, TaskIcon
End Sub

Public Sub drawTrayIcon()
    TaskIcon.hIcon = frmGH.picTray
    TaskIcon.szTip = "Game Hunter" & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, TaskIcon
End Sub
