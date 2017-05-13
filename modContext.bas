Attribute VB_Name = "modContext"
Option Explicit

'==================================================
' FrmMain.Rich1_DblClick()
Declare Function ReleaseCapture Lib "user32" () As Long

' FrmMain.Form_Resize()
Declare Function MoveWindow Lib "user32" _
                            (ByVal hWnd As Long, _
                            ByVal x As Long, _
                            ByVal y As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            ByVal bRepaint As Long) As Long
                            
'=====================================================
' FrmMain.Form_KeyDown()

Declare Function GetClientRect Lib "user32" _
                            (ByVal hWnd As Long, _
                            lpRect As RECT) As Long

Declare Function GetCursorPos Lib "user32" _
                            (lpPoint As POINTAPI) As Long

Declare Function GetWindowRect Lib "user32" _
                            (ByVal hWnd As Long, _
                            lpRect As RECT) As Long

Declare Function InvalidateRect Lib "user32" _
                            (ByVal hWnd As Long, _
                            ByVal lpRect As Long, _
                            ByVal bErase As Long) As Long

Declare Function PtInRect Lib "user32" _
                            (lpRect As RECT, _
                            ByVal ptX As Long, _
                            ByVal ptY As Long) As Long

Declare Function ScreenToClient Lib "user32" _
                            (ByVal hWnd As Long, _
                            lpPoint As POINTAPI) As Long

Type POINTAPI
    x As Long
    y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'=====================================================

Public Const WM_SETREDRAW = &HB
'Public Const WM_GETTEXTLENGTH = &HE

Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303

Private Const WM_USER = &H400

Public Const EM_SCROLLCARET = &HB7
Public Const EM_REPLACESEL = &HC2
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_SETREADONLY = &HCF
Public Const EM_POSFROMCHAR = &HD6
Public Const EM_GETSELTEXT = (WM_USER + 62)

' RichEdit control specific messages,
' all handle selection beyond 64K (&HFFFF&)
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_GETTEXTRANGE = (WM_USER + 75)

' EM_FINDTEXTEX lParam
Type CHARRANGE   'cr
    cpMin As Long
    cpMax As Long
End Type
 
Type TEXTRANGE
  chrg As CHARRANGE
  lpstrText As String   ' allocated by caller, zero terminated by RichEdit
End Type

