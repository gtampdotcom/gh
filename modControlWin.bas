Attribute VB_Name = "modControlWin"
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
'Public Const WM_ACTIVATE = &H6
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Const WM_VSCROLL = &H115
Public Const SB_BOTTOM = 7
Public Const WM_COMMAND = &H111
Public Const WM_SETTEXT = &HC
Public Const EM_SETSEL = &HB1

Public intPrevListCount As Integer

Public lngProcessID As Long
Public lngHandleGTA2 As Long
Public lngHandleLV As Long
Public lngHandleHistory As Long
Public lngHandleJoinHistory As Long
Public lngHandleReject As Long
Public lngHandleStart As Long
Public lngHandleCancel As Long
Public lngHandleChat As Long
'Public lngHandleJoinChat As Long
Public lngHandleSend As Long
Public lngHandleMaps As Long
Public lngHandlePlayersRequired As Long
Public lngHandleGameSpeed As Long
Public lngHandleSpeed As Long
Public lngHandleGameType As Long
Public lngHandleScoreLimit As Long
Public lngHandleTimeLimit As Long
Public lngHandleCops As Long
Public lngHandleScoreLimitLabel As Long
Public lnghandleTimeLimitLabel As Long

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)

Public Const WM_USER = &H400
Public Const TBM_SETPOS = (WM_USER + 5)
Public Const BM_CLICK As Long = &HF5
Public Const WM_GETTEXT As Integer = &HD
Public Const WM_GETTEXTLENGTH As Long = &HE

Public Const GWL_ID = (-12)

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32.dll" _
(ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Function EnumChildWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    On Error GoTo oops
    EnumChildWindowsProc = 1
    
    Select Case Val(GetWindowLong(hwnd, GWL_ID))
        Case 1024
            lngHandleLV = hwnd
        Case 1022
            lngHandleHistory = hwnd
        Case 1051
            lngHandleJoinHistory = hwnd
        Case 1020
            lngHandleReject = hwnd
        Case 1021
            lngHandleStart = hwnd
        Case 2
            lngHandleCancel = hwnd
        Case 1025
            lngHandleChat = hwnd
        'Case 1053
        '    lngHandleJoinChat = hwnd
        Case 1023
            lngHandleSend = hwnd
        Case 1026
            lngHandleMaps = hwnd
            'Call getMapDescFromLV
        Case 1033
            lngHandlePlayersRequired = hwnd
        'Case 1031
        '    lngHandleSpeed = hWnd
        Case 1032
             lngHandleGameSpeed = hwnd
        Case 1036
            lngHandleGameType = hwnd
        Case 1059
            lngHandleScoreLimit = hwnd
        Case 1038
            lngHandleTimeLimit = hwnd
        Case 1027
            lngHandleCops = hwnd
        Case 1035
            lngHandleScoreLimitLabel = hwnd
        Case 1037
            If IsWindowVisible(hwnd) Then
                blnReadyForJoiners = True
                Dim intListCount As Integer
                Dim strString As String
                If intPlayerCount <> 0 Then
                    intListCount = GetListViewCount(lngHandleLV)
                    If intListCount = 0 Then Exit Function
                    strString = intListCount & "/" & intPlayerCount
                    If strStatusMsg <> strString And strStatusMsg <> "A" And strStatusMsg <> "=HW" Then
                        If blnchkSoundJoin = True And intListCount > 1 And intPrevListCount <> intListCount Then
                            intPrevListCount = intListCount
                            PlaySound strSoundJoin, ByVal 0&, SND_ASYNC
                        End If
                        
                        'Clear strStatusMsg if it's not =AFK or =HW or anything with =
                        If InStr(strStatusMsg, "=") = 0 Then
                            strStatusMsg = strString
                            frmGH.changeStatus strString
                            send "NOTICE " & strChannel & " S" & strStatusMsg
                        End If
                    End If
                End If
            Else
                blnReadyForJoiners = False
            End If
            lnghandleTimeLimitLabel = hwnd
    End Select

Exit Function

oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "EnumChildWindowsProc: " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :EnumChildWindowsProc " & strErrdesc & " Line: " & strErrLine
End Function

'Some of this function was from here: http://www.ex-designz.net/apidetail.asp?api_id=316
Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
On Error GoTo oops
'Static winnum As Integer ' counter keeps track of how many windows have been enumerated
'winnum = winnum + 1 ' one more window enumerated....
''Rebug.Print winnum
EnumWindowsProc = 1 ' return value of 1 means continue enumeration

Dim lngTemp As Long

Call GetWindowThreadProcessId(hwnd, lngTemp)
If lngTemp = lngPID Then
    lngProcessID = lngTemp
Else
    Exit Function
End If

'Dim hwndTarget As Long
'hwndTarget = FindWindow(vbNullString, "Game Hunter v1.548")
'Const GWL_HINSTANCE = (-6)
'Dim hInstance As Long
'hInstance = GetWindowLong(hwndTarget, GWL_HINSTANCE)
'Debug.Print hInstance & " " & WindowTitle(hwndTarget) & " " & App.hInstance
'SendMessage hwndTarget, WM_SETTEXT, 0, ByVal "test: " & Rand(0, 999)


If InStr(WindowTitle(hwnd), "GTA2") Then
    
    Select Case ClassName(hwnd)
        Case "#32770"
            If blnchkTime = True Then
                If frmGH.timStamp.Enabled = False Then frmGH.timStamp.Enabled = True
            End If
            
            If frmGH.timUpdateMap.Enabled = False Then frmGH.timUpdateMap.Enabled = True
            blnLobby = True
            blnInGame = False
            lngProcessID = lngTemp
            lngWindowHandle = hwnd
            EnumChildWindows hwnd, AddressOf EnumChildWindowsProc, 1

        Case "WinMain"
            blnLobby = False
            blnInGame = True
            blnReadyForJoiners = False
            frmGH.timUpdateMap.Enabled = False
            frmGH.timStamp.Enabled = False
            EnumWindowsProc = 0
    End Select
End If

Exit Function

oops:
    Call ErrorHandler("EnumWindowsProc", Err.Description, Erl)


'370       strErrdesc = Err.Description
'380       strErrLine = Erl
'390       displaychat strDestTab, vbRed, "EnumWindowsProc: " & strErrdesc
'400       send "PRIVMSG " & gta2ghbot & " :EnumWindowsProc " & strErrdesc & " Line: " & strErrLine
End Function

Public Function WindowTitle(ByVal lHwnd As Long) As String
On Error GoTo oops

Dim slength As Long
Dim Buffer As String
Dim retval As Long

slength = GetWindowTextLength(lHwnd) + 1 ' get length of title bar text
If slength > 1 Then ' if return value refers to non-empty string
    Buffer = Space(slength) ' make room in the buffer
    retval = GetWindowText(lHwnd, Buffer, slength) ' get title bar text
    WindowTitle = Left(Buffer, slength - 1) ' display title bar text of enumerated window
    'frmHax.txtHistory.Text = frmHax.txtHistory.Text & WindowTitle & vbNewLine
End If
   
Exit Function

oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "WindowTitle error: " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :WindowTitle " & strErrdesc & " Line: " & strErrLine
End Function

Public Function ClassName(ByVal lHwnd As Long) As String
On Error GoTo oops
Dim lLen As Long
Dim sBuf As String
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lHwnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If

Exit Function

oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "ClassName error: " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :ClassName " & strErrdesc & " Line: " & strErrLine
End Function

Private Function GetListViewCount(ByVal hwnd As Long) As Long
    'this simply get number of items
    GetListViewCount = SendMessage(hwnd, LVM_GETITEMCOUNT, 0, ByVal 0)
End Function


Private Sub getMapDescFromLV()

Const CB_GETCURSEL = &H147
Const CB_GETLBTEXT = &H148
Const CB_GETLBTEXTLEN = &H149
Const CB_GETCOUNT = &H146
Dim count As Long       ' number of items in the combo box
count = SendMessage(lngHandleMaps, CB_GETCOUNT, ByVal CLng(0), ByVal CLng(0)) - 1

' Display the text of whatever item in combo box Combo1
' is currently selected.  If no list box item is selected, say so.
Dim Index As Long       ' index to the selected item
Dim itemtext As String  ' the text of the selected item
Dim textlen As Long     ' the length of the selected item's text

' Determine the index of the selected item.
Index = SendMessage(lngHandleMaps, CB_GETCURSEL, ByVal CLng(0), ByVal CLng(0))
textlen = SendMessage(lngHandleMaps, CB_GETLBTEXTLEN, ByVal CLng(Index), ByVal CLng(0))
' Make enough room in the string to receive the text, including the terminating null.
itemtext = Space(textlen) & vbNullChar
' Retrieve that item's text and display it.
textlen = SendMessage(lngHandleMaps, CB_GETLBTEXT, ByVal CLng(Index), ByVal itemtext)
itemtext = Left(itemtext, textlen)
'Debug.Print itemtext

End Sub
