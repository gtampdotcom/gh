Attribute VB_Name = "modDetectURL"
Option Explicit

'Hyperlink detection code copied from iQWordPad
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=69067

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
'Const SEE_MASK_INVOKEIDLIST = &HC
'Const SEE_MASK_NOCLOSEPROCESS = &H40
'Const SEE_MASK_FLAG_NO_UI = &H400
'Type SHELLEXECUTEINFO
'       cbSize As Long
'       fMask As Long
'       hwnd As Long
'       lpVerb As String
'       lpFile As String
'       lpParameters As String
'       lpDirectory As String
'       nShow As Long
'       hInstApp As Long
'       lpIDList As Long
'       lpClass As String
'       hkeyClass As Long
'       dwHotKey As Long
'       hIcon As Long
'       hProcess As Long
'End Type
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long

'for URL Linking
Private Const EM_CHARFROMPOS& = &HD7

Private Type POINTAPI
    x As Long
    y As Long
End Type


Private Const WM_USER = &H400
'Private Const WM_NCLBUTTONDOWN = &HA1
'Private Const HTBOTTOMRIGHT = 17
'Private Const EM_AUTOURLDETECT = (WM_USER + 91)

'Public Sub EnableAutoURLDetection(rtb As RichTextBox)
'    'enable auto URL detection
'    SendMessage rtb.hwnd, EM_AUTOURLDETECT, 1&, ByVal 0&
'End Sub
' Return the word the mouse is over.
Public Function RichWordOver(rch As RichTextBox, x As Single, y As Single) As String
Dim pt As POINTAPI
Dim pos As Long
Dim start_pos As Long
Dim end_pos As Long
Dim ch As String
Dim txt As String
Dim txtlen As Long

On Error GoTo oops

    ' Convert the position to pixels.
    pt.x = x \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(rch.hwnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then Exit Function
    
    ' Find the start of the word.
    txt = rch.Text
    For start_pos = pos To 1 Step -1
        ch = Mid$(rch.Text, start_pos, 1)
        If ch = " " Or ch = vbCr Or ch = vbTab Then Exit For
    Next start_pos
    
    start_pos = start_pos + 1

    ' Find the end of the word.
    txtlen = Len(txt)
    
    For end_pos = pos To txtlen
        ch = Mid$(txt, end_pos, 1)
        If ch = " " Or ch = vbCr Or ch = vbTab Then Exit For
    Next end_pos
    
    end_pos = end_pos - 1
    
    If start_pos <= end_pos Then
        RichWordOver = Mid$(txt, start_pos, end_pos - start_pos + 1)
    End If
    
    RichWordOver = cleanURL(RichWordOver)
       
    Exit Function
oops:
    strErrdesc = Err.Description
    displaychat strDestTab, vbRed, "Hyperlink detection error: " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :Hyperlink detection error: " & strErrdesc
End Function

Public Function cleanURL(ByVal strURL As String) As String
    Dim i As Integer
    Dim vbComma As String
    vbComma = Chr$(44)
    
    'Remove these characters from the end of URL
    For i = Len(strURL) To 4 Step -1
        Select Case Right$(strURL, 1)
            Case ".", vbComma, ";", "-", "{", "}", "<", ">", "?", "'", vbQuote '"(", ")"
                strURL = Left$(strURL, Len(strURL) - 1)
            Case Else
                Exit For
        End Select
    Next
    
     'Remove these characters from the start of URL
    For i = 1 To Len(strURL)
        Select Case Left$(strURL, 1)
            Case ".", vbComma, ";", "-", "{", "}", "<", ">", "(", ")", "?", "!", "'", vbQuote
                strURL = Right$(strURL, Len(strURL) - 1)
            Case Else
                Exit For
        End Select
    Next
    
    cleanURL = strURL
End Function
