Attribute VB_Name = "modComdlg"
Option Explicit

Public PROGRAM_FILES As String
Public WINDOWS_FOLDER As String
Public DOCUMENTS As String

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

       Public Type OPENFILENAME
         lStructSize As Long
         hwndOwner As Long
         hInstance As Long
         lpstrFilter As String
         lpstrCustomFilter As String
         nMaxCustFilter As Long
         nFilterIndex As Long
         lpstrFile As String
         nMaxFile As Long
         lpstrFileTitle As String
         nMaxFileTitle As Long
         lpstrInitialDir As String
         lpstrTitle As String
         Flags As Long
         nFileOffset As Integer
         nFileExtension As Integer
         lpstrDefExt As String
         lCustData As Long
         lpfnHook As Long
         lpTemplateName As String
       End Type


Public Function BrowseFile(ByVal strPath As String, Optional ByVal blnWave As Boolean)
On Error GoTo oops
Dim OpenFile As OPENFILENAME

Dim strFile As String
    
    If strPath <> vbNullString And InStr(strPath, "\") Then
        'Search backwards until I find a \ for storing filename and initial folder
        strFile = Mid$(strPath, InStrRev(strPath, "\", , vbBinaryCompare) + 1, 666)
        strPath = Left$(strPath, InStrRev(strPath, "\", , vbBinaryCompare) - 1)
    Else
        If blnWave = True Then
            strPath = WINDOWS_FOLDER & "\media"
        Else
            strPath = PROGRAM_FILES
        End If
    End If
    
start:
    Dim lReturn As Long
    Dim sFilter As String
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = frmGH.hwnd
    OpenFile.hInstance = App.hInstance
    
    If blnWave = True Then
        sFilter = "Wave Files (*.wav)" & vbNullChar & "*.wav" & vbNullChar
    Else
        sFilter = "gta2.exe" & vbNullChar & "gta2.exe" & vbNullChar
    End If
    
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = strPath
    If blnWave = False Then
        OpenFile.lpstrTitle = "Select " & TXT_GTA2EXE
    Else
        OpenFile.lpstrTitle = "Select a wave file"
    End If
    OpenFile.Flags = 0
    lReturn = GetOpenFileName(OpenFile)
    If lReturn <> 0 Then
        BrowseFile = StripNull(OpenFile.lpstrFile)
        If blnWave = False Then
            If LCase(Right$(BrowseFile, Len(strFile))) <> strFile Then GoTo start
            If DetectGTA2version(StripNull(OpenFile.lpstrFile)) = False Then GoTo start
        End If
    End If
    Exit Function
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Browse for file error: " & strErrdesc & " Line: " & strErrLine
End Function

Public Function StripNull(ByVal InString As String) As String

'Input: String containing null terminator (vbNullChar)
'Returns: all character before the null terminator

Dim iNull As Integer
If Len(InString) > 0 Then
    iNull = InStr(InString, vbNullChar)
    Select Case iNull
    Case 0
        StripNull = InString
    Case 1
        StripNull = vbNullString
    Case Else
       StripNull = Left$(InString, iNull - 1)
   End Select
End If

End Function
