Attribute VB_Name = "modFiles"
Option Explicit

Private Declare Function GetTempPath Lib "Kernel32" Alias _
"GetTempPathA" (ByVal nBufferLength As Long, ByVal _
lpBuffer As String) As Long

Public Function Exists(ByVal File As String) As Boolean
On Error GoTo notfound
   
    'Not using FSO since it didn't work in Wine (it probably can but requires extra files)
    'Dim FS As New FileSystemObject
    'If FS.FileExists(File) Then
    'If (GetAttr(File) And vbDirectory) = vbDirectory Then
    '    Exists = True
    '    Exit Function
    'End If
    
    Dim strString As String
    Dim i As Integer
    For i = Len(File) To 1 Step -1
        If Mid$(File, i, 1) = "\" Then
            strString = Mid$(File, i + 1, 666)
            Exit For
        End If
    Next i
    
    If File = vbNullString Or strString = vbNullString Then
        Exists = False
        Exit Function
    End If
    
    'nested/threaded Dir won't work
    If UCase(Dir(File, vbHidden + vbNormal + vbSystem + vbReadOnly)) = UCase(strString) Then
        Exists = True
    Else
        Exists = False
    End If
    
    Exit Function
notfound:
    Exists = False
    displaychat strChannel, vbRed, File & " " & Err.Description
End Function

'Public Function modFileOpen(filename As String, access_mode As String)
'On Error GoTo oops
'modFileOpen = 1
'
'Select Case access_mode
'    Case "output"
'        Open filename For Output As #1
'    Case "input"
'        Open filename For Input As #1
'End Select
'Exit Function
'oops:
'    modFileOpen = 0
'    strErrdesc = Err.Description
'    displaychat strDestTab, vbRed, "Unable to open " & filename & " for " & access_mode & ".  Error " & strErrdesc
'End Function

Public Sub modFileCopy(source As String, Destination As String)
On Error GoTo oops

FileCopy source, Destination
Exit Sub
oops:
    strErrdesc = Err.Description
    displaychat strDestTab, vbRed, "Unable to copy " & source & " to " & Destination & ".  Error " & strErrdesc
End Sub

'Public Function rename(filename As String, newfilename As String) As Boolean
'On Error GoTo oops
'rename = True
'Name filename As newfilename
'Exit Function
'oops:
'    rename = False
'    strErrdesc = Err.Description
'End Function

Public Function modFileKill(FileName As String) As Boolean
On Error GoTo oops
modFileKill = True
Kill FileName
Exit Function
oops:
    modFileKill = False
    strErrdesc = Err.Description
End Function

Public Sub modMkDir(FileName As String)
On Error GoTo oops
    MkDir FileName
Exit Sub
oops:
    strErrdesc = Err.Description
    'displaychat strDestTab, vbRed, "Unable to create " & filename & ".  Error " & strErrdesc
End Sub

Public Sub MoveMMPfiles()
On Error Resume Next
    Dim oFileSystem As New FileSystemObject
    Dim oFolder As Folder
    Dim oCurrentFile As File
    Dim oFileColl As Files

    Set oFolder = oFileSystem.GetFolder(strGTA2path & "data\tempMMP\")
    Set oFileColl = oFolder.Files

    If oFileSystem.FolderExists(oFolder) = False Then Exit Sub

    displaychat strChannel, strGHColor, "Moving MMP files from " & strGTA2path & "data\tempMMP\ to " & strGTA2path & "data\"

    'Move all files in gta2\data\tempMMP to gta2\data
    If oFileColl.count > 0 Then
        For Each oCurrentFile In oFileColl
            oFileSystem.MoveFile strGTA2path & "data\tempMMP\" & oCurrentFile.Name, strGTA2path & "data\" & oCurrentFile.Name
        Next
    End If

    oFileSystem.DeleteFolder oFolder

    Set oFileSystem = Nothing
    Set oFolder = Nothing
    Set oFileColl = Nothing
    Set oCurrentFile = Nothing
    
'Dir is faster than FSO and works in Wine but I don't know about Copy and I want to move anyway
'    'If gta2\data\tempMMP exists then move all MMP files into gta2\data folder
'    If Dir(strGTA2path & "data\tempMMP\*.mmp") <> vbNullString Then
'    'Copy data\*.mmp to data\tempMMP
'        strString = Dir(strGTA2path & "data\tempMMP\*.mmp")
'        Do While strString <> vbNullString
'            modFileCopy (strGTA2path & "data\tempMMP\" & strString), (strGTA2path & "data\" & strString)
'            strString = Dir
'        Loop
'    End If
    
    Exit Sub
oops:
    strErrdesc = Err.Description
    displaychat strDestTab, vbRed, "Unable to move " & oCurrentFile.Name & ".  Error " & strErrdesc
End Sub

' Keywords: Get Temporary Folder, Temporary Folder Visual Basic Code, VB Function Get Temp Folder, VBA Temporary Folder, VB6 Temporary Folder, GetTempPath, Windows API Functions

Public Function GetTmpPath()

Dim sFolder As String ' Name of the folder
Dim lRet As Long ' Return Value

sFolder = String(MAX_PATH, 0)
lRet = GetTempPath(MAX_PATH, sFolder)

If lRet <> 0 Then
GetTmpPath = Left(sFolder, InStr(sFolder, _
Chr(0)) - 1)
Else
GetTmpPath = vbNullString
End If

End Function

Public Sub setGTA2path()

Dim cr As New cRegistry

With cr 'read GTA2 path from registry
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter"
    .ValueKey = "GTA2folder"
    If LenB(.Value) > 0 Then
        strGTA2path = .Value
        If strGTA2path = "\\" Or strGTA2path = "\" Or strGTA2path = "/" Then
            strGTA2path = vbNullString
        Else
            If Right$(strGTA2path, 1) <> "\" Then strGTA2path = strGTA2path & "\"
        End If
    End If
    
    .ClassKey = HKEY_CLASSES_ROOT
    .SectionKey = "GTA2"
    .ValueKey = vbNullString
    .Value = "GTA2 GTA2GH"
    .ValueKey = "URL Protocol"
    .Value = "gta2://"
    .SectionKey = "GTA2\DefaultIcon"
    .ValueKey = vbNullString
    .Value = App.Path & "\" & App.EXEName
    .SectionKey = "GTA2\Shell\Open\Command"
    .Value = App.Path & "\" & App.EXEName & " %1"
'[HKEY_CLASSES_ROOT\gta2]
'@="GTA2GH gta2"
'[HKEY_CLASSES_ROOT\gta2\DefaultIcon]
'@="C:\\gh\\15\\gta2gh.exe"
'[HKEY_CLASSES_ROOT\gta2\shell]
'@="open"
'[HKEY_CLASSES_ROOT\gta2]
'"Url Protocol"=""
'[HKEY_CLASSES_ROOT\gta2\shell\open\command]
'@="C:\\gh\\15\\gta2gh.exe %1"
   
    
End With

Set cr = Nothing

End Sub

'Public Function writeTest() As Boolean
'
'writeTest = False
'
'If App.LogMode = False Then 'we are running in the VB6 IDE
'    If InStr(GetCommandOutput("c:\path\7za.exe a " & vbQuote & strGTA2path & "gta2ghwritetest.7z" & vbQuote & " " & vbQuote & strGTA2path & "readme.txt" & vbQuote, True, False, True), "Everything is Ok") Then
'        writeTest = True
'    End If
'Else
'    If InStr(GetCommandOutput(App.Path & "\7za.exe a " & vbQuote & strGTA2path & "gta2ghwritetest.7z" & vbQuote & " " & vbQuote & strGTA2path & "readme.txt" & vbQuote, True, False, True), "Everything is Ok") Then
'        writeTest = True
'    End If
'End If
'
'End Function
