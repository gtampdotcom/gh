Attribute VB_Name = "modVersion"
Option Explicit

'Used for getting the gta2.exe version information
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     ' e.g. = &h0000 = 0
   dwStrucVersionh As Integer     ' e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    ' e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    ' e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    ' e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    ' e.g. = &h0031 = .31
   dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
   dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
   dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
   dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
   dwFileFlagsMask As Long        ' = &h3F for version "0.42"
   dwFileFlags As Long            ' e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               ' e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             ' e.g. VFT_DRIVER
   dwFileSubtype As Long          ' e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           ' e.g. 0
   dwFileDateLS As Long           ' e.g. 0
End Type

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal source As Long, ByVal length As Long)

Public Function DetectGTA2version(Optional strProcessPath As String) As Boolean
    On Error GoTo oops
    
    Dim FileName As String, Directory As String, FullFileName As String
    Dim FileVer As String
    Dim rc As Long, lDummy As Long, sBuffer() As Byte
    Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
    Dim lVerbufferLen As Long
    
    Call setGTA2path
    
    If strProcessPath = vbNullString Then
        FileName = TXT_GTA2EXE
        Directory = strGTA2path
        FullFileName = Directory & FileName
    Else
        FullFileName = strProcessPath
        If LCase(Right$(strProcessPath, 8)) <> TXT_GTA2EXE Then Exit Function
    End If
    
    '** Get size **
    lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
    If lBufferLen < 1 Then
       Exit Function
    End If
    
    '** Store info to udtVerBuffer struct **
    ReDim sBuffer(lBufferLen)
    rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    
    '** Determine File Version number **
    FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)
    '** Determine Product Version number **
    ''ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)
    
     If Right$(FileVer, 2) = ".0" Then FileVer = Left$(FileVer, Len(FileVer) - 2)
     If Right$(FileVer, 2) = ".0" Then FileVer = Left$(FileVer, Len(FileVer) - 2)
    
    If strProcessPath <> vbNullString Then
        If FileVer = "7.1.100.1248" Then
            displaychat strChannel, strGHColor, "Derp! " & strProcessPath & " is the GTA2 classic installer. I'm not interested in that file."
            DetectGTA2version = False
        Else
            DetectGTA2version = True
        End If
        Exit Function
    End If
    
    strGTA2version = "?"
    strGTA2version = FileVer
    DetectGTA2version = True
     
     If Val(strGTA2version) < 11 Then
        If strProcessPath = vbNullString Then
            displaychat strDestTab, vbRed, "Update GTA2 to the latest version: https://gtamp.com/forum/viewtopic.php?f=4&t=73"
        End If
        
        DetectGTA2version = False
        
        'If exists(app.path & "\update.exe") Then
        '   modFileCopy app.path & "update.exe", strGTA2path & TXT_GTA2EXE
        'end if
     End If
     
     Exit Function
    
oops:
     strErrdesc = Err.Description
     displaychat strDestTab, strTextColor, "Error determining GTA2 version number: " & strErrdesc & " " & Erl
     send "PRIVMSG " & gta2ghbot & " :Error determining GTA2 version number: " & strErrdesc
    End Function
