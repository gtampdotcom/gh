Attribute VB_Name = "modDownload"
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'Download a web page on the internet (works with proxy servers)
'
'The following routine uses API calls to read/download an internet file or an html page from a remote web site. The code will work with a proxy server and a routine demonstrating how to use this code can be found at the bottom of the post.

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias _
    "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, _
    ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As Long

'http://www.visualbasic.happycodings.com/Internet_Web_Mail_Stuff/code9.html
'Purpose     :  Get text from a web site
'Inputs      :  sServerName             The server name where the file is located eg.
'               sFileName               The file name to download eg. "index.asp" or "/code/codetoc.asp"
'               [sUsername]             If required, the login user name.
'               [sPassword]             If required, the user's password.
'Outputs     :  The contents of the specified file
'Notes       :  Can be used through a proxy server by specifying a username and password
'Revisions   :

Function CopyURLToFile(ByVal URL As String, Optional ByVal FileName As String) As String
    Dim i As Integer
    Dim strString As String
    Dim lngDownloadSize As Long
    Dim FileNum As Integer
    Dim ok As Boolean
    Dim NumberOfBytesRead As Long
    Dim Buffer As String
    Dim fileIsOpen As Boolean
    Dim bln404 As Boolean
    Dim strFile As String
    
    Dim hInternetSession As Long, hInternetConnect As Long, hHttpOpenRequest As Long
    Dim lRetVal As Long, lLenFile As Long, lNumberOfBytesRead As Long, lResLen As Long
    Dim sBuffer As String, lTotalBytesRead As Long
    
    Static blnDownloading As Boolean
    
    Const scUserAgent As String = "GH" & TXT_GHVER
    Const INTERNET_OPEN_TYPE_PRECONFIG = 0, INTERNET_FLAG_EXISTING_CONNECT = &H20000000
    Const INTERNET_OPEN_TYPE_DIRECT = 1, INTERNET_OPEN_TYPE_PROXY = 3
    Const INTERNET_DEFAULT_HTTP_PORT = 80, INTERNET_FLAG_RELOAD = &H80000000
    Const INTERNET_SERVICE_HTTP = 3
    Const HTTP_QUERY_CONTENT_LENGTH = 5
    
    On Error GoTo ErrorHandler
    
    If lngMaster = 0 Then 'if this is the master GH process then launch a 2nd process to handle the download
        Dim lngGHPID As Long
        Dim lngCursor As Long
        lngCursor = frmGH.MousePointer
        frmGH.MousePointer = vbHourglass
        
        If App.EXEName = "prjGH" Then 'we are running in the VB6 IDE
            lngGHPID = shellandwait("c:\gh\15\gta2gh.exe" & " -m " & frmGH.txtSlave.hwnd & " -d " & URL, "c:\gh\15")
        Else
            lngGHPID = shellandwait(App.Path & "\" & App.EXEName & " -m " & frmGH.txtSlave.hwnd & " -d " & URL, App.Path)
            'ShellExecute(Me.hwnd, "Open", strChatExcludingCommand, vbNullString, Mid$(strChatExcludingCommand, 1, InStrRev(strChatExcludingCommand, "\")), vbNormalFocus) = 2 Then
        End If
        frmGH.MousePointer = lngCursor
        Exit Function
    End If
    
    If URL = "http://gtamp.com/gta2.7z" Then
        If Exists(strGTA2path & "data\audio\wil.raw") = False Then
            strGTA2path = DOCUMENTS & "\gta2\"
            Dim cr As New cRegistry
            With cr
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                .ValueType = REG_SZ
                .ValueKey = "GTA2Folder"
                .Value = strGTA2path
            End With
            Set cr = Nothing
        End If
    Else
        If DetectGTA2version = False Then
            strString = "Download cancelled. Select your GTA2 folder in settings"
            If lngMaster = 1 Then
                MsgBox strString
            Else
                i = SendMessage(lngMaster, WM_SETTEXT, 0, ByVal strString)
            End If
            End
        End If
    End If
    
    If Exists(App.Path & "\7za.exe") = False Then
        strString = "Download cancelled. Failed to find " & App.Path & "\7za.exe"
        MsgBox strString
        If lngMaster <> 1 Then
            If lngMaster Then End
            i = SendMessage(lngMaster, WM_SETTEXT, 0, ByVal strString)
        End If
        End
    End If
    
    If blnDownloading = True Then
        If lngMaster <> 1 Then
            i = SendMessage(lngMaster, WM_SETTEXT, 0, ByVal "A file is already being downloaded.")
            If lngMaster Then End
        End If
        End
    End If
    
    blnDownloading = True
    
    'Initializes an application's use of the Win32 Internet functions
    hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

    'Opens an FTP, Gopher, or HTTP session for a given site
    
    Dim sServerName As String, sUsername As String, sPassword As String, sFileName As String
    
    sServerName = URL
    
    'remove http:// from URL
    If LCase$(Left$(URL, 7)) = "http://" Then
        sServerName = Right$(URL, Len(URL) - 7)
    End If
    
    If LCase$(Left$(URL, 8)) = "https://" Then
        sServerName = Right$(URL, Len(URL) - 8)
    End If
    
    'Split URL into servername and filename
    i = InStr(sServerName, "/")
    If i Then
        sFileName = Mid$(sServerName, i, 666)
        sServerName = Left$(sServerName, i - 1)
    Else
        sFileName = "/"
    End If
    
    hInternetConnect = InternetConnect(hInternetSession, sServerName, INTERNET_DEFAULT_HTTP_PORT, sUsername, sPassword, INTERNET_SERVICE_HTTP, 0, 0)
    'Create an HTTP request handle
    hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", sFileName, "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    
    'Creates a new HTTP request handle and stores the specified parameters in that handle

    lRetVal = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, 0, 0)
    
    If lRetVal Then
        'Determine the file size
        sBuffer = Space(65536)
'710           sBuffer = Space(1024)
        lResLen = Len(sBuffer)
        lRetVal = HttpQueryInfo(hHttpOpenRequest, HTTP_QUERY_CONTENT_LENGTH, ByVal sBuffer, Len(sBuffer), 0)
        
        If lRetVal Then
            'Successfully returned file length
            lLenFile = Val(Left$(sBuffer, lResLen))
        Else
            'Unable to establish file length
            lLenFile = -1
        End If
    End If
        
    strFile = Right$(URL, Len(URL) - InStrRev(URL, "/"))
    
    If lLenFile <> -1 Then
        'ensure that there is no local file
        On Error Resume Next
    
        If FileName <> vbNullString Then Kill FileName
    
        'On Error GoTo ErrorHandler
            
        'open the local file
        FileNum = FreeFile
        Open FileName For Binary As FileNum
        fileIsOpen = True
       
        If lngMaster <> 1 Then
            Call SendMessage(lngMaster, WM_SETTEXT, 0, ByVal "Downloading " & strFile)
        End If
        
        Dim lngStartTime As Long
        Dim lngTotalTime As Long
        lngStartTime = GetTickCount
          
        'Display download window so download can be cancelled
        If lngMaster = 1 Then
            frmDownload.Show
            frmDownload.Icon = frmGH.Icon
        End If
           
        'Read the file
        Do
            'DoEvents 'allows download to be cancelled
            If blnCancel = True Then Exit Do
            
            'Store the results
            lngTotalTime = FormatNumber((GetTickCount - lngStartTime) / 1000, 0, vbFalse, vbFalse, vbFalse)
            If lTotalBytesRead > 0 And lngTotalTime > 0 Then lngSpeed = (lTotalBytesRead / lngTotalTime) / 1000
            lRetVal = InternetReadFile(hHttpOpenRequest, sBuffer, Len(sBuffer), lNumberOfBytesRead)
            lTotalBytesRead = lTotalBytesRead + lNumberOfBytesRead
            'Change the titlebar of the master GH process to show download progress
            strString = Int(lTotalBytesRead / 1000) & "kB/" & Int(lLenFile / 1000) & "kB at " & lngSpeed & "kB/s (" & Int((lTotalBytesRead / lLenFile) * 100) & "%)"
            frmDownload.Caption = "Downloading " & strFile & " " & strString
            If lngMaster <> 1 Then
                i = SendMessage(lngMaster, WM_SETTEXT, 0, ByVal "Game Hunter v" & TXT_GHVER & " - " & strString)
                If i <> 1 Then End
            End If
            'frmGH.Caption = "Game Hunter v" & TXT_GHVER & " - " & Int(lTotalBytesRead / 1000) & "kB/" & Int(lLenFile / 1000) & "kB at " & lngSpeed & "kB/s (" & Int((lTotalBytesRead / lLenFile) * 100) & "%)"
            lngDownloadSize = lTotalBytesRead
            
            If Left$(sBuffer, 2) = "<!" Then
                bln404 = True
                Exit Do
            End If
            
            'save the data to the local file
            Put #FileNum, , Left$(sBuffer, lNumberOfBytesRead)
            
            'Finished reading file
            If lNumberOfBytesRead = 0 Or lTotalBytesRead = lLenFile Or lRetVal = 0 Then
                Exit Do
            End If
        Loop
       
        lTotalBytesRead = Int(lTotalBytesRead / 1000)
        If lngSpeed = 0 Then lngSpeed = lTotalBytesRead
    Else
        bln404 = True
    End If
    
    If lngMaster <> 1 Then
        i = SendMessage(lngMaster, WM_SETTEXT, 0, ByVal "Game Hunter v" & TXT_GHVER)
    End If
    ' flow into the error handler

ErrorHandler:
        
'1260      Open "c:\test.txt" For Output As #1
'1270      Print #1, "Error: " & Err.Description & " " & Erl
'1280      Close #1
    
    frmGH.cmdToolbar(BTN_CANCEL).Visible = False
    lngDownloadSize = 0
    ' close the local file, if necessary
    If fileIsOpen Then Close #FileNum

    'Close handles
    InternetCloseHandle hHttpOpenRequest
    InternetCloseHandle hInternetSession
    InternetCloseHandle hInternetConnect
    blnDownloading = False

    ' report the error to the client, if there is one
    'If Err Then Err.Raise Err.Number, , Err.Description

    If bln404 = True Then
        If InStr(LCase(URL), "gtamp.com") Then
            strString = strFile & " isn't on gtamp.com/maps. Check https://gtamp.com/mapscript/maplist/download.php?mmp=" & LCase(Replace(strFile, " ", "%20")) & ".mmp"
            If lngMaster <> 1 Then
                Call SendMessage(lngMaster, WM_SETTEXT, 0, ByVal strString)
            Else
                MsgBox strString
            End If
        Else
            strString = strFile & " is not at that address."
            frmDownload.Caption = strString
            If lngMaster <> 1 Then
                i = SendMessage(lngMaster, WM_SETTEXT, 0, ByVal strString)
            Else
                MsgBox strString
            End If
        End If
        End
    End If

    If blnCancel = True Then
        strString = "Download cancelled " & strFile & " - " & lTotalBytesRead & "kB in " & Duration(lngTotalTime, 2) & "(" & lngSpeed & "kB/s)"
        frmDownload.Caption = strString
        If lngMaster <> 1 Then
            i = SendMessage(lngMaster, WM_SETTEXT, 0, ByVal strString)
        End If
        blnCancel = False
    Else
        strString = "Download complete: " & strFile & " - " & lTotalBytesRead & "kB in " & Duration(lngTotalTime, 2) & "(" & lngSpeed & "kB/s)"
        frmDownload.Caption = strString
        If lngMaster <> 1 Then
            i = SendMessage(lngMaster, WM_SETTEXT, 0, ByVal strString)
        End If
        Dim strFolder As String
        Dim strFileType As String
        strFolder = "data\"
        strFileType = "Map"
        If URL = "http://gtamp.com/gta2.7z" Or URL = "http://gtamp.com/gta2patch.7z" Or URL = "http://127.0.0.1/gta2.7z" Then
            strFolder = vbNullString
            strFileType = "GTA2"
        End If
        
        If InStr(GetCommandOutput(App.Path & "\7za.exe x -y -o" & vbQuote & strGTA2path & strFolder & vbQuote & " " & vbQuote & GetTmpPath & "gta2map.7z" & vbQuote, True, False, True), "Everything is Ok") Then
            strString = strFileType & " successfully installed."
            frmDownload.Caption = frmDownload.Caption & " " & strString
            If lngMaster <> 1 Then SendMessage lngMaster, WM_SETTEXT, 0, ByVal strString
        Else
            strString = Replace(Replace(Replace(Replace(Replace(Replace(GetCommandOutput(App.Path & "\7za.exe x -y -o" & vbQuote & strGTA2path & "data\" & vbQuote & " " & vbQuote & GetTmpPath & "gta2map.7z" & vbQuote, True, False, True), "7-Zip (A) 9.20  Copyright (c) 1999-2010 Igor Pavlov  2010-11-18", vbNullString), vbNewLine, " "), "   ", vbNullString), "Extracting ", vbNullString), "  ", " "), "  ", " ")
            If lngMaster <> 1 Then
                SendMessage lngMaster, WM_SETTEXT, 0, ByVal strString
            Else
                MsgBox strString
            End If
            strString = strFileType & " failed to install."
            frmDownload.Caption = strString
            If lngMaster <> 1 Then SendMessage lngMaster, WM_SETTEXT, 0, ByVal strString
        End If
    End If
    Sleep 5000
    End
End Function
