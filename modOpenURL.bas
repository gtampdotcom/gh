Attribute VB_Name = "modOpenURL"
Option Explicit

Const INTERNET_FLAG_RELOAD = &H80000000
Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
    (ByVal lpszAgent As String, ByVal dwAccessType As Long, _
    ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, _
    ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias _
    "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, _
    ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, _
    ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As _
    Long) As Integer
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As _
    Long, ByVal lpBuffer As String, ByVal dwNumberOfBytesToRead As Long, _
    lNumberOfBytesRead As Long) As Integer

' Download a file from Internet and save it to a local file
'
' it works with HTTP and FTP, but you must explicitly include
' the protocol name in the URL, as in
'    CopyURLToRAM "http://gtamp.com/server.txt")

Function CopyURLToRAM(ByVal URL As String) As String
    On Error Resume Next
    Dim hInternetSession As Long
    Dim hURL As Long
    Dim ok As Boolean
    Dim NumberOfBytesRead As Long
    Dim Buffer As String
    
    ' check obvious syntax errors
    If Len(URL) = 0 Then Err.Raise 5

    ' open an Internet session, and retrieve its handle
    hInternetSession = InternetOpen("GH" & TXT_GHVER, INTERNET_OPEN_TYPE_PRECONFIG, _
        vbNullString, vbNullString, 0)
    If hInternetSession = 0 Then Err.Raise vbObjectError + 1000, , _
        "An error occurred calling InternetOpen function"
        
    hURL = InternetOpenUrl(hInternetSession, URL, vbNullString, 0, _
        INTERNET_FLAG_RELOAD, 0)
        
    ' prepare the receiving buffer
    Buffer = Space(256)
    ok = InternetReadFile(hURL, Buffer, Len(Buffer), NumberOfBytesRead)
    CopyURLToRAM = Left$(Buffer, NumberOfBytesRead)
     ' close internet handles, if necessary
    If hURL Then InternetCloseHandle hURL
    If hInternetSession Then InternetCloseHandle hInternetSession
End Function
