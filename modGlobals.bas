Attribute VB_Name = "modGlobalsOMG"
'Application scope:
'Any variable declared as Application Scope is available anywhere in your application, anytime, for whatever reason.
'Using Application-scoped variables is generally frowned upon, because the value can change anywhere and you have no
'real control over when they get changed, or in what order of code execution. They are handy for mostly static values,
'like dynamic Const's. You declare this type of variable in a Module, with the keyword Public.
'
'Uh Oh!

Option Explicit

Public Declare Function GetTickCount Lib "Kernel32" () As Long
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Public Const vbQuote = """"
Public Const TXT_COUNTRY_DETECTION_FAILED = "Country detection failed"
Public Const TXT_GEOSITE = "http://ip2c.org/s"
Public Const TXT_GTA2EXE = "gta2.exe"       'name of GTA2 executable
Public Const TXT_GHVER = "1.5991"            'GTA2 Game Hunter version number
Public Const TXT_YOUR_GAME_REMOVED = "Your game was removed from the list."
Public Const TXT_PRIVATE = "Private chat with "
Public Const gta2ghbot = "gta2ghbot"
Public Const MAX_PATH = 260

Public strOSV As String 'Operating system version
'Public strPlayerColors(255, 1) As String
Public lngMaster As Long 'hWnd of master GH txtSlave
Public blnChkSync As Boolean
Public intTheme As Integer
Public strRaw As String
Public lngWindowHandle As Long 'GTA2 window handle
Public lngPID As Long 'process ID of the gta2.exe launched by GH
Public lngSpeed As Long
Public blnCancel As Boolean 'true if download should be cancelled
Public blnInGame As Boolean 'true if GTA2 is not in the lobby
Public bln98 As Boolean 'true in the rare chance someone is still using Win95/98
Public blnGotCC As Boolean
Public strMMPfile As String 'This is the MMP file that the host was using when you tried to join
Public blnSync As Boolean
Public blnPlayReplay As Boolean
Public lngGTA2RunningTime As Long 'time GTA2 has been running for
Public strMacAddress As String

Public strFailedCountryNick As String
Public strFailedCountryIP As String
Public blnCountryDetectFail As Boolean

'Color variables, most are set in frmGH load
Public strCTCPcolor As String
Public strConnectionColor As String
Public strHelpColor As String
Public strActionColor As String
Public strTopicColor As String
Public strBannedColor As String
Public strGHColor As String
Public strPrivateMessageColor As String
Public strQuitColor As String
Public strJoinColor As String
Public strBackColor As String
Public strForeColor As String
Public strTextColor As String
Public strServerColor As String
Public strLinkColor As String
Public blnLinkUL As Boolean

'Font settings
Public strFontName As String
Public strFontSize As Integer
Public blnFontBold As Boolean
Public blnFontItalic As Boolean
Public blnUnderline As Boolean

Public blnHidden As Boolean 'true if hide registry key exists
Public strMacAddresses As String
Public intNickservWaitTime As Integer '-1 if nickserv says a name is registered and protected
Public blnConnected As Boolean  'this var will be used to check if we timed out, and will be set to true if get connected
Public blnPrivmsg As Boolean 'true if we received a private message
Public strDestTab As String
Public strExecutableChecksum As String 'Your GTA2.exe checksum
Public strMapChecksum As String 'Your currently select GMP checksum
Public strScriptChecksum As String 'Your currently select SCR checksum
Public intTimeSinceLastServerData As Integer
Public strData As String 'the var that will hold the data of a single IRC command
Public strChannel As String
Public strNick As String     'our global nickname var
'Public blnDONOTCHANGESERVER As Boolean
Public blnDisconnectClick As Boolean 'true if disconnect is clicked
Public strStatusMsg As String 'stores your status message
Public strAwayMsg As String 'stores your away message
Public strNickLastJoined As String 'nick of last player you tried to join
'Public blnFocus As Boolean 'stores whether app has focus
Public intLinePosition As Integer
Public strIPAddress As String
Public strExternalHostName As String
Public strHostNick As String 'the nick of the last host you tried to join
Public blnHosted As Boolean
Public blnSystray As Boolean
Public strGTA2version As String
Public strKey As String 'channel key
Public blnLogin As Boolean 'if login complete message has been displayed set to true
Public blnRouter As Boolean

'---------- frmOptions variables-------
Public intCountryIndex As Integer
Public strLocation1 As String
Public strLocation2 As String
Public blnchkMuteAlertSound As Boolean
Public blnchkAutoDownload As Boolean
Public blnchkVPN As Boolean
Public blnchkHide As Boolean
Public strTxtWordAlert As String
Public strAlertWords(20) As String
Public blnchkFlash1 As Boolean
Public blnchkFlash2 As Boolean
Public blnchkFlash3 As Boolean
Public blnchkSoundLocation1 As Boolean
Public blnchkSoundLocation2 As Boolean
Public blnchkSoundWordAlert As Boolean
Public blnchkSoundPrivmsg As Boolean
Public blnchkSoundJoin As Boolean
Public blnchkSoundHosted As Boolean
Public blnchkTime As Boolean
Public blnchkWine As Boolean
Public blnchkConnectOnStartup As Boolean
Public blnchkStartup As Boolean
Public blnchkTray As Boolean
Public blnchkCloseTray As Boolean
Public blnchkStartTray As Boolean
Public blnchkMinToTray As Boolean

'------------frmCreateGame variables------------
Public blnchkPad As Boolean
Public blnHighlight As Boolean 'Highlight alert words
Public blnchkGameClear As Boolean
Public blnReadyForJoiners As Boolean 'true when GTA2 is ready for joiners
Public blnLobby As Boolean
Public strPreviousMapDesc As String
Public strServer(5) As String 'stores the IRC server
Public strPort As String 'stores the IRC port
Public intServerNum As Integer 'stores the IRC server number
Public strPassword As String 'stores your IRC password
Public strErrdesc As String 'stores error desciption
Public strErrNum As String   'stores error number
Public strErrLine As String 'stores error line number
Public strPasswordProtectGame As String 'Yes if hosting with password, No if not
Public strYourGamePassword As String 'the password used to protect your hosted games from unwanted joiners
Public strCountryCode As String
Public strSelectedCountryCode As String
Public strCountries As Variant
'Public strEurope As Variant
Public strWave As String
Public strSoundLocation1 As String
Public strSoundLocation2 As String
Public strSoundWordAlert As String
Public strSoundPrivmsg As String
Public strSoundJoin As String
Public strSoundHosted As String
Public intColumnHeader4Width As Integer
Public strPreferedNick As String
Public strGTA2path As String
Public intPlayerCount  As Integer
Public strGTA2MapDesc As String 'stores the GTA2 map description
Public strGTA2MMP As String 'GTA2 .MMP (MapMultiPlayer filename)
Public strComment As String

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Function StrZToStr(s As String) As String
   StrZToStr = Left$(s, Len(s) - 1)
End Function

Function encode(strText As String) As String

Dim i As Integer            'Loop counter
Dim intKeyChar As Integer   'Character within the key that we'll use to encrypt
Dim strTemp As String       'Store the encrypted string as it grows
Const strKey = "G"          'The encryption key
Dim strChar1 As String * 1  'The first character to XOR
Dim strChar2 As String * 1  'The second character to XOR

    'Loop through each character in the text
    For i = 1 To Len(strText)
        'Get the next character from the text
        strChar1 = Mid(strText, i, 1)
        'Find the current "frame" within the key
        intKeyChar = ((i - 1) Mod Len(strKey)) + 1
        'Get the next character from the key
        strChar2 = Mid(strKey, intKeyChar, 1)
        'Convert the charaters to ASCII, XOR them, and convert to a character again
        strTemp = strTemp & Chr(Asc(strChar1) Xor Asc(strChar2))
    Next i
    
    'Return the resultant encoded/decoded string
    encode = strTemp
End Function

Public Sub ErrorHandler(strFunction As String, strErrdesc As String, strLine As String)
    If Val(strLine) > 0 Then strLine = "Line: " & strLine
    displaychat strDestTab, strTextColor, "Error during " & strFunction & " " & strErrdesc & " " & strLine
    send "PRIVMSG " & gta2ghbot & " :Error during " & strFunction & "  " & strErrdesc & " " & strLine
End Sub

