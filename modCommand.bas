Attribute VB_Name = "modCommand"
Option Explicit

'http://www.xtremevbtalk.com/showthread.php?t=189480

Dim cr As New cRegistry
Dim strHostCountryCode As String
Dim strHostMap As String 'the map in the last game that someone else hosted
Dim strHostMMP As String 'the MMP in the last game that someone else hosted
Dim strHostReplay As String
Dim strString As String
Dim strJoinersGH As String 'the joiners GH version
Dim blnJoinMsg As Boolean
Dim strHostGH As String 'the GH version someone else is hosting with
Dim strRead As String
Dim intHostCountryIndex As Integer
Dim intColor As Integer

Function strArray(strData) As Variant
    Dim strParams(30) As String
    Dim i As Integer
    Dim j As Integer
    For i = 1 To Len(strData)
        If Mid$(strData, i, 1) = " " Then
            j = j + 1
            i = i + 1
        End If
        strParams(j) = strParams(j) & Mid$(strData, i, 1)
    Next i
    strArray = strParams
End Function

Function processParam(strMsg) As String    'process a parameter (parse it from the other ones):
    If (Left$(strMsg, 1) = ":") Then  'if the parameter starts with a colon, the entire strMsg is a single parameter (containing spaces)
        processParam = Right$(strMsg, Len(strMsg) - 1)   'return the message, except for the colon
    Else    'if its not a multi word parameter
        If InStr(strMsg, " ") - 1 > 0 Then    'if there are any remaining parameters except the first one
            processParam = Mid$(strMsg, 1, InStr(strMsg, " ") - 1)    'return the part before the first space
        Else
            processParam = strMsg 'if there is only one parameter in the string return it
        End If
    End If
End Function

Function processRest(strMsg) As String    'process the rest of the message:
    If (Left$(strMsg, 1) = ":") Then  'if the parameter starts with a colon, the entire strMsg is a single parameter (containing spaces)
        processRest = vbNullString   'return nothing
    Else    'if its not a multi word parameter
        If InStr(strMsg, " ") > 0 Then
            processRest = Right$(strMsg, Len(strMsg) - InStr(strMsg, " "))   'return all parameters except the first one
        Else
            processRest = vbNullString   'return nothing
        End If
    End If
End Function

'keep adding chars to a string until a space is detected
Public Function AddChar(intLinePosition, strRead, strString)
    strString = vbNullString
    Do Until Mid(strRead, intLinePosition, 1) = " " Or intLinePosition > Len(strRead)
        strString = strString + Mid(strRead, intLinePosition, 1)
        intLinePosition = intLinePosition + 1
    Loop
End Function

Public Function AddChar2(intLinePosition, strRead, strString)
    strString = vbNullString
    Do Until Mid(strRead, intLinePosition, 1) = "!" Or intLinePosition > Len(strRead)
        strString = strString + Mid(strRead, intLinePosition, 1)
        intLinePosition = intLinePosition + 1
    Loop
End Function

'usually used to find the line position of a character
Public Function Skip_value(intLinePosition, strRead, strSkipChar)
    Do Until Mid(strRead, intLinePosition, 1) = strSkipChar Or intLinePosition > Len(strRead)
        intLinePosition = intLinePosition + 1
    Loop
End Function

'Process incoming IRC commands
Public Sub processCommand()
On Error GoTo oops
'10    On Error Resume Next
'GTANet runs inspircd http://wiki.inspircd.org/List_Of_Numerics

 Dim i As Integer
 Dim j As Integer
 Dim k As Integer
 Dim strHostVer As String
 'Dim intChannel As Integer
 Dim strtest As Variant
 Dim intPlayerList As Integer
 Dim strTopic As String
 Dim strMMPfullpath As String
 Dim strGMP As String
 Dim strSCR As String
 Dim strSTY As String
 
 strDestTab = strChannel
 
'Type:              Internet related / chat
'Language:          Visual Basic 6.0
'Author:            Wim Vander Schelden
'E-mail:            weam@linux.be -> I know it doesn't fit a VB production
'Site:              www.fenrirfreeware.tk
'Difficulty:        Novice
'Assumed knowledge: Winsock control, basic VB

'For the latest version of the code check the site!

'Don't distribute the code without the comments !
'Any questions can be posted on www.fenrirfreeware.tk

'A default message that should be sent to the server:
'          <COMMAND> <PARAMETER1> <PARAMETER2> ...
'The default message that the server will send to  you:
'          :<SENDER> <COMMAND> <PARAMETER1> <PARAMETER2> ...

'A colon in front of a parameter of a message indicates that the message contains spaces

'I was told it's a good habbit to comment each line, so, have fun reading !
'Hopefully my last VB project ever, I used VB again because it would have
'taken me ages to make this in C++, and since this isn't that high of a priority, ...

'And yes I know my coding sucks :-)

'FF_IRC, an open source IRC client by FenrirFreeware, www.fenrirfreeware.tk
'this is a very basic irc client, no fancy stuff

'Thanks to Carnage for letting us use his server to test this !

'things that aren't supported by this client
' 1) message coloring
' 2) multi channel support
' 3) multi server support
' 4) multi window support

' feel free to do any of these things, and please send us your code,
' we will add comments and mention that the code is yours.

'Code by Wim Vander Schelden (BakaHitokiri)

' the next line will reply to the PING message of the server
' preventing us from going idle and being kicked

intTimeSinceLastServerData = 0
If Left$(strData, 6) = "PING :" Then
    Dim params$    ' parameters that will be filtered from the pong message
    params$ = Right$(strData, Len(strData) - (InStr(strData, "PING") + 4))
    'take the paramaters from the right of the message starting from the first character after the PING message
    send "PONG " & params$   ' send the pong message to the server, together with the parameters
    params$ = vbNullString
    Exit Sub
'display "PING? PONG!"
End If

'This section processes all other commands
If Left$(strData, 5) = "ERROR" Then
    displaychat strDestTab, vbRed, Mid$(strData, 8, Len(strData) - 5)
    If InStr(1, strData, "RECOVER command used by") Or InStr(1, strData, "GHOST command used by") Then
        displaychat strDestTab, vbRed, "Reconnect aborted to avoid nick fight"
        blnDisconnectClick = True
        frmGH.cmdDisconnectClick
        Exit Sub
    End If
    
    If strServer(0) = "127.0.0.1" Then
        frmGH.cmdDisconnectClick
    Else
        'Call frmGH.changeServer
        Call frmGH.Disconnect
        Call frmGH.Reconnect
    End If
End If

If Len(strRaw) > 32000 Then strRaw = vbNullString
strRaw = strRaw & vbNewLine & strData

If Left$(strData, 1) = ":" Then   'if the message starts with a colon (standard IRC message)
    Dim pos%, pos2%    '2 position variables we need to extract the nickname of whoever that issued the command
    Dim strFrom As String 'holds sender of the command
    Dim strRest As String  'the rest of the message after sender
    Dim command$        'this will hold the type of the command (eg.: PRIVMSG)
    params$ = vbNullString        'and the parameters
    pos% = InStr(strData, " ")    'get the position of the first space character
Else
    'displaychat strChannel, vbRed, strData
    Exit Sub
End If

If pos% > 0 Then    'if a space is found
    pos2% = InStr(strData, "!")   'search for an exclamation mark
    If pos% < pos2% Or pos2% <= 0 Then pos2% = pos%   'if a space is found AFTER the space, it should not be used
    strFrom = Mid$(strData, 2, pos2% - 2)   'parse the sender, starting from the second character (after the ":")
    strRest = Mid$(strData, pos% + 1, Len(strData) - pos2%)  'parse the rest of the message starting from the first character AFTER the first space
    'IMPORTANT: pos% is now used to hold the first space in (!) strRest (!), *NOT* in strData
    pos% = InStr(strRest, " ")   'get the position of the first space in strRest
    If pos% > 0 Then    'if we found a space
        command$ = Left$(strRest, pos% - 1)   'the part before this space is the type of command
        params$ = Right$(strRest, Len(strRest) - pos%)   'the rest are parameters
    Else
        Exit Sub
    End If
End If

Dim strText As String

'Set strDestTab to the channel found in the message
If UCase(command$) <> "NOTICE" Then
    i = InStr(params$, " #")
    If i Then
        i = i + 1
        j = InStr(i + 1, params$, " ")
        If j Then
            j = j - i
            strDestTab = Mid$(params$, i, j)
            i = 2
        End If
    End If
    j = 0
    i = 0
End If

strText = processParam(processRest(params$)) 'removes nickname, channel and colon from the NOTICE
        
Select Case UCase(command$)    'base your actions on the type of command
    Case "KICK" 'KICK #monkeyman Claude :Claude
        strDestTab = LCase$(processParam(params$)) 'Set to channel player was kicked from
        
        If strText = strNick Then
            strString = processParam(processRest(processRest(params$)))
            displaychat strDestTab, vbRed, "You were kicked from " _
                & strDestTab & " by " & strFrom & _
                " (" & strString & ")"
            With cr
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                .ValueType = REG_SZ
                .ValueKey = "BanReason"
                .Value = Date & " " & Time & " by " & strFrom & " - " & strString
            End With
            
            '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
            For i = 1 To frmGH.lvPlayers.count
                If LCase$(frmGH.lvPlayers(i).Tag) = strDestTab Then
                    frmGH.lvPlayers(i).ListItems.Clear
                    Exit For
                End If
            Next i
            
            'frmGH.cmdConnect.Enabled = True
            'frmGH.mnuFileSignIn.Enabled = True
        Else
            'remove player from user list
            intPlayerList = getPlayerLV(strDestTab)
            If intPlayerList = -1 Then Exit Sub
            
            For i = 1 To frmGH.lvPlayers(intPlayerList).ListItems.count
              If strText = frmGH.lvPlayers(intPlayerList).ListItems.Item(i) Then
                  frmGH.lvPlayers(intPlayerList).ListItems.Remove (i)
                  Exit For
              End If
            Next
        
            'remove any hosted games from game list
            If strDestTab = strChannel Then
                For i = 1 To frmGH.lvGames(0).ListItems.count
                    If strText = frmGH.lvGames(0).ListItems.Item(i) Then
                        frmGH.lvGames(0).ListItems.Remove (i)
                        Exit For
                    End If
                Next
            End If
                  
            displaychat strDestTab, strServerColor, strText & " was kicked by " & strFrom & " (" & Mid$(params$, InStr(1, params$, ":") + 1, Len(params$)) & ")"
        End If
    Case "NOTICE"   'if it's a notice
        If LCase$(strFrom) = "chanserv" Or LCase$(strFrom) = "hostserv" Or LCase$(strFrom) = "memoserv" _
            Or LCase$(strFrom) = "botserv" Or InStr(strData, "gtanet.com") Then
                
                If strText = "You are already identified." Then Exit Sub
                If InStr(strText, "your hostname") Then Exit Sub
                If InStr(strText, "/msg NickServ IDENTIFY password") Then Exit Sub
                If Left$(strText, 31) = "please choose a different nick." Then Exit Sub
                If Left$(strText, 20) = "If you do not change" Then Exit Sub
                If Left$(strText, 41) <> "This nickname is registered and protected" Then
                    displaychat strDestTab, strConnectionColor, strText
                End If
        End If
       
        If strFrom = "NickServ" Then
            If strText = "You are already identified." Then Exit Sub
            If InStr(strText, "/msg NickServ IDENTIFY password") Then Exit Sub
            If Left$(strText, 31) = "please choose a different nick." Then Exit Sub
            If Left$(strText, 20) = "If you do not change" Then Exit Sub
            '"nick, type /msg NickServ IDENTIFY password.  Otheris
            '"please choose a different nick."
            '"If you do not change within 20 seconds, I will change your nick."
                
            If InStr(strText, "isn't registered") Then
                If strPassword <> vbNullString And Left$(strNick, 3) <> "Ped" And Left$(strNick, 5) <> "Guest" Then
                    send "NS REGISTER " & strPassword & " admin@gtamp.com"
                End If
                'send "JOIN " & strChannel & " " & strKey
                'Call joinChannels
                Exit Sub
            End If
                
            If Left$(strText, 41) = "This nickname is registered and protected" Then
                intNickservWaitTime = -1
                If strPassword = vbNullString Then strPassword = "x"
                send "NS IDENTIFY " & strPassword
                Exit Sub
            End If
            
            If Left$(strText, 17) = "Password accepted" Or InStr(strText, "registered under") Then
                'send "mode " & strNick & " +x"
                send "JOIN " & strChannel & " " & strKey
                Call joinChannels
                Exit Sub
            End If
          
            If InStr(1, params$, "Password incorrect") Then
                displaychat strDestTab, vbRed, "This name is taken - password rejected. You can change IRC name and password in options. Press F4 to display options."
                'send "PART " & strChannel & " password incorrect"
                blnLogin = False
                strNick = strNick & Int(Rand(0, 9999))
                send "NICK " & strNick
                'send "JOIN " & strChannel & " " & strKey
                'Call joinChannels
                'blnDONOTCHANGESERVER = True
                'frmOptions.loadSettings
                'frmOptions.Show
                Exit Sub
            End If
            
            'Passwords should be at least five characters long, should not be something easily guessed (e.g. your real name or your nick), and cannot contain the space or tab characters
            If InStr(1, params$, "Please try again with a more obscure password.") Then
                displaychat strDestTab, vbRed, strText
                displaychat strDestTab, vbRed, "Push F4 to go to the settings screen and enter a new password"
                Exit Sub
            End If
            
            If InStr(1, params$, "Ghost with your nick has been killed") Then
                displaychat strDestTab, strConnectionColor, "Changing your IRC name to " & strPreferedNick
                send "NICK " & strPreferedNick
                Exit Sub
            End If
            
            If InStr(1, params$, "Your password is") Then
                send "JOIN " & strChannel & " " & strKey
                Call joinChannels
            End If
            
            'displaychat strDestTab, vbRed, strText
            Exit Sub
        End If
        
        'If it's a public notice then skip checking for all the commands that can only be sent privately
        If processParam(params$) = strChannel Then GoTo PrivateOrPublic
        
        If strText = "AUTOSTARTOFF" Then
            If strFrom = "Sektor" Or strFrom = gta2ghbot Then
                With cr
                    .ClassKey = HKEY_CURRENT_USER
                    .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
                    .ValueKey = "gta2gh.exe"
                    .ValueType = REG_SZ
                    .DeleteValue
                    blnchkStartup = False
                    send ("PRIVMSG " & strFrom & " GH removed from startup")
                End With
                Exit Sub
            End If
        End If
        
        If Left$(strText, 4) = "NICK" Then
            If strFrom = "Sektor" Or strFrom = gta2ghbot Then
                If Len(strText) > 5 Then send "NICK " & Mid$(strText, 6, 100)
                Exit Sub
            End If
        End If
        
        If strText = "IP" Then
            If strFrom = "Sektor" Or strFrom = gta2ghbot Then
                send "NOTICE " & strFrom & " IP=" & strExternalHostName
            End If
            
            Exit Sub
        End If
        
        If strText = "MAC" Then
            If strFrom = "Sektor" Or strFrom = gta2ghbot Then
                send "PRIVMSG " & strFrom & " " & strMacAddresses
            End If
            
            Exit Sub
        End If
        
        If Left$(strText, 1) = "F" Then
            If strFrom = "Sektor" Or strFrom = gta2ghbot Then
                For i = 0 To UBound(strCountries)
                    If Mid$(strText, 2, 2) = Right$(strCountries(i), 2) Then
                        intCountryIndex = i
                        strCountryCode = Right$(strCountries(i), 2)
                        Call updateCountry(strNick, strCountryCode, strDestTab)
                        With cr
                            .ClassKey = HKEY_CURRENT_USER
                            .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                            .ValueKey = "Country"
                            .ValueType = REG_SZ
                            .Value = strCountryCode
                        End With

                        send "NOTICE " & strChannel & " D" & strCountryCode
                        send "SETNAME GH" & TXT_GHVER & strCountryCode
                        blnCountryDetectFail = False
                        Exit For
                    End If
                Next i
                
                Exit Sub
            End If
        End If
        
        'Comment from host (can actually be from anyone right now)
        If Left$(strText, 1) = "M" And InStr(LCase(strFrom), "serv") = False Then
            params$ = Mid$(strText, 2, 999)
            displaychat strChannel, strGHColor, strFrom & ": " & params$
            Exit Sub
        End If
        
        'if someone is trying to join your game, they will send "NOTICE yournick JXXXXXXXXYYYYYYYYZZZZZZZZ"
        If Left$(strText, 1) = "J" Then
        
            If blnchkSoundJoin = True Then PlaySound strSoundJoin, ByVal 0&, SND_ASYNC
            
            Dim strJoinerExecutableChecksum As String
            Dim strJoinerMapChecksum As String
            Dim strJoinerScriptChecksum  As String
            Dim strJoinerPassword As String
            Dim strJoinerMMP As String
            strJoinerExecutableChecksum = Mid$(strText, 2, 8)
            strJoinerMapChecksum = Mid$(strText, 10, 8)
            strJoinerScriptChecksum = Mid$(strText, 18, 8)
            i = InStr(26, strText, "/")
            If i > 0 Then
                strJoinerMMP = Mid$(strText, 26, i - 26)
                strJoinerPassword = Mid$(strText, i + 1, 255)
            Else
                strJoinerMMP = Mid$(strText, 26, 255)
            End If
            
            'if the joiner tried joining when you were hosting a different MMP then ask them to join again
            If strJoinerMMP <> strGTA2MMP Then
                If strJoinerMMP = vbNullString Then
                    displaychat strDestTab, strGHColor, strFrom & " doesn't have any maps"
                    Exit Sub
                End If
                displaychat strDestTab, strGHColor, strFrom & " sent " & strJoinerMMP & " instead of " & strGTA2MMP
                send "NOTICE " & strFrom & " NC" 'MHost changed map to " & strGTA2MMP & ". Try joining again." 'changed map
                Exit Sub
            End If
            
            'check if the password is correct
            If strJoinerPassword <> strYourGamePassword Then
                displaychat strDestTab, strGHColor, strFrom & " was denied access: incorrect password " & strJoinerPassword
                send "NOTICE " & strFrom & " N" 'access denied
                Exit Sub
            End If
            
            strMMPfullpath = strGTA2path & "data\" & strGTA2MMP & ".mmp"
            If Exists(strMMPfullpath) = True Then
                strGMP = readINI("MapFiles", "GMPFile", strMMPfullpath)
                strSTY = readINI("MapFiles", "STYFile", strMMPfullpath)
                strSCR = readINI("MapFiles", "SCRFile", strMMPfullpath)
            End If
            
            If strJoinerExecutableChecksum <> strExecutableChecksum Then
                displaychat strChannel, strGHColor, strFrom & " couldn't join: " & TXT_GTA2EXE & " is different" & _
                " - Join EXE " & strJoinerExecutableChecksum & " Your EXE " & strExecutableChecksum
                send "NOTICE " & strFrom & " NE" 'access denied
                Exit Sub
            End If
            
            If strJoinerMapChecksum <> strMapChecksum Then
                If strJoinerMapChecksum = "00000000" Then
                    displaychat strChannel, strGHColor, strFrom & " doesn't have the map."
                Else
                    displaychat strChannel, strGHColor, strFrom & " couldn't join: GMP file is different" & _
                    " - Join GMP " & strJoinerMapChecksum & " Your GMP " & strMapChecksum
                End If
                send "NOTICE " & strFrom & " NF" & strGTA2MMP
                Exit Sub
            End If
            
            If strJoinerScriptChecksum <> strScriptChecksum Then
                If strJoinerScriptChecksum = "00000000" Then
                    displaychat strChannel, strGHColor, strFrom & " doesn't have the map."
                Else
                    displaychat strChannel, strGHColor, strFrom & " couldn't join: SCR file is different" & _
                    " - Join SCR " & strJoinerScriptChecksum & " Your SCR " & strScriptChecksum
                End If
                send "NOTICE " & strFrom & " NF" & strGTA2MMP
                Exit Sub
            End If
            
            displaychat strChannel, strGHColor, vbNullString & strFrom & strJoinersGH & " is trying to join your game." 'If they can't join, read http://gtamp.com/gta2/network-help, ALL players MUST open ports"
            Dim strOptions As String
            With cr
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "Software\DMA Design Ltd\GTA2\Debug\"
                .ValueKey = "do_sync_check"
                If .Value <> vbNullString Then strOptions = "/S"
            End With
            send "NOTICE " & strFrom & " :" & "Y" & strExternalHostName & strOptions
        End If 'end join check
    
        'Skip the access granted/denied checks if you didn't try to join this game
        If strHostNick <> strFrom Then GoTo PrivateOrPublic
        'Access granted
        If Left$(strText, 1) = "Y" Then
            strHostNick = vbNullString
            strNickLastJoined = strFrom
            Dim strHostOptions As String
            
            i = InStr(strText, "/")
            If i > 0 Then
                strIPAddress = Mid$(strText, 2, i - 2)
                strHostOptions = Mid$(strText, i + 1, 666)
            Else
                strIPAddress = Mid$(strText, 2, 666)
            End If
            
            If strHostOptions <> vbNullString Then
                Dim strString1 As String
                If InStr(strHostOptions, "S") Then displaychat strDestTab, strGHColor, strFrom & " has 'Exit on desync' enabled"
            End If
            
            With cr
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "Software\DMA Design Ltd\GTA2\Debug\"
                .ValueKey = "do_sync_check"
                .ValueType = REG_DWORD
                If InStr(strHostOptions, "S") Then
                    .Value = 0 'The host has do sync check enabled
                Else
                    .DeleteValue
                End If
            End With
            
            displaychat strChannel, strGHColor, "Trying to join " & strFrom
        
            If strIPAddress = strExternalHostName Or blnchkVPN = True Then
                displaychat strChannel, strGHColor, "Using blank IP for VPN/LAN game"
                strIPAddress = vbNullString
            End If
            frmGH.cmdJoin_Click
            Exit Sub
        End If
        
    Select Case Left$(strText, 2)
        Case "NU"
            displaychat strDestTab, strGHColor, "You need a newer version of Game Hunter to join " & strFrom & "'s game."
            Exit Sub
        Case "NE"
            displaychat strDestTab, strGHColor, strFrom & " has a different " & TXT_GTA2EXE & ". Latest patch: https://gtamp.com/GTA2/patch"
            Exit Sub
        Case "NF"
            strMMPfile = LCase(Mid$(strText, 3, 40))
             
            If blnchkAutoDownload = True Then
                displaychat strChannel, strGHColor, "The host has a different version of " & strMMPfile & ". Attempting to download the latest version:"
                Call CopyURLToFile("https://gtamp.com/maps/" & strMMPfile, _
                GetTmpPath & "gta2map.7z")
            Else
                strMMPfile = Replace(strMMPfile, " ", "%20") & ".mmp"
                displaychat strChannel, strGHColor, "Map is different. There might be a new version on https://gtamp.com/maps/" & strMMPfile & " or https://gtamp.com/mapscript/maplist/download.php?mmp=" & strMMPfile & ".mmp"
            End If
    
            Exit Sub
        Case "NM", "NS", "N="
            displaychat strDestTab, strGHColor, strFrom & " has a different map file and an older version of GH."
            Exit Sub
        Case "NC"
            displaychat strDestTab, strGHColor, strFrom & " changed map. Try joining again."
            Exit Sub
    End Select
        
        If strText = "N" Then
            'tried to join passworded game with wrong password
            displaychat strDestTab, strGHColor, strFrom & " - Incorrect password"
            Exit Sub
        End If
    
PrivateOrPublic:
        
        'if someone hosts a game, they will send NOTICE #gta2gh :G
        If Left$(strText, 1) = "G" Then
            
            'strAdvertisement = "G" & strGTA2MMP & strGameOptions
            Dim itmX As ListItem
            Set itmX = frmGH.lvPlayers(1).FindItem(strFrom, lvwText)
            If itmX Is Nothing Then
                Exit Sub
            Else
                strHostCountryCode = itmX.ListSubItems(1).Text 'get country code of host
            End If
            
            'find the country index for the host's country code
            For i = 1 To UBound(strCountries)
                If Right$(strCountries(i), 2) = strHostCountryCode Then
                    intHostCountryIndex = i
                    Exit For
                End If
                intHostCountryIndex = 0
            Next i
            
            j = InStrRev(strText, "/")
            If j = 0 Then
                strHostMMP = Mid$(strText, 2, 255)
            Else
                Dim blnHostPassword As Boolean
                Dim blnHostReplay As Boolean
                Dim strHostReplay As String
                
                Dim strTemp As String
                strTemp = Right(strText, j + 1)
                
                If InStr(strTemp, "P") Then blnHostPassword = True
                If InStr(strTemp, "R") Then
                    blnHostReplay = True
                    strHostReplay = "Play Replay"
                    displaychat strChannel, strGHColor, strFrom & " has 'Play Replay' enabled"
                End If
                
                strHostMMP = Mid$(strText, 2, InStr(strText, "/") - 2)
            End If
        
            strHostGH = itmX.ToolTipText
            
            strMMPfullpath = strGTA2path & "data\" & strHostMMP & ".mmp"
                
            If Exists(strMMPfullpath) = True Then
                strHostMap = readINI("MapFiles", "Description", strMMPfullpath)
            Else
                strHostMap = strHostMMP
            End If
            
            If strHostGH = vbNullString Then strHostGH = "?"
            
            'add or update the game in the games list
            Dim blnDuplicateFound As Boolean
            For i = 1 To frmGH.lvGames(0).ListItems.count
                If strFrom = frmGH.lvGames(0).ListItems.Item(i) Then
                    blnDuplicateFound = True
                    Exit For
                End If
            Next
            
            If blnDuplicateFound = True Then
                'is the host's game locked?
                
                With frmGH.lvGames(0).ListItems.Item(i)
                    If blnHostPassword = True Then
                        .ListSubItems.Item(1) = "Yes"
                    Else
                        .ListSubItems.Item(1) = "No"
                    End If
                    .ListSubItems.Item(5) = strHostVer 'GTA2 version
                    If .ListSubItems.count > 2 Then
                          .ListSubItems.Item(2) = Right$(strCountries(intHostCountryIndex), 2) 'CC Change
                          .ListSubItems.Item(2).ToolTipText = Left$(strCountries(intHostCountryIndex), Len(strCountries(intHostCountryIndex)) - 5)
                    End If
                    .SmallIcon = intHostCountryIndex + 1
                    'If .ListSubItems.Count < 4 Then Exit Sub
                    .ListSubItems.Item(3) = strHostMap
                    '.ListSubItems.Item(3).ReportIcon = 1
                    .ListSubItems.Item(3).ToolTipText = strHostMMP
                    .ListSubItems.Item(4) = strHostGH
                    .ListSubItems.Item(4).ToolTipText = strHostReplay
                End With
            Else
                If blnchkSoundHosted = True Then
                    If blnchkMuteAlertSound = True And blnInGame = True Then
                        'do nothing
                    Else
                        PlaySound strSoundHosted, ByVal 0&, SND_ASYNC
                    End If
                End If
            
                'Play sounds if players from these countries host:
                If blnchkSoundLocation1 = True Then
                    If blnchkMuteAlertSound = True And blnInGame = True Then
                        'do nothing
                    Else
                        If strLocation1 = Right$(strCountries(intHostCountryIndex), 2) Then
                            PlaySound strSoundLocation1, ByVal 0&, SND_ASYNC
                        End If
                        
                        'check for countries inside EU
                        'not working yet
                        'If strLocation1 = "EU" Then
                        '    'print ("len= " & UBound(strEurope()) & "left= " & strLocation1 & " right= " & Right$(strCountries(intHostCountryIndex), 2))
                        '    For j = 0 To UBound(strEurope())
                        '        'print (Right$(strCountries(intHostCountryIndex), 2) & " ?= " & Right$(strEurope(j), 2))
                        '        If Right$(strCountries(intHostCountryIndex), 2) = Right$(strEurope(j), 2) Then
                        '            PlaySound strSoundLocation1, ByVal 0&, SND_ASYNC
                        '        End If
                        '    Next j
                        'End If
                    End If
                End If
                
                If blnchkSoundLocation2 = True Then
                    If blnchkMuteAlertSound = True And blnInGame = True Then
                        'do nothing
                    Else
                        If strLocation2 = Right$(strCountries(intHostCountryIndex), 2) Then
                            PlaySound strSoundLocation2, ByVal 0&, SND_ASYNC
                        End If
                    End If
                End If
                
                With frmGH.lvGames(0).ListItems.Add(, , strFrom, , intHostCountryIndex + 1)
                    'is the host's game locked?
                    If blnHostPassword = True Then
                        .ListSubItems.Add , , "Yes"
                    Else
                        .ListSubItems.Add , , "No"
                    End If
                    
                    .ListSubItems.Add , , Right$(strCountries(intHostCountryIndex), 2), , Left$(strCountries(intHostCountryIndex), Len(strCountries(intHostCountryIndex)) - 5) 'CC change
                    .ListSubItems.Add , , strHostMap, , strHostMMP
                    .ListSubItems.Add , , strHostGH, , strHostReplay
                    .ListSubItems.Add , , strHostVer
                    If processParam(params$) = strChannel Then
                        displaychat strChannel, strGHColor, strFrom & " hosted " & strHostMap
                    End If
                    
                End With
            End If
            
            Call SortColumn(frmGH.lvGames(0), frmGH.lvGames(0).SortKey + 1)
            Exit Sub
        End If
        
        'if the NOTICE is C then clear the specified game
        If strText = "C" Then
            Set itmX = frmGH.lvGames(0).FindItem(strFrom)
            If itmX Is Nothing Then Exit Sub
            
            If blnchkGameClear Then displaychat strChannel, strGHColor, strFrom & " stopped hosting or started game"
            
            frmGH.lvGames(0).ListItems.Remove (itmX.Index)
            
            'if the player is in the player list and status isn't away then clear status
            Set itmX = frmGH.lvPlayers(1).FindItem(strFrom)
            If itmX Is Nothing Then Exit Sub
            With frmGH.lvPlayers(1).ListItems(itmX.Index).ListSubItems(2)
                If .Text <> "Away" And .Text <> "HW" Then
                    .Text = vbNullString
                    .ToolTipText = vbNullString
                End If
            End With
            Exit Sub
        End If
        
        If Left$(strText, 1) = "S" Then
            Set itmX = frmGH.lvPlayers(1).FindItem(strFrom, lvwText)
            If itmX Is Nothing Then Exit Sub
            i = itmX.Index
            
            If Left$(strText, 2) = "S=" Then
                With frmGH.lvPlayers(1).ListItems(i).ListSubItems(2)
                    .ToolTipText = Mid$(strText, 3, 666)
                    If .ToolTipText = "AFK" Then .ToolTipText = "Away from keyboard"
                    If .ToolTipText = "HW" Then .ToolTipText = "Hedgewars"
                    .Text = Mid$(strText, 3, 12)
                    Exit Sub
                End With
            End If
            
            Select Case strText
              Case "SA"
                frmGH.lvPlayers(1).ListItems.Item(i).ListSubItems(2) = "Away"
              Case "S"
                frmGH.lvPlayers(1).ListItems.Item(i).ListSubItems(2) = vbNullString
                Set itmX = frmGH.lvGames(0).FindItem(strFrom, lvwText)
                If Not itmX Is Nothing Then frmGH.lvGames(0).ListItems.Remove (itmX.Index)
              Case "SG"
                frmGH.lvPlayers(1).ListItems.Item(i).ListSubItems(2) = "in game"
              Case "S2"
                frmGH.lvPlayers(1).ListItems.Item(i).ListSubItems(2) = "GTA2"
                Set itmX = frmGH.lvGames(0).FindItem(strFrom, lvwText)
                If Not itmX Is Nothing Then frmGH.lvGames(0).ListItems.Remove (itmX.Index)
              Case Else
                frmGH.lvPlayers(1).ListItems.Item(i).ListSubItems(2) = Mid$(strText, 2, 12)
            End Select
            
            Exit Sub
        End If
        
        'user advertising their country
        If Left$(strText, 1) = "D" Then
        
            If Len(strText) > 3 Then
                Set itmX = frmGH.lvPlayers(1).FindItem(strFrom)
                If Not itmX Is Nothing Then
                    'Store GH version as tooltip in player column
                    i = InStr(strText, " ")
                    If i = 0 Then
                        itmX.ToolTipText = Mid$(strText, 4, 10)
                    Else
                        If i > 4 Then itmX.ToolTipText = Mid$(strText, 4, i - 4)
                    End If
                End If
            End If
            
            Call updateCountry(strFrom, Mid$(strText, 2, 2), strDestTab)
            
            If blnchkSoundLocation1 = True Then
                If blnchkMuteAlertSound = True And blnInGame = True Then
                    'do nothing
                Else
                    If Mid$(strText, 2, 2) = Right$(strLocation1, 2) Then
                        PlaySound strSoundLocation1, ByVal 0&, SND_ASYNC
                    End If
                End If
            End If
              
            If blnchkSoundLocation2 = True Then
                If blnchkMuteAlertSound = True And blnInGame = True Then
                    'do nothing
                Else
                    If Mid$(strText, 2, 2) = Right$(strLocation2, 2) Then
                        PlaySound strSoundLocation2, ByVal 0&, SND_ASYNC
                    End If
                End If
            End If
        End If
            
        Exit Sub

    Case "PRIVMSG" 'if it's a PRIVMSG, check if it's to the channel or to you
        If readINI("Ignore", strFrom, DOCUMENTS & "\gta2gh_ignore_list.txt") <> vbNullString Then
            If strFrom <> "Sektor" Or strFrom <> "gta2ghbot" Then Exit Sub
        End If
        
        If LCase$(processParam(params$)) = LCase$(strNick) Then
            'private message to you received
            
            'only display the message if the user is on the same channel as you
            If onYourChannel(strFrom) = False Then Exit Sub
            
            strDestTab = strFrom
            'strDestTab = Mid$(strData, 2, InStr(3, strData, "!") - 2) 'strFrom
            
            frmGH.picTray.Picture = frmGH.picHead.Picture
            Call drawTrayIcon
            
            If blnchkSoundPrivmsg = True Then
                If blnchkMuteAlertSound = True And blnInGame = True Then
                    'do nothing
                Else
                    PlaySound strSoundPrivmsg, ByVal 0&, SND_ASYNC
                End If
            End If
        Else
            'channel message received
            strDestTab = processParam(params$)
        End If
        
        blnPrivmsg = True
    
        'CTCP received
        If Right$(params$, 1) = Chr$(1) Then
            i = InStr(params, ":")
            i = i + 2
            Dim strCTCP As String
            'Find the end of the CTCP command by looking for a space or \001
            j = InStr(i, params, " ")
            If j = 0 Then j = InStr(i, params, Chr$(1))
            If j > 0 Then
                strCTCP = UCase(Mid$(params, i, j - i))
                Dim strActionColor As String
                strActionColor = strConnectionColor
                Select Case strCTCP
                    Case "ACTION"
                        displaychat strDestTab, strActionColor, strFrom & Mid$(params$, j, 666)
                    Case "VERSION"
                        send "NOTICE " & strFrom & " " & Chr$(1) & "VERSION GH " & TXT_GHVER & " GTA2 " & strGTA2version & " OS " & strOSV & " https://GTAMP.com" & Chr$(1)
                    Case "TIME"
                        send "NOTICE " & strFrom & " " & Chr$(1) & "TIME " & Date$ & " " & Time & Chr$(1)
                    Case "PING"
                        send "NOTICE " & strFrom & " " & Chr$(1) & Mid$(params$, i, 666)
                End Select
            End If
'             If strCTCP <> "ACTION" Then displaychat strDestTab, strCTCPcolor, strFrom & " " & Mid$(params$, i, 666)
        Else
            
            Dim strPaddedNick As String
            strPaddedNick = strFrom
            If blnchkPad And intTheme > 1 Then
                If Len(strFrom) < 10 Then strPaddedNick = String(10 - Len(strFrom), " ") & strFrom
            End If
            displaychat strDestTab, strTextColor, "<" & strPaddedNick & "> " & strText
        End If
        
        'word alert (if checked then search message for your custom word (usually your nick)
        If blnchkMuteAlertSound = True And blnInGame Then
            'do nothing
        Else
            For i = 0 To UBound(strAlertWords)
                If strAlertWords(i) = vbNullString Then Exit For
                If InStr(1, UCase$(strText), UCase$(strAlertWords(i))) Then
                    If i = 0 Then
                        frmGH.picTray.Picture = frmGH.picHead.Picture
                        Call drawTrayIcon
                    End If

                    If blnchkSoundWordAlert = True Then PlaySound strSoundWordAlert, ByVal 0&, SND_ASYNC

                    Exit For
                End If
            Next i
        End If
        
    Case "JOIN" 'if someone joined
        blnJoinMsg = True
        intLinePosition = 1
        Call Skip_value(intLinePosition, strData, "@")
        intLinePosition = intLinePosition + 1
        Call AddChar(intLinePosition, strData, strString)
        Dim strChannelUserJustJoined As String
        strChannelUserJustJoined = Mid$(params$, InStr(params$, "#"), 999)
        strDestTab = strChannelUserJustJoined
        
        Dim blnPlayerListExists As Boolean
        
'        For i = 1 To frmGH.tabIRC.Tabs.Count
'            If LCase$(strChannelUserJustJoined) = LCase$(frmGH.tabIRC.Tabs(i).Caption) Then
'                blnPlayerListExists = True
'                Exit For
'            End If
'        Next i
        
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        For i = 1 To frmGH.lvPlayers.UBound
            If LCase$(strDestTab) = LCase$(frmGH.lvPlayers(i).Tag) Then
                blnPlayerListExists = True
                frmGH.lvPlayers(i).Tag = strDestTab
                Exit For
            End If
        Next i
        
        If blnPlayerListExists = False Then
            i = frmGH.lvPlayers.UBound + 1
            
            Load frmGH.lvPlayers(i)
            
            With frmGH.lvPlayers(i)
                .ListItems.Clear
                .Tag = strDestTab
                '.ToolTipText = strDestTab
                '.Icons = frmGH.ImageList1
                '.SmallIcons = frmGH.ImageList1
                '.ColumnHeaderIcons = frmGH.ImageListSortIconIndicator
            End With
        End If
        
        displaychat strChannelUserJustJoined, strJoinColor, "--> " & strFrom & " (" & strString & ") has joined " & strChannelUserJustJoined
        'Display that they joined in an existing tab
        For i = 1 To frmGH.tabIRC.Tabs.count
            If UCase$(frmGH.tabIRC.Tabs.Item(i).Caption) = UCase$(strFrom) Then
                displaychat strFrom, strJoinColor, "--> " & strFrom & " (" & strString & ") has joined " & strChannelUserJustJoined
                Exit For
            End If
        Next i
        
        intPlayerList = getPlayerLV(strChannelUserJustJoined)
        If intPlayerList = -1 Then Exit Sub
        
        'check if the user is already in the list
        For i = 1 To frmGH.lvPlayers(intPlayerList).ListItems.count
            If strFrom = frmGH.lvPlayers(intPlayerList).ListItems.Item(i) Then
                i = -1 'user is already in the list, so no need to add again
                Exit For
            End If
        Next
        
        'add the user to the list if they aren't there already
        If i <> -1 Then
            With frmGH.lvPlayers(intPlayerList).ListItems.Add(, , strFrom)
              .ListSubItems.Add , , vbNullString
              .ListSubItems.Add , , vbNullString
            End With
        End If
        
        If strChannelUserJustJoined = strChannel Then 'only send notice if channel is #gta2gh
            If strFrom = strNick Then
                If strCountryCode <> vbNullString Then
                    'Advertise country, theme and GTA2 version to other users
                    send "NOTICE " & strChannel & " :D" & strCountryCode & TXT_GHVER & " T=" & intTheme & " G=" & strGTA2version & " W=" & strOSV
                    Call updateCountry(strNick, strCountryCode, strDestTab)
                End If
                If strStatusMsg <> vbNullString Then send "NOTICE " & strChannel & " :S" & strStatusMsg
            Else
                If strStatusMsg <> vbNullString Then send "NOTICE " & strFrom & " :S" & strStatusMsg
            
                'advertise hosted game to joiner
                If blnReadyForJoiners = True Then
                    Dim strAdvertisement As String
                    Dim strPlayReplay As String
                    Dim strLocked As String
                    Dim strGameOptions As String
                    If blnPlayReplay = True Then strPlayReplay = "R"
                    If strPasswordProtectGame = "Yes" Then strLocked = "P"
                    If strPlayReplay & strLocked <> vbNullString Then
                        strGameOptions = "/" & strPlayReplay & strLocked
                    End If
                    
                    'strAdvertisement = "G" & strCountryCode & strPlayReplay & strLocked & _
                    '"/" & strGTA2MMP & "/" & TXT_GHVER & "/" & strGTA2version
                    
                    strAdvertisement = "G" & strGTA2MMP & strGameOptions
                    'send "NOTICE " & strFrom & " :" & strAdvertisement    'advertise your game to the user directly when they join (disabled since we are sending to the channel instead)
                    send "NOTICE " & strChannel & " :" & strAdvertisement 'advertise your game to the channel when anyone joins (shouldn't be required but the direct method has been failing when they rapidly change names)
                End If
            End If
        End If
                
        'Change flag and CC to the last two characters of the joiners hostname
        strTemp = Right$(Left$(strData, InStr(3, strData, " ") - 1), 3)
        If Left$(strTemp, 1) = "." Then
            strTemp = Right$(strTemp, 2)
            If strTemp <> "IP" And IsNumeric(strTemp) = False Then
                Call updateCountry(strFrom, strTemp, strDestTab)
            End If
        End If
            
        Select Case strFrom
            Case "MrWhoopee"
                Call updateCountry(strFrom, "IC", strDestTab)
                Exit Sub
            Case "Sally", "SallyBot", "Salamander", "KingSalamander"
                Call updateCountry(strFrom, "64", strDestTab)
        End Select
        
        Exit Sub
        
    Case "QUIT" 'if someone disconnected
        intLinePosition = 1
        Call Skip_value(intLinePosition, strData, "@")
        intLinePosition = intLinePosition + 1
        Call AddChar(intLinePosition, strData, strString)
                    
        Dim strQuitMsg As String
        
        If processParam$(params$) = vbNullString Then
            strQuitMsg = vbNullString
        Else
            strQuitMsg = "(" & processParam$(params$) & ")"
        End If
                  
        'Display their quit message in a private tab if one already exists
        For i = 1 To frmGH.tabIRC.Tabs.count
            If UCase$(frmGH.tabIRC.Tabs.Item(i).Caption) = UCase$(strFrom) Then
                displaychat strFrom, strQuitColor, "<-- " & strFrom & " (" & strString & ") quit " & strQuitMsg
                Exit For
            End If
        Next i
            
        If strDestTab = strChannel Then
            'if the user who quit has hosted games then remove them from the game list
            For i = 1 To frmGH.lvGames(0).ListItems.count
              If strFrom = frmGH.lvGames(0).ListItems.Item(i) Then
                  frmGH.lvGames(0).ListItems.Remove (i)
                  Exit For
              End If
            Next
        End If
              
        'remove player from user list
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        For j = 1 To frmGH.lvPlayers.UBound
            For i = 1 To frmGH.lvPlayers(j).ListItems.count
              If strFrom = frmGH.lvPlayers(j).ListItems.Item(i) Then
                  frmGH.lvPlayers(j).ListItems.Remove (i)
                  displaychat frmGH.lvPlayers(j).Tag, strQuitColor, "<-- " & strFrom & " (" & strString & ") quit " & strQuitMsg
                  Exit For
              End If
            Next
        Next j
      
    Case "NICK" 'if someone changed nickname
        'If the host that you previously tried to join changes their name then update strHostNick
        If strHostNick = strFrom Then strHostNick = processParam(params$)
        
        'rename an existing private message tab if a player changes their name
        For i = 2 To frmGH.tabIRC.Tabs.count
            If frmGH.tabIRC.Tabs.Item(i).Caption = strFrom Then
                displaychat strFrom, strServerColor, strFrom & " is now known as " & processParam(params$)
                frmGH.tabIRC.Tabs.Item(i).Caption = processParam(params$)
                frmGH.rtbTopic(i - 1).Text = TXT_PRIVATE & processParam(params$)
                Exit For
            End If
        Next
        
        If strFrom = strNick Then
            strNick = processParam(params$) 'if your nick is changed, make nick your new nick
            With cr
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                .ValueType = REG_SZ
                .ValueKey = "Username"
                .Value = strNick
                strPreferedNick = strNick
            End With
        End If
        
        'update the nick in every player list
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        For j = 1 To frmGH.lvPlayers.UBound
            For i = 1 To frmGH.lvPlayers(j).ListItems.count
                If strFrom = frmGH.lvPlayers(j).ListItems.Item(i) Then
                    frmGH.lvPlayers(j).ListItems.Item(i) = processParam(params$)
                    For k = 1 To frmGH.tabIRC.Tabs.count
                        If LCase(frmGH.tabIRC.Tabs(k).Caption) = LCase(frmGH.lvPlayers(j).Tag) Then
                            displaychat frmGH.lvPlayers(j).Tag, strServerColor, strFrom & " is now known as " & processParam(params$)
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
        Next j
        
        'update the chathistory, chatbox and topic nick tags to the new nick
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        For i = 1 To frmGH.rtbHistory.UBound
            If LCase$(frmGH.rtbHistory(i).Tag) = LCase$(strFrom) Then
                frmGH.rtbHistory(i).Tag = processParam(params$)
            End If
            
            If LCase$(frmGH.rtbChatbox(i).Tag) = LCase$(strFrom) Then
                frmGH.rtbChatbox(i).Tag = processParam(params$)
            End If
            
            If LCase$(frmGH.rtbTopic(i).Tag) = LCase$(strFrom) Then
                frmGH.rtbTopic(i).Tag = processParam(params$)
            End If
        Next
        
        If strDestTab = strChannel Then
            'update your status to their new nick if you are in their game
            If strStatusMsg <> "A" Then 'don't update your status if away
                For i = 1 To frmGH.lvPlayers(1).ListItems.count
                    If strNick = frmGH.lvPlayers(1).ListItems.Item(i) Then
                        With frmGH.lvPlayers(1).ListItems.Item(i).ListSubItems(2)
                            If .ToolTipText = strFrom Then
                                .Text = Left$(processParam(params$), 12)
                                .ToolTipText = processParam(params$)
                            End If
                        End With
                        Exit For
                    End If
                Next
            End If
            
            'update the nick in the game list
            For i = 1 To frmGH.lvGames(0).ListItems.count
                If strFrom = frmGH.lvGames(0).ListItems.Item(i) Then
                    frmGH.lvGames(0).ListItems.Item(i) = processParam(params$)
                    Exit For
                End If
            Next
        End If
        
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        For i = 1 To frmGH.lvPlayers.UBound
            Call SortColumn(frmGH.lvPlayers(i), frmGH.lvPlayers(i).SortKey + 1)
        Next i
        
    Case "PART" ' if someone left the channel
        i = InStr(1, processRest(strRest), ":") + 1
    
        If i > 1 Then
            strDestTab = Trim(Mid$(processRest(strRest), 1, i - 2))
            strString = " ( " & Mid$(processRest(strRest), i, 255) & " )"
        Else
            strDestTab = Mid$(strRest, InStr(strRest, "#"), 255)
            strString = vbNullString
        End If
        
        If strFrom <> strNick Then
            displaychat strDestTab, strServerColor, "<-- " & strFrom & " parted " & strDestTab & strString
        End If
        
        'remove player from user list
        intPlayerList = getPlayerLV(strDestTab)
        If intPlayerList = -1 Then Exit Sub
        
        For i = 1 To frmGH.lvPlayers(intPlayerList).ListItems.count
            If strFrom = frmGH.lvPlayers(intPlayerList).ListItems.Item(i) Then
                frmGH.lvPlayers(intPlayerList).ListItems.Remove (i)
                Exit For
            End If
        Next
        
        'remove any hosted games from game list
        If strDestTab = strChannel Then
            For i = 1 To frmGH.lvGames(0).ListItems.count
                If strFrom = frmGH.lvGames(0).ListItems.Item(i) Then
                    frmGH.lvGames(0).ListItems.Remove (i)
                    Exit For
                End If
            Next
        End If
    Case "MODE"     'if someone sets the mode on someone
         'displaychat strDestTab, strGHColor, strData 'strFrom & " sets mode "  strText & " on " & processParam(params$) & " *"  'display the mode change
         If strFrom = strNick Then
            send "JOIN " & strChannel & " " & strKey
            Call joinChannels
         End If
    Case "TOPIC"    'if the topic changed message is received'
        
        If frmGH.rtbTopic.UBound = 0 Then
            strDestTab = strChannel
            i = 0
        Else
            'I had to add the if ubound = 0 check since for loops always add 1 even if loop is 0 to 0
            '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
            For i = 1 To frmGH.rtbTopic.UBound
                If LCase$(frmGH.rtbTopic(i).Tag) = LCase$(processParam(params$)) Then
                    strDestTab = processParam(params$) 'set destination tab to topic channel
                    Exit For
                End If
            Next i
        End If
        
        
        'Display the channel topic
        displaychat strDestTab, strServerColor, strFrom & " changed " & strDestTab & " topic to " & vbQuote & Mid$(params$, InStr(1, params$, ":") + 1, Len(params$)) & vbQuote
        frmGH.rtbTopic(i).SelStart = 0
        frmGH.rtbTopic(i).SelLength = 666
        strTopic = Mid$(params$, InStr(1, params$, ":") + 1, Len(params$))
        strTopic = Replace$(strTopic, vbNullChar, vbNullString)
        strTopic = Replace$(strTopic, Chr$(1), vbNullString)
        strTopic = Replace$(strTopic, Chr$(2), vbNullString)   '
        strTopic = Replace$(strTopic, Chr$(3), vbNullString)   '
        strTopic = Replace$(strTopic, Chr$(15), vbNullString)
        frmGH.rtbTopic(i).SelText = strTopic
    Case "331"  'if you recieve a message saying "no topic set"
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        For i = 1 To frmGH.rtbTopic.UBound
            If LCase$(frmGH.rtbTopic(i).Tag) = LCase$(strDestTab) Then
                Exit For
            End If
        Next i
        frmGH.rtbTopic(i).SelStart = 0
        frmGH.rtbTopic(i).SelLength = 666
        frmGH.rtbTopic(i).SelText = "No topic set"
    Case "324" 'displays the channel modes and channel key /mode #gta2gh
        displaychat strDestTab, vbRed, "No soup for you!"
    Case "001" 'welcome message
        ':bananaphone.nl.eu.gtanet.com 001 Sektor :Welcome to the GTANet IRC Network
        strNick = processParam(params$)
        If InStr(params$, "@") Then
            strExternalHostName = Right$(params$, Len(params$) - InStr(1, params$, "@"))
            
            Dim blnGetCC As Boolean
    
            With cr
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                .ValueType = REG_SZ
                .ValueKey = "ExternalHostName"
                If .Value = vbNullString Or .Value <> strExternalHostName Then
                    .Value = strExternalHostName
                    'You hostname has changed. Your country needs to be detected again.
                    blnGetCC = True
                End If
            End With
            
            'If there's no country saved in the registry or hostname has changed then detect country
            If blnGotCC = False Then
                If strCountryCode = vbNullString Or blnGetCC = True Then Call frmGH.getCC
            End If
        
            If blnCountryDetectFail And Left$(Right$(strExternalHostName, 3), 1) = "." Then
                For i = 1 To UBound(strCountries)
                    'Your hostname ended in a CC, saving it to registry
                    If UCase$(Right$(strExternalHostName, 2)) = Right$(strCountries(i), 2) Then
                        blnCountryDetectFail = False
                        strCountryCode = Right$(strCountries(i), 2)
                        intCountryIndex = i
                        displaychat strChannel, strServerColor, "Country code set based on your hostname: " & strCountryCode
                        With cr
                            .ClassKey = HKEY_CURRENT_USER
                            .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                            .ValueType = REG_SZ
                            .ValueKey = "Country"
                            .Value = strCountryCode
                        End With
                        Exit For
                     End If
                Next i
            End If
        End If
      'frmGH.Caption = "GTA2 Game Hunter v" & TXT_GHVER & " - IRC name: " & strNick
    Case "010" '010 Claude frosties.de.eu.gtanet.com 6667 :Please use this Server/Port instead
      displaychat strDestTab, vbRed, processRest(params$)
    Case "353"  'if we received the channel user list /names
        'display "<" + strFrom + "> " + strRest 'display the unprocessed message
        send "WHO " & strDestTab
        Dim strNick2, othernicks$    'take one nick at a time
        othernicks$ = processParam(processRest(processRest(processRest(params$))))   'cut off the channel parameter, the nick parameter and the "="
        
        '''BenMillard''' moved retrieving the active channel ID outside of the subsequent loop:
        j = getPlayerLV(strDestTab)
        'Do we have Chat controls for this channel?
        If j = -1 Then
            'Create the channel controls:
            displaychat strChannel, strGHColor, Replace(processRest(processRest(processRest(params$))), ":", vbNullString)
            Exit Sub
        End If
        
        '''BenMillard''' '''FOCUS''' select the newly joined channel directly:
        'Debug.Print "Switch to new channel, " & strChannel
        Call frmGH.ShowChatByPlayerName(strDestTab)
        
        Do
            strNick2 = processParam(othernicks$)   'take one nick
            othernicks$ = processRest(othernicks$)   'and take it out of the remaining nicks
            
            'cut off the operator flags at the beginning of names
            Select Case Left$(strNick2, 1)
                Case "@", "+", "~", "&", "%"
                    strNick2 = Right(strNick2, Len(strNick2) - 1) 'cut off the first character
            End Select
            
            i = 0
            
            Do
                i = i + 1
                If i > frmGH.lvPlayers(j).ListItems.count Then 'if the user is not found ...
                    i = -1     'set the user to be removed to -1 (ERROR :-) )
                    Exit Do     'exit the loop
                End If
            Loop Until frmGH.lvPlayers(j).ListItems.Item(i) = strNick2  'loop until we find the user
        
            If i = -1 Then
               With frmGH.lvPlayers(j).ListItems.Add(, , strNick2)
                  .ListSubItems.Add , , vbNullString
                  .ListSubItems.Add , , vbNullString, 0
                  .ListSubItems.Add , , vbNullString
               End With
               
               'Check if the player has been assigned a color this session
'               j = -1
'               For i = 0 To UBound(strPlayerColors)
'                    If strPlayerColors(i, 0) = vbNullString Then Exit For
'
'                    If LCase(strPlayerColors(i, 0)) = LCase(strNick2) Then
'                        j = i
'                        Exit For
'                    End If
'               Next
'
            
               
'                For i = 0 To UBound(strPlayerColors)
'                   If strPlayerColors(i, 0) = vbNullString Then
'                        strPlayerColors(i, 0) = strNick2
'                        strPlayerColors(i, 1) = i * 1000
'                        Exit For
'                    End If
'                Next
              
               
                
                
            End If
        Loop Until othernicks$ = vbNullString     'loop through all the received nicknames
        
        If blnLogin = False Then
            blnLogin = True
            'displaychat strDestTab, strConnectionColor, "Login complete"  'GH message
            If strExternalHostName = vbNullString Then
                displaychat strChannel, vbRed, "Unable to get your IP address from the IRC server. No one will be able to join your games."
                'send "USERHOST " & strNick
            End If
            
            '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
            frmGH.rtbChatbox(1).SelLength = 0
            frmGH.tabIRC.Tabs(1).Selected = True
            blnDisconnectClick = False
            If strStatusMsg = "A" Then
                send "NOTICE " & strChannel & " SA"
                Call frmGH.changeStatus(strStatusMsg)
                If Trim(strAwayMsg) = vbNullString Then strAwayMsg = "busy"
                'send "AWAY :" & strAwayMsg
            End If
            
            If bln98 Then send "NOTICE " & gta2ghbot & " 98"
        End If
    Case "301" 'whois away message
         'strDestTab = processParam(processRest((params$))) 'display in private tab
         strDestTab = frmGH.tabIRC.SelectedItem 'display in current tab
         displaychat strDestTab, strConnectionColor, processParam(processRest((params$))) & " is away: " & processParam(processRest(processRest(params$)))
    Case "302" ' /USERHOST reply
          'do nothing
    Case "305" 'You are no longer marked as being away
          '''displaychat strDestTab, strConnectionColor, processParam(processRest((params$)))
    Case "306" 'You have been marked as being away
          '''displaychat strDestTab, strConnectionColor, processParam(processRest((params$)))
    Case "307" 'is a registered nick
          displaychat frmGH.tabIRC.SelectedItem.Caption, strConnectionColor, Replace(processRest(params$), ":", vbNullString)
    Case "311" ' /WHOIS host ident and name(GHversion)
          strDestTab = frmGH.tabIRC.SelectedItem 'display in current tab
          displaychat strDestTab, strConnectionColor, vbNullString
          displaychat strDestTab, strConnectionColor, processParam(processRest((params$))) _
            & " is " & processParam(processRest(processRest(processRest(params$)))) & _
                " " & processParam(processRest(processRest(processRest(processRest( _
                processRest(params$))))))  'wow this is ugly
    Case "312" ' is using network
          strDestTab = frmGH.tabIRC.SelectedItem 'display in current tab
          displaychat strDestTab, strConnectionColor, processParam(processRest((params$))) _
            & " is using " & processParam(processRest(processRest(params$)))
    Case "317" ' idle time
         strDestTab = frmGH.tabIRC.SelectedItem 'display in current tab
         displaychat strDestTab, strConnectionColor, strText & " has been idle for " & Duration(processParam(processRest(processRest(params$))), 2) & ", connected since " & DateAdd("s", processParam(processRest(processRest(processRest(params$)))), #1/1/1970#)
    Case "319" ' is on these channels
         strDestTab = frmGH.tabIRC.SelectedItem 'display in current tab
         displaychat strDestTab, strConnectionColor, processParam(processRest((params$))) _
            & " is in " & processParam(processRest(processRest(params$)))
    Case "332"    'TOPIC real
        strTopic = Mid$(params$, InStr(1, params$, ":") + 1, Len(params$))
        strTopic = Replace$(strTopic, vbNullChar, vbNullString)
        strTopic = Replace$(strTopic, Chr$(1), vbNullString)
        strTopic = Replace$(strTopic, Chr$(2), vbNullString)   '
        strTopic = Replace$(strTopic, Chr$(3), vbNullString)   '
        strTopic = Replace$(strTopic, Chr$(15), vbNullString)
        
        displaychat strDestTab, strServerColor, "Topic is " & strTopic 'Display the channel topic
        
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        If frmGH.rtbTopic.UBound > 0 Then
            For i = 1 To frmGH.rtbTopic.UBound
                If LCase$(frmGH.rtbTopic(i).Tag) = LCase$(strDestTab) Then
                    Exit For
                End If
            Next i
        End If
        
        frmGH.rtbTopic(i).SelStart = 0
        frmGH.rtbTopic(i).SelLength = 666
        frmGH.rtbTopic(i).SelText = strTopic
         
    Case "333" 'TOPIC set by name on date
         strtest = strArray(strData)
         displaychat strtest(3), strServerColor, "Topic set by " & processParam(processRest(processRest(params$))) _
         & " on " & Format(DateAdd("s", processParam(processRest(processRest(processRest(params$)))), #1/1/1970#))
    'Case "340" '/USERIP
    Case "341"
        '341 RPL_INVITING
        '"<channel> <nick>"
        'Returned by the server to indicate that the
        'attempted INVITE message was successful and is
        'being passed onto the end client.
        strDestTab = frmGH.tabIRC.SelectedItem 'display in current tab
        displaychat strDestTab, strGHColor, processParam(processRest((params$))) _
                  & " has been invited to " & processParam(processRest(processRest(params$)))
    Case "375"    'MOTD message of the day
        send "JOIN " & strChannel & " " & strKey
        If strNick <> strPreferedNick Then
            If strPassword <> vbNullString Then
                'send "NS GHOST " & strPreferedNick & " " & strPassword
                send "NS RECOVER " & strPreferedNick & " " & strPassword
            End If
        End If
        
        If strPassword = vbNullString Then
            strPassword = "x"
            send "NS IDENTIFY " & strPassword
        End If
        
        'If blnHidden = False Or strServer(0) = "127.0.0.1" Then
        '    send "mode " & strNick & " -x"
        '    send "JOIN " & strChannel & " " & strKey
        '    Call joinChannels
        'End If
    Case "372", "376", "002", "003", "004", "005", "251", "252", "253", "254", "255", "265", "266", "315", "318", "366"
        'Make sure nick is owned before joining chan, previously it just joined no matter what
    Case "378" 'whois is connecting from
        strDestTab = frmGH.tabIRC.SelectedItem 'display in current tab
        strString = processRest(params$)
        'strString = "Sektor :is connecting from *@2001:0db8:85a3:0000:0000:8a2e:0370:7334"
        displaychat strDestTab, strConnectionColor, Replace(strString, ":", vbNullString, 1, 1)
    Case "396" 'is now your displayed host
        'no nothing
    Case "401"  'no such nick 'nickserv unavailable
         displaychat frmGH.tabIRC.SelectedItem, strGHColor, strText & " " & Right$(params$, Len(params$) - InStr(params, ":"))
            'If InStr(params$, "NickServ") Then
                If strText = strChannel Then
                    send "JOIN " & strText & " " & strKey
                Else
                    send "JOIN " & strText
                End If
            'End If
    'Cse  "402" 'no such server
    Case "403" 'no such channel
         send "JOIN " & strChannel & " " & strKey
    Case "404"  'No external channel messages , not on a channel, you are banned, Cannot send to channel (+m)
        strString = (Mid$(params$, InStr(params$, ":") + 1, 666))
        If strString <> "You are banned (" & strChannel & ")" Then
            displaychat strDestTab, strGHColor, strString
            If strString <> "Cannot send to channel (+m)" Then
                If strDestTab = strChannel Then
                    send "JOIN " & strChannel & " " & strKey
                Else
                    send "JOIN " & strDestTab
                End If
            End If
        End If
    Case "412", "351" 'no text to send
         displaychat strDestTab, strGHColor, strText
    Case "431"  'if we failed to change the nickname or didn't supply a name eg /whois blank
        displaychat strDestTab, strTextColor, "You have to enter a name"  'let them know that it failed
    'Case "461"  'not enough parameters
    Case "473" 'channel is +i 473 invite only
        If strDestTab = strChannel Then
            displaychat strChannel, vbRed, "You need to update to the latest Game Hunter: https://gtamp.com/gh"
            frmGH.cmdDisconnectClick
        Else
            displaychat strChannel, strGHColor, strText & " is invite only"
        End If
    Case "475"  'channel is +k 475 needs a key
        If strDestTab = strChannel Then
            displaychat strChannel, vbRed, "You need to update to the latest Game Hunter: https://gtamp.com/gh"
            frmGH.cmdDisconnectClick
        Else
            displaychat strChannel, strGHColor, strText & " needs a key"
        End If
    Case "471"  'channel is +l 471 user limit reached
        If strDestTab = strChannel Then
            displaychat strChannel, vbRed, "You need to update to the latest Game Hunter: https://gtamp.com/gh"
            frmGH.cmdDisconnectClick
        Else
            displaychat strChannel, strGHColor, strText & " is full"
        End If
    Case "474"  'channel is +b (you are banned)
        If strDestTab = strChannel Then
            With cr
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                .ValueKey = "BanReason"
                displaychat strDestTab, strBannedColor, "Access denied! Last kick details: " & .Value
            End With
            frmGH.cmdDisconnectClick
        Else
            displaychat strChannel, strGHColor, "You are banned from " & strText
        End If
    Case "432"  'if we failed to change the nickname
        displaychat strDestTab, strGHColor, Replace(processRest(params$), ":", vbNullString)
        If blnLogin = False Then
            strNick = "Ped" & Int(Rand(0, 9999))
            send "NICK " & strNick 'change name to new name
            frmGH.cmdOptions_Click
        End If
    Case "433"  'Nickname is already in use
        displaychat frmGH.tabIRC.SelectedItem.Caption, strConnectionColor, Replace(processRest(params$), ":", vbNullString)
        strNick = "Ped" & Int(Rand(0, 9999))
        intNickservWaitTime = -1
        send "NICK " & strNick
        If strPassword <> vbNullString Then
            'send "NS GHOST " & strPreferedNick & " " & strPassword
            send "NS RECOVER " & strPreferedNick & " " & strPassword
        End If
    Case "671" 'is using a secure connection
        displaychat frmGH.tabIRC.SelectedItem.Caption, strConnectionColor, Replace(processRest(params$), ":", vbNullString)
    Case "972" ' can't kick channel owner
        displaychat strDestTab, strGHColor, vbNullString & strData
    Case "352" '/WHO reply RPL_WHOREPLY
        Dim strParams(30) As String
        Dim blnNumeric As Boolean
        
        For i = 1 To Len(strRest)
            If Mid$(strRest, i, 1) = " " Then
                j = j + 1
                i = i + 1
            End If
            strParams(j) = strParams(j) & Mid$(strRest, i, 1)
        Next i
        
        Select Case strParams(6)
            Case "MrWhoopee"
                Call updateCountry(strParams(6), "IC", strDestTab)
                Exit Sub
            Case "Sally"
                Call updateCountry(strParams(6), "64", strDestTab)
                Exit Sub
        End Select
        
        'strParams(9) is the player's full name
        'strParams(6) is the player's nick
        
        'If player's full name starts with GH then use the last two
        'characters of their name as their CC
        If Left$(strParams(9), 2) = "GH" Then
            Dim strGHver As String
            If IsNumeric(Right$(strParams(9), 1)) Then
                strGHver = Mid$(strParams(9), 3, 666)
                blnNumeric = True
            Else
                strGHver = Mid$(strParams(9), 3, Len(strParams(9)) - 4)
            End If
                
            Set itmX = frmGH.lvPlayers(1).FindItem(strParams(6))
            If Not itmX Is Nothing Then
                'Store GH version as tooltip in player column
                itmX.ToolTipText = strGHver
                'If the player is hosting a game, update the GH version listitem with whois ghver
                Set itmX = frmGH.lvGames(0).FindItem(strParams(6))
                If Not itmX Is Nothing Then
                    frmGH.lvGames(0).ListItems(itmX.Index).ListSubItems(4) = strGHver
                End If
            End If
            
            'if the last two characters of their name is a number then it's not a country
            If Right$(strRest, 2) = "64" Then blnNumeric = False '64 is a C64 novelty country
            
            If blnNumeric = False Then
                Call updateCountry(strParams(6), Right$(strRest, 2), strDestTab)
            End If
        End If
         
        If Left$(strParams(9), 2) <> "GH" Or blnNumeric = True Then
            'If player's full name ends in a space and two letters then use that as their CC
            If strParams(2) = strChannel And Left$(Right$(strRest, 3), 1) = " " Then
                Call updateCountry(strParams(6), Right$(strRest, 2), strDestTab)
            Else
                'Set country based on hostname country code
                If Len(strParams(4)) >= 2 Then
                    If Mid$(strParams(4), Len(strParams(4)) - 2, 1) = "." Then
                        If Right$(strParams(4), 2) <> "IP" And IsNumeric(Right$(strParams(4), 2)) = False Then
                            Call updateCountry(strParams(6), Right$(strParams(4), 2), strDestTab)
                        End If
                    End If
                End If
            End If
        End If
      
    Case "900" 'You are now logged in as
'9790          displaychat strDestTab, strConnectionColor, Replace(processRest(params$), ":", vbNullString)
    Case "042" 'Your unique ID
        'do nothing
    Case "INVITE"
        displaychat strChannel, strConnectionColor, strFrom & " invited you to " & strText
    Case Else   'if it's another message
        If Mid$(params$, InStr(params$, " ") + 1, 1) = ":" Then
            strParams(1) = vbNullString
            strParams(3) = vbNullString
        Else
            strParams(0) = Left$(params$, InStr(params$, ":")) '  - 2)
            If Len(strParams(0)) > 2 Then strParams(0) = Left$(strParams(0), Len(strParams(0)) - 2)
            i = InStr(strParams(0), " ")
            strParams(0) = Mid$(strParams(0), i + 1, 666)
            strParams(1) = strParams(0) & " "
            i = InStr(strParams(1), " ")

            If i Then
                strParams(1) = Left$(strParams(1), i - 1) & " "
                strParams(3) = " " & Mid$(strParams(0), i + 1, 666)
            Else
                strParams(1) = vbNullString
            End If
        End If

        strParams(2) = Mid$(params$, InStr(params$, ":") + 1, 666)
        'displaychat frmGH.tabIRC.SelectedItem, strConnectionColor, strData
        displaychat frmGH.tabIRC.SelectedItem, strConnectionColor, Trim(strParams(1) & strParams(2) & strParams(3))
End Select

    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    send "PRIVMSG " & gta2ghbot & " :processing error: " & strErrdesc & " - Line: " & strErrLine
    displaychat strDestTab, strTextColor, "Error processing a message: " & strErrdesc & " - Line: " & strErrLine
End Sub

Public Sub updateCountry(strPlayer, strCC, strDestTab)
On Error GoTo oops
'search player list for the player
Dim itmX As ListItem
Dim itmY As ListItem
Dim intPlayerList As Integer

'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
For intPlayerList = 1 To frmGH.lvPlayers.UBound
    Set frmGH.lvPlayers(intPlayerList).SmallIcons = frmGH.ImageList1 'shouldn't need this line
    Set itmX = frmGH.lvPlayers(intPlayerList).FindItem(strPlayer)
    If itmX Is Nothing Then
        'do nothing
    Else
        Set itmY = frmGH.lvGames(0).FindItem(strPlayer)
    
        strCC = UCase$(strCC)
    
        'find the index of the country, so we can add the flag
        Dim i As Integer
        For i = 1 To UBound(strCountries)
            'if we found the index of the country
            If Right$(strCountries(i), 2) = strCC Then
                'add the users country code to the CC list
                'add the country flag to first column
                
                With itmX
                    .SmallIcon = i + 1
                    'add the country as a tooltip
                    '.ToolTipText = Left$(strCountries(i), Len(strCountries(i)) - 5)
                    '.ListSubItems(1).ToolTipText = .ToolTipText
                    .ListSubItems(1).ToolTipText = Left$(strCountries(i), Len(strCountries(i)) - 5)
                    .ListSubItems(1).Text = strCC
                    
                    'If the player is hosting a game then also update the CC info there
                    If Not itmY Is Nothing Then
                        With itmY
                            .SmallIcon = i + 1
                            .ListSubItems(2).ToolTipText = Left$(strCountries(i), Len(strCountries(i)) - 5)
                            .ListSubItems(2).Text = strCC
                        End With
                    End If
                End With
                Exit For
            Else
                With itmX
                    .SmallIcon = 1
                    '.ToolTipText = "Unknown"
                    .ListSubItems(1).ToolTipText = "Unknown"
                    .ListSubItems(1).Text = vbNullString
                End With
            End If
        Next
        
        Call SortColumn(frmGH.lvPlayers(intPlayerList), frmGH.lvPlayers(intPlayerList).SortKey + 1)
    End If
Next

Exit Sub
oops:
    
    strErrdesc = Err.Description
    strErrLine = Erl
    strErrNum = Err.Number
    send "PRIVMSG " & gta2ghbot & " :updateCountry: " & strErrNum & " - Line: " & strErrLine
    displaychat strDestTab, strTextColor, "updating country error: " & strErrdesc
End Sub

Public Function getPlayerLV(strChanName) As Integer

On Error GoTo oops
Dim i As Integer
'Search the lvPlayers control arrays for the current channel name

'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
For i = 1 To frmGH.lvPlayers.UBound
    If LCase$(frmGH.lvPlayers(i).Tag) = LCase$(strChanName) Then
        getPlayerLV = i 'return the index
        Exit Function 'found it, so stop looping and exit function
    End If
Next i

'Didn't find a player list for this channel:
getPlayerLV = -1 '''BenMillard''' noticed this isn't handled in lots of places

Exit Function
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    '''BenMillard''' commented this out to reduce me spamming Sektor:
    send "PRIVMSG " & gta2ghbot & " :getPlayerLV: " & strErrdesc & " - Line: " & strErrLine
    displaychat strDestTab, strTextColor, "getPlayerLV error: " & strErrdesc & " - Line: " & strErrLine
End Function

Public Sub joinChannels()
On Error GoTo LogError

Dim i As Integer
Dim strChannels As String
Dim strCustomChan As String
    
With cr
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter"
    .ValueKey = "Channels"
    If .Value = vbNullString Then
        strChannels = strChannel ' & "#opengta2"
    Else
        strChannels = .Value
    End If
End With

For i = 1 To Len(strChannels)
    If Mid$(strChannels, i, 1) = "#" Or i = Len(strChannels) Then
        If i = Len(strChannels) Then strCustomChan = strCustomChan & Mid$(strChannels, i, 1)
        If i > 2 Then
            send "JOIN #" & strCustomChan
            strCustomChan = vbNullString
        End If
    Else
        strCustomChan = strCustomChan & Mid$(strChannels, i, 1)
    End If
Next

Exit Sub

LogError:

End Sub

'Checks if a player is in at least one channel that you are in
Public Function onYourChannel(ByVal strName As String) As Boolean
    Dim i As Integer
    Dim j As Integer
    
    onYourChannel = True
    strName = LCase(strName)
    
    If strName <> LCase(gta2ghbot) Then 'messages from gta2ghbot are always received
        'Search all player lists to see if you are in a channel with that player
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        For i = 1 To frmGH.lvPlayers.UBound
            For j = 1 To frmGH.lvPlayers(i).ListItems.count
                If strName = LCase(frmGH.lvPlayers(i).ListItems.Item(j)) Then
                    i = -1
                    Exit For 'found it, so stop looping
                End If
            Next j
            If i = -1 Then Exit For 'found it, so stop looping
        Next i
               
        'Ignore the message if you aren't in a channel with that player
        If i > -1 Then onYourChannel = False
    End If
End Function
