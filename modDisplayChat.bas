Attribute VB_Name = "modDisplayChat"
Option Explicit
Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

'add a message to the chat history
Public Sub displaychat(ByVal strDestTab As String, ByVal strColor As String, ByVal strMsg As String, Optional blnSkip As Boolean)

Dim intDestinationTab As Integer
Dim intChatHistory As Integer
Dim i As Integer
Dim j As Long

On Error GoTo oops

If strColor = vbNullString Then
    If frmGH.rtbHistory(0).BackColor = 15658734 Then
        strColor = vbBlack
    Else
        strColor = 15658734
    End If
End If

'Replace all control codes excluding Line Feed and Carriage Return
For i = 1 To 31
    If i <> 10 And i <> 13 Then strMsg = Replace$(strMsg, Chr$(i), vbNullString)
Next

If strFontName = "DejaVu Sans Mono" Then
    strMsg = Replace$(strMsg, Chr$(133), vbNullString) '…
    strMsg = Replace$(strMsg, Chr$(135), vbNullString) '‡
    strMsg = Replace$(strMsg, Chr$(155), vbNullString) '›
    strMsg = Replace$(strMsg, Chr$(196), vbNullString) 'Ä
End If

If blnchkTime = True Then strMsg = Time & " " & strMsg

intChatHistory = -1

'Search for existing chat history
'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
'''Will need to refactor all these index values now!
'''BenMillard''' thinks this kind of duplicates the new ShowChatByPlayerName
For i = 1 To frmGH.rtbHistory.UBound
    If UCase$(frmGH.rtbHistory(i).Tag) = UCase$(strDestTab) Then
        'if the tab is not a channel then change topic to "Private chat with name"
        If Left$(strDestTab, 1) <> "#" Then frmGH.rtbTopic(i).Text = TXT_PRIVATE & strDestTab
        frmGH.rtbTopic(i).Tag = strDestTab
        intChatHistory = i
        Exit For
    End If
Next i

intDestinationTab = -1

'''BenMillard''' thinks this kind of duplicates the new ShowChatByPlayerName
If strDestTab <> vbNullString Then
    For i = 1 To frmGH.tabIRC.Tabs.count
        If UCase$(frmGH.tabIRC.Tabs.Item(i).Caption) = UCase$(strDestTab) Then
            frmGH.tabIRC.Tabs.Item(i).Caption = strDestTab 'update caption just in case the casing is different
            intDestinationTab = i
            Exit For
        End If
    Next i
    
    '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
    If intDestinationTab < 0 Then
        With frmGH
            .tabIRC.Tabs.Add , , strDestTab
            intDestinationTab = .tabIRC.Tabs.count
            
            'If there isn't a hidden chat then load a new rtb
            '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
            If intChatHistory < 0 Then
                Load .rtbHistory(.rtbHistory.UBound + 1)
                Load .rtbTopic(.rtbTopic.UBound + 1)
                Load .rtbChatbox(.rtbChatbox.UBound + 1)
                intChatHistory = frmGH.rtbHistory.UBound
                
                '''EnableAutoURLDetection .rtbHistory(.rtbHistory.ubound)
                With .rtbTopic(.rtbTopic.UBound)
                    .Tag = strDestTab
                    .Text = vbNullString
                End With
                
                '''EnableAutoURLDetection .rtbTopic(.rtbChatbox.ubound)
                With .rtbHistory(frmGH.rtbHistory.UBound)
                    .Tag = strDestTab
                    .Text = vbNullString
                End With
                
                With .rtbChatbox(.rtbChatbox.UBound)
                    .Tag = strDestTab
                    '.ToolTipText = strDestTab
                    .Text = vbNullString
                    .SelStart = 0
                    .SelLength = 500
                    .SelColor = strTextColor
                    .SelLength = 0
                End With
            End If
            intDestinationTab = .tabIRC.Tabs.count
        End With
    End If
    
    If blnPrivmsg = True Then
        With frmGH.tabIRC
            If intDestinationTab <> .SelectedItem.Index Then
                .Tabs.Item(intDestinationTab).HighLighted = True
            End If
        End With
    End If
Else
    intDestinationTab = 1
End If

'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
If intChatHistory < 1 Then intChatHistory = 1

'Add strMsg to rtbHistory and apply color
With frmGH.rtbHistory(intChatHistory)
    .SelStart = Len(.Text)
    .SelColor = strColor
    .SelUnderline = False
    .SelBold = blnFontBold
    .SelItalic = blnFontItalic
    .SelFontName = strFontName
    .SelText = vbNewLine & strMsg
    
    'highlight URLs
    Dim strArray() As String
    Dim strClean As String
    strArray = Split(strMsg)
    j = Len(.Text) - Len(strMsg)
    
    For i = 0 To UBound(strArray)
        
        If strArray(i) <> strDestTab Then 'if you are in the channel already then don't make the channel name a link
            If Left$(strArray(i), 1) = "#" Then
                .SelStart = j
                .SelLength = Len(strArray(i))
                .SelColor = strLinkColor
                .SelUnderline = False
                .SelBold = blnFontBold
                .SelItalic = blnFontItalic
                .SelFontName = strFontName
                .SelText = .SelText
            Else
                If isURL(strArray(i)) = True Then
                    strClean = cleanURL(strArray(i))
                    .SelStart = (j - 1) + InStr(strArray(i), strClean)
                    .SelLength = Len(strClean)
                    .SelColor = strLinkColor
                    .SelUnderline = blnUnderline
                    .SelBold = blnFontBold
                    .SelItalic = blnFontItalic
                    .SelFontName = strFontName
                    .SelText = .SelText
                End If
            End If
        End If
        
        j = j + Len(strArray(i)) + 1
    Next

    If blnHighlight = False Then Exit Sub
    
    'Colorize alert words
    Dim x As Integer
    For x = 0 To UBound(strAlertWords)

        i = InStr(1, UCase$(strMsg), UCase$(strAlertWords(x)))
        j = Len(strMsg)
        
        If i Then
            'If blnchkSoundWordAlert = True Then PlaySound strSoundWordAlert, ByVal 0&, SND_ASYNC
            .SelStart = Len(.Text) - (j - i) - 1
            .SelLength = Len(strAlertWords(x))
            .SelColor = strActionColor
            .SelBold = blnFontBold
            .SelItalic = blnFontItalic
            .SelFontName = strFontName
            .SelText = .SelText
        End If
    Next x
    
    'Colorize player names
'    For x = 0 To UBound(strPlayerColors)
'
'        i = InStr(1, UCase$(strMsg), UCase$(strPlayerColors(x, 0)))
'        j = Len(strMsg)
'
'        If i Then
'            .SelStart = Len(.Text) - (j - i) - 1
'            .SelLength = Len(strPlayerColors(x, 0))
'            .SelColor = Val(strPlayerColors(x, 1))
'            .SelBold = blnFontBold
'            .SelItalic = blnFontItalic
'            .SelFontName = strFontName
'            .SelText = .SelText
'        End If
'    Next x
        
End With
    
Exit Sub

oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    Debug.Print "displaychat error: " & Err.Description & " " & strErrLine
End Sub

Sub send(strMsg)  'send a message to the IRC server
On Error GoTo oops
    If blnConnected = True Then
        frmGH.sockIRC.SendData strMsg & vbNewLine 'send the data, along with a carrige return and a line feed
    Else
        displaychat strDestTab, strGHColor, "Sign in to send messages."
        Exit Sub
    End If
    
    Exit Sub    'skip the error handling section
oops:
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Unable to send data.  Not connected."
    blnConnected = False
    blnLogin = False
End Sub

'''FOCUS'''
Public Sub giveFocus(varFocus As Variant)
On Error Resume Next
varFocus.SetFocus
End Sub

'Returns true if string looks like a URL
Public Function isURL(ByVal strLow As String) As Boolean
    strLow = LCase$(strLow)
    
    If Left$(strLow, Len(strChannel)) = frmGH.tabIRC.SelectedItem Then Exit Function
    
    If InStr(strLow, "www.") Or InStr(strLow, "://") Or Left$(strLow, 1) = "#" Then
        isURL = True
        Exit Function
    Else
        isURL = False
    End If

End Function
