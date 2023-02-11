VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX%"
Begin VB.Form frmCreateGame 
   AutoRedraw      =   -1  'True
   Caption         =   "Create Game"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8970
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Ra&ndom"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      ToolTipText     =   "Random map"
      Top             =   5040
      Width           =   855
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      ItemData        =   "frmCreateGame.frx":0000
      Left            =   1440
      List            =   "frmCreateGame.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   855
      Left            =   6360
      TabIndex        =   13
      Top             =   120
      Width           =   1695
      Begin VB.CheckBox chkSync 
         Caption         =   "E&xit On Desync"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1455
      End
      Begin VB.CheckBox chkPlayReplay 
         Caption         =   "&Play Replay"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.Frame fraFilter 
      Caption         =   "&Find this map:"
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlayAlone 
      Caption         =   "Play &Alone"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame fraLock 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   12
      Top             =   120
      Width           =   1695
      Begin VB.CheckBox chkHostPassword 
         Caption         =   "Pass&word"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
      Begin VB.TextBox txtHostPassword 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lvMaps 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "STRING"
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "STRING"
         Text            =   "GMPfile"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "STRING"
         Text            =   "STYfile"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "STRING"
         Text            =   "SCRfile"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "STRING"
         Text            =   "MMP filename"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "STRING"
         Text            =   "PlayerCount"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdLaunchGTA2 
      Caption         =   "Create &Game"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Image imgMap 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   4815
   End
End
Attribute VB_Name = "frmCreateGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright GTAMP.com gtamulti@gmail.com
'License: Do whatever you want with this code. No warranty.
'The integrity of this product cannot be guaranteed for high voltage operation.
'The Zaibatsu Corporation reserves the right to change the specifications without notice.
'Conditions apply.

Option Explicit
Dim cr As New cRegistry
Dim strGMP As String
Dim strSTY As String
Dim strSCR As String
Dim strLastGMP As String
Dim strLastSTY As String
Dim strLastSCR As String

Dim strLastPicFileName As String
Const SIDESPACE = 240
Const STANDARD_SPACE = 120

Private Sub cboFilter_click()
    Call saveSettings
    Call Host
End Sub

Private Sub chkSync_Click()
blnChkSync = chkSync.Value
With cr
.ClassKey = HKEY_CURRENT_USER
.SectionKey = "Software\DMA Design Ltd\GTA2\Debug\"
.ValueKey = "do_sync_check"
.ValueType = REG_DWORD
If chkSync = vbChecked Then
    .Value = 0
    blnChkSync = True
Else
    blnChkSync = False
    .DeleteValue
End If
End With
End Sub

Private Sub cmdRandom_Click()
    If lvMaps.ListItems.count = 0 Then Exit Sub
    lvMaps.ListItems(Rand(1, lvMaps.ListItems.count)).Selected = True
    lvMaps.SelectedItem.EnsureVisible
    Dim itmRandom As MSComctlLib.ListItem
    Call lvMaps_ItemClick(itmRandom)
End Sub

Private Sub txtFind_change()
    Call saveSettings
    strLastGMP = vbNullString
    strLastSTY = vbNullString
    strLastSCR = vbNullString
    Call Host
End Sub

Private Sub txtFind_Keypress(key As Integer)
    'Select text if ctrl+a is pushed
    If key = 1 Then
        key = 0
        With txtFind
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub

Private Sub txtHostPassword_Keypress(key As Integer)
    'Select text if ctrl+a is pushed
    If key = 1 Then
        key = 0
        With txtHostPassword
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub

Private Sub chkHostPassword_Click()
    txtHostPassword.Enabled = chkHostPassword
End Sub

Public Sub cmdRefresh_Click()
    Call saveSettings
    If frmGH.PreHost = True Then Call Host
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call cmdClose_Click
End Sub

Private Sub lvMaps_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SortLV(frmCreateGame.lvMaps, ColumnHeader)
    Dim intUniqueCount As Integer
    Dim strPrevItem As String
    Dim i As Integer
    
    If ColumnHeader.Index > 1 Then
        For i = 1 To lvMaps.ListItems.count
            If LCase$(lvMaps.ListItems.Item(i).ListSubItems((ColumnHeader.Index) - 1).Text) <> strPrevItem Then
                If lvMaps.ListItems.Item(i).ListSubItems((ColumnHeader.Index) - 1).Text <> vbNullString Then intUniqueCount = intUniqueCount + 1
                'Debug.Print lvMaps.ListItems.Item(i).ListSubItems(ColumnHeader.Index).Text
                strPrevItem = LCase$(lvMaps.ListItems.Item(i).ListSubItems((ColumnHeader.Index) - 1).Text)
            End If
        Next
        
        With lvMaps.ColumnHeaders(ColumnHeader.Index)
            If InStr(.Text, "(") Then
                .Text = Left$(.Text, InStr(.Text, " ")) & "(" & intUniqueCount & ")"
            Else
                .Text = .Text & " " & "(" & intUniqueCount & ")"
            End If
        End With
    End If
    
If lvMaps.ListItems.count Then lvMaps.SelectedItem.EnsureVisible
End Sub

Private Sub lvMaps_DblClick()
'''Shortcut for hosting:
    Call cmdLaunchGTA2_Click
End Sub

Public Sub cmdPlayAlone_Click()
On Error GoTo oops
    Dim strReplay As String
    Dim i As Integer
    Dim j As Integer
    If lvMaps.ListItems.count = 0 Then
        Exit Sub
    End If
    
    i = lvMaps.SelectedItem.Index 'This stores the index of the currently selected map
   
    For j = 2 To lvMaps.ColumnHeaders.count
        If LCase$(lvMaps.ColumnHeaders.Item(j).Text) = "gmpfile" Then
            strGMP = lvMaps.ListItems.Item(i).ListSubItems(j - 1) & ".gmp"
        End If
        
        If LCase$(lvMaps.ColumnHeaders.Item(j).Text) = "styfile" Then
            strSTY = lvMaps.ListItems.Item(i).ListSubItems(j - 1) & ".sty"
        End If
        
        If LCase$(lvMaps.ColumnHeaders.Item(j).Text) = "scrfile" Then
            strSCR = lvMaps.ListItems.Item(i).ListSubItems(j - 1) & ".scr"
        End If
    Next
    
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\DMA Design Ltd\GTA2\Debug\"
        .ValueType = REG_SZ
        .ValueKey = "mapname"
        .Value = strGMP
        .ValueKey = "scriptName"
        .Value = strSCR
        .ValueKey = "stylename"
        .Value = strSTY
        .ValueType = REG_DWORD
        .ValueKey = "skip_frontend"
        .Value = 0
    End With
        
    Call setGTA2path
        
    If Exists(strGTA2path & TXT_GTA2EXE) Then
        If modVersion.DetectGTA2version = False Then Exit Sub
        If chkPlayReplay = vbChecked Then
            If Exists(strGTA2path & "test\replay.rep") Then strReplay = " -r"
        End If
        modMkDir strGTA2path & "test"
        Call shellandwait(strGTA2path & TXT_GTA2EXE & strReplay, strGTA2path)
        
    Else
        displaychat strDestTab, vbRed, "Can't find " & strGTA2path & TXT_GTA2EXE
    End If
    
    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error during Play Alone: " & strErrdesc & " - Line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :Play Alone error: " & strErrdesc & " Line: " & strErrLine
End Sub

Private Sub cmdLaunchGTA2_Click()
On Error GoTo oops
    If txtHostPassword <> vbNullString And txtHostPassword = strPassword Then
        displaychat strDestTab, strGHColor, "For security reasons, your game password should not be the same as your IRC password."
        Exit Sub
    End If
    
    If saveSettings = False Then Exit Sub
    
    If chkPlayReplay.Value = vbChecked Then
        blnPlayReplay = True
    Else
        blnPlayReplay = False
    End If
    
    frmGH.Host
    '''FOCUS'''Call frmGH.giveChatFocus
    Hide
    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error during launch: " & strErrdesc & " - Line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :cmdLaunchGTA2 error: " & strErrdesc & " Line: " & strErrLine
End Sub

Private Function saveWindowSettings()
    With cr
        'Save window size and position:
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
        .ValueKey = "HostWindowState"
        If WindowState <> vbMinimized Then .Value = WindowState
        If WindowState = vbNormal Then
            .ValueKey = "HostHeight"
            .ValueType = REG_SZ
            .Value = Height
            
            .ValueKey = "HostWidth"
            .ValueType = REG_SZ
            .Value = Width
            
            .ValueKey = "HostLeft"
            .ValueType = REG_SZ
            .Value = Left
            
            .ValueKey = "HostTop"
            .ValueType = REG_SZ
            .Value = Top
        End If
    End With
End Function

Private Function saveSettings() 'Why is this function called twice when creating a game???
    Dim i As Integer
    Call saveWindowSettings
    
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
        'lvMaps save sort settings
        .ValueKey = "lvMaps.SortOrder"
        .ValueType = REG_SZ
        .Value = lvMaps.SortOrder
        .ValueKey = "lvMaps.SortKey"
        .ValueType = REG_SZ
        .Value = lvMaps.SortKey
        .ValueKey = "Filter"
        .Value = cboFilter.ListIndex
    End With

    saveSettings = False
    
    If lvMaps.ListItems.count = 0 Then Exit Function
    
    'Sort array in ASCII order, case sensitive
    'the sort really shouldn't be case sensitive but that was the easiest way to make it put [ before letters
    
    Dim SortedArray As New cSortArray
    'add all map descriptions from frGH.lvMMPlist to cSortArray (sorted array)
    For i = 1 To frmGH.lvMMPlist.ListItems.count
        SortedArray.AddItem LCase(Trim(frmGH.lvMMPlist.ListItems(i).ListSubItems(4).Text)) 'add map description to array (trim and convert to lcase)
    Next i
    
    'mpaddon If frmProbeGTA2.listenForGameName = False Then Exit Function
    
    'write host password, comment and other selection settings to registry
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueType = REG_SZ
        .ValueKey = "chkHostPassword"
        .Value = chkHostPassword
        .ValueKey = "HostPassword"
        .Value = txtHostPassword
        txtHostPassword.Enabled = chkHostPassword
    End With
    
    If txtHostPassword <> vbNullString And chkHostPassword = vbChecked Then
        strPasswordProtectGame = "Yes"
        strYourGamePassword = txtHostPassword
    Else
        chkHostPassword = vbUnchecked
        strYourGamePassword = vbNullString
        strPasswordProtectGame = "No"
    End If
   
    strGTA2MapDesc = lvMaps.SelectedItem
    
    'save the current map description to the registry, so it can be searched for next time
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueType = REG_SZ
        .ValueKey = "MapDesc"
        .Value = strGTA2MapDesc
    End With
    
    'Search the sorted array for the selected map
    i = SortedArray.SearchArray(LCase(strGTA2MapDesc))
    
    'Now that we found the index of the map, write it to the registry, so GTA2 can read it
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\DMA Design Ltd\GTA2\Network"
        .ValueKey = "map_index"
        .ValueType = REG_DWORD
        .Value = i
        
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\DMA Design Ltd\GTA2\Network"
        .ValueKey = "map_index"
        .ValueType = REG_DWORD
        .Value = i
                
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueKey = "map_desc"
        .ValueType = REG_DWORD
        .Value = vbNullString
    End With
    
    saveSettings = True
End Function

Private Sub lvMaps_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo oops:

'''Variables:
Dim strPreviewPath As String '''renamed for clarity

'''No maps?
If lvMaps.ListItems.count = 0 Then Exit Sub

'''Get paths for selected map:
strGMP = lvMaps.ListItems(lvMaps.SelectedItem.Index).ListSubItems(1) & ".gmp"
strSTY = lvMaps.ListItems(lvMaps.SelectedItem.Index).ListSubItems(2) & ".sty"
'txtFind = lvMaps.SelectedItem.Text

'''Get path for preview image:
strPreviewPath = strGTA2path & "data\" & lvMaps.ListItems(lvMaps.SelectedItem.Index).ListSubItems(1) '''GMPfile
'strPreviewPath =  ' & Left$(strPreviewPath, Len(strPreviewPath) - 4)

'''Already showing this preview?
If strPreviewPath = strLastPicFileName Then Exit Sub '''restructured

'''Search for new preview image:
strLastPicFileName = strPreviewPath
'if filename has changed then see if it exists
If Exists(strPreviewPath & ".gif") Then
    strPreviewPath = strPreviewPath & ".gif"
ElseIf Exists(strPreviewPath & ".jpg") Then
    strPreviewPath = strPreviewPath & ".jpg"
Else
    strPreviewPath = vbNullString '''no preview
End If

'''No preview found? '''Restructured
If Len(strPreviewPath) = 0 Then
     '''Moved from a labelled GOTO into this structure:
    strLastPicFileName = vbNullString
    imgMap.Picture = Nothing
    imgMap.Visible = False
    lvMaps.Width = frmCreateGame.ScaleWidth - SIDESPACE 'this map doesn't have an image :(
    Exit Sub
Else
    'Yay image exists!
    imgMap.Picture = LoadPicture(strPreviewPath)
    Call SizeImage(imgMap)
End If

'Update controls:
Call Form_Resize

Exit Sub

oops:
    strErrdesc = Err.Description
    strErrNum = Err.Number
    strErrLine = Erl
    If strErrNum = 481 Then
        displaychat strChannel, vbRed, strErrdesc & " " & strPreviewPath 'invalid picture
    Else
        displaychat strChannel, vbRed, "Error during map click: " & strErrdesc & " - Line: " & strErrLine & strErrNum
        send "PRIVMSG " & gta2ghbot & " :Error during map click: " & strErrdesc & " Line: " & strErrLine & " " & strErrNum
    End If
    imgMap.Picture = Nothing
    imgMap.Visible = False
    lvMaps.Width = frmCreateGame.ScaleWidth - SIDESPACE 'this map doesn't have an image :(
End Sub

'Scale image, right align and vertically centre
'Based on blindwig's code: http://www.xtremevbtalk.com/showpost.php?postid=526291
'Modified by Sektor to align right instead of centre and also resize lvMaps
'Frame/Container wasn't needed and it caused a graphic flicker, so it was removed
Public Sub SizeImage(imgTarget As Image)
On Error GoTo oops:
    If Visible = False Then Exit Sub
    '''Renamed to avoid confusion with form properties:
    Dim sngLeft As Single, sngTop As Single, sngWidth As Single, sngHeight As Single
    Dim AspectRate As Double
    
    '''Hide during resize:
    imgTarget.Visible = False
    imgTarget.Stretch = False
    
    '''Minimum width for lvMaps: (Logic adapted from frmGH.AutoSizeLV)
    With lvMaps
        '1st column + any scrollbar + borders
        sngWidth = .ColumnHeaders(1).Width
        
        'Scrollbar?
        If .Height < (.ListItems.count * 210) + 150 Then 'more items than can be shown:
            sngWidth = sngWidth + 150 'typical scrollbar width
        End If
        
        'Apply width only if it will be different:
        sngWidth = sngWidth + 60 'add borders
        If CLng(.Width) <> CLng(sngWidth) Then
            .Width = sngWidth
        End If
        
        '''lvMaps.Width = fraLock.Width + 400
        ''''Rebug.Print .Name; .ColumnHeaders(1).Width + 150 + 60; .Width;
    End With
    
    'Aspect ratio:
    AspectRate = imgTarget.Width / imgTarget.Height
    If AspectRate > (frmCreateGame.ScaleWidth _
                     - lvMaps.Width - lvMaps.Left _
                     - (STANDARD_SPACE * 3)) _
                     / (lvMaps.Height) Then '''includes gaps
        'Picture is Wide
        sngWidth = frmCreateGame.ScaleWidth _
                     - lvMaps.Width - lvMaps.Left _
                     - (STANDARD_SPACE * 3) '''must match If statement!
        sngHeight = sngWidth / AspectRate
    Else
        'Picture is High
        sngHeight = lvMaps.Height
        sngWidth = sngHeight * AspectRate
    End If
    
    'Move and size the image:
    sngLeft = frmCreateGame.ScaleWidth - (sngWidth + 120) 'as far right as possible
    sngTop = lvMaps.Top + (lvMaps.Height - sngHeight) / 2 'vertically centered with Maps list
    'sngTop = fraLock.Top + (lvMaps.Height - sngHeight) / 2
    imgTarget.Stretch = True
    imgTarget.Move sngLeft, sngTop, sngWidth, sngHeight
    
    'If the image size with correct ratio leaves some space then increase lvMaps size to fill
    imgTarget.Visible = True '''moved
    lvMaps.Width = imgTarget.Left - SIDESPACE
    ''''Rebug.Print ; ; lvMaps.Width
    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    strErrNum = Err.Number
    displaychat strDestTab, vbRed, "Error resizing image: " & strErrdesc & " - Line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :Resizing image: " & strErrdesc & " Line: " & strErrLine
End Sub

'Cancel
Private Sub cmdClose_Click()

On Error GoTo oops:
    
    'write do_sync_check to GTA2 reg key based on chkSync checkbox
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\DMA Design Ltd\GTA2\Debug\"
        .ValueKey = "do_sync_check"
        .ValueType = REG_DWORD
        If chkSync = vbChecked Then
            .Value = 0
            blnChkSync = True
        Else
            blnChkSync = False
            .DeleteValue
        End If
    End With
    
    Call saveSettings
    Hide
    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrNum = Err.Number
    displaychat strDestTab, vbRed, "Error closing Create Game form: " & strErrdesc & " - Line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :cmdClose_Click(): " & strErrdesc
End Sub

Private Sub Form_Resize()
On Error GoTo oops
'''Resizable.
'Variables:

Dim lngWidth As Long '''BenMillard

'Sizing limits:
If WindowState = vbMinimized Then Exit Sub

If Width < 8000 Then Width = 8000
If Height < 4000 Then Height = 4000

'Buttons:
cmdLaunchGTA2.Top = ScaleHeight - 120 - cmdLaunchGTA2.Height
cmdPlayAlone.Top = cmdLaunchGTA2.Top
cmdClose.Top = cmdLaunchGTA2.Top
cmdClose.Left = ScaleWidth - 120 - cmdClose.Width
cmdPlayAlone.Left = cmdClose.Left - 120 - cmdPlayAlone.Width
cmdLaunchGTA2.Left = cmdPlayAlone.Left - 120 - cmdLaunchGTA2.Width
cmdRefresh.Top = cmdClose.Top
cmdRandom.Top = cmdClose.Top
cboFilter.Top = cmdClose.Top + 50

'Options frame is right aligned, constant width:
fraOptions.Left = ScaleWidth - 120 - fraOptions.Width

'Map list and Preview Image:
lvMaps.Height = cmdLaunchGTA2.Top - 120 - lvMaps.Top
If imgMap.Visible Then
    Call SizeImage(imgMap)
    lngWidth = lvMaps.Width
End If

'if the preview is too small for the password box to fit above it then don't reserve any room for the image
If imgMap.Visible = False Or lngWidth > Width - 4000 Then
    lvMaps.Width = frmCreateGame.ScaleWidth - 120 - 120 '''no preview
    lngWidth = lvMaps.Width / 3 '''search frame uses 1/3 of full width
End If

'''Search frame matches Maps list:
fraFilter.Width = lngWidth

txtFind.Width = lngWidth - 120 - 120

'''Password frame fills remaining space:
fraLock.Left = 120 + lngWidth + 120
fraLock.Width = fraOptions.Left - 120 - fraLock.Left
txtHostPassword.Width = fraLock.Width - 120 - 120
chkHostPassword.Width = chkHostPassword.Font.Weight * 3

'After resize, ensure the selected map is visible
If lvMaps.ListItems.count Then lvMaps.SelectedItem.EnsureVisible

Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error resizing host window: " & strErrdesc & " - Line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :error resizing host window: " & strErrdesc & " Line: " & strErrLine

End Sub
        
Public Sub loadDisplaySettings()
'Load display settings
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
        .ValueKey = "HostWindowState"
        
        'If WindowState = vbMinimized Then WindowState = vbNormal
        WindowState = Val(.Value)
        
        If .Value <> vbMaximized Then
            .ClassKey = HKEY_CURRENT_USER
            .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
            .ValueKey = "HostTop"
            If .Value <> vbNullString And .Value >= -20000 And .Value <= 50000 Then Top = .Value
            .ValueKey = "HostLeft"
            If .Value <> vbNullString And .Value >= -20000 And .Value <= 50000 Then Left = .Value
            .ValueKey = "HostWidth"
            If .Value <> vbNullString And .Value >= -20000 And .Value <= 50000 Then Width = .Value
            .ValueKey = "HostHeight"
            If .Value <> vbNullString And .Value >= 1000 And .Value <= 50000 Then Height = .Value
            .ValueKey = "HostWindowState"
            If .Value = vbNullString Then .Value = vbMaximized
            If .Value = "0" Or .Value = "2" Then
                WindowState = .Value
            End If
            
            .ValueKey = "HostHeight"
            If .Value = vbNullString Or Val(.Value) < 4000 Then
                Height = 4000
            Else
                Height = .Value
            End If
            
            .ValueKey = "HostWidth"
            If .Value = vbNullString Or Val(.Value) < 5000 Then
                Width = 5000
            Else
                Width = .Value
            End If
            
            .ValueKey = "HostLeft"
            If .Value > 0 Then Left = .Value
            .ValueKey = "HostTop"
            If .Value > 0 Then Top = .Value
       End If
    End With
    
    Call loadSortSettings
End Sub

Public Sub loadFilterSettings()
    'load filter settings
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
        .ValueKey = "Filter"
        If .Value = vbNullString Or .Value < 0 Or .Value > 6 Then
            cboFilter.ListIndex = 1
        Else
            cboFilter.ListIndex = Val(.Value)
        End If
    End With
End Sub

Private Function Host() 'form_load
On Error GoTo oops:
    Dim i As Integer
    
    If cboFilter.ListIndex = -1 Then cboFilter.ListIndex = 1
    
    strLastPicFileName = vbNullString
    blnDisconnectClick = True
    
    'add the sort arrow images to the Map ListView on the host form
    lvMaps.ColumnHeaderIcons = frmGH.ImageListSortIconIndicator
    lvMaps.AllowColumnReorder = False
    lvMaps.ListItems.Clear
    
    Me.Icon = frmGH.Icon
    Dim intNumMMPFiles As Integer
    Dim strMMPFolder As String
    
    strMMPFolder = strGTA2path & "\data"
      
    With lvMaps
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
    End With
    
    Const TXT_MMP_COLUMN = "MMP filename"
    
    intNumMMPFiles = 1
    
    'Find the column number with TXT_MMP_COLUMN in the title
    Dim intMMPcolumn As Integer
    For intMMPcolumn = 1 To lvMaps.ColumnHeaders.count
        If lvMaps.ColumnHeaders(intMMPcolumn).Text = TXT_MMP_COLUMN Then
            Exit For
        End If
    Next
    
    'Copy all items from frmGH.lvMMPlist to lvMaps
    Dim lstItem As ListItem
    Dim blnSkip As Boolean
    For i = 1 To frmGH.lvMMPlist.ListItems.count
        blnSkip = False
        With frmGH.lvMMPlist.ListItems(i)
            
            Select Case cboFilter.ListIndex
                Case 6 'All maps
                    blnSkip = False
                'Case 7 'Unique maps
                    'If strLastGMP = .SubItems(1) And strLastSTY = .SubItems(2) _
                    'And strLastSCR = .SubItems(3) Then blnSkip = True
                Case Else 'Maps with a specific PlayerCount
                    If Val(.SubItems(5)) <> cboFilter.ListIndex + 1 _
                        And .SubItems(5) <> vbNullString Then blnSkip = True
            End Select
            
            If blnSkip = False Then
                If InStr(LCase(.SubItems(4)), LCase(txtFind.Text)) > 0 Then
                    Set lstItem = lvMaps.ListItems.Add(, , .SubItems(4))
                    lstItem.SubItems(1) = .SubItems(1)
                    lstItem.SubItems(2) = .SubItems(2)
                    lstItem.SubItems(3) = .SubItems(3)
                    lstItem.SubItems(4) = frmGH.lvMMPlist.ListItems(i).Text
                    lstItem.SubItems(5) = .SubItems(5)
                    strLastGMP = .SubItems(1) & ".gmp"
                    strLastSTY = .SubItems(2) & ".sty"
                    strLastSCR = .SubItems(3) & ".scr"
                End If
            End If
        End With
    Next i
    Set lstItem = Nothing
    
    '''Show the number of maps in column header:
    With lvMaps
        .ColumnHeaders.Item(1) = "Maps (" & lvMaps.ListItems.count & ")"
    End With
    
    lvMaps.HideSelection = False
    
    i = 1
    
    Call loadSortSettings
    
    '''AutoSize the columns in this ListView:
    Call AutoSizeLV(lvMaps)
    'If lngDownloadSize = 0 Then DoEvents
    Call Form_Resize
    
    Exit Function
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error during hosting: " & strErrdesc & " - Line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :frmCreateGame load error: " & strErrdesc & " Line: " & strErrLine
End Function

Public Sub loadSortSettings()
  Dim i As Integer
  
  i = 1
  
  If lvMaps.ListItems.count = 0 Then Exit Sub
  
  With cr
        'lvMaps load sort settings
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
        .ValueKey = "lvMaps.SortOrder"
        lvMaps.SortOrder = .Value
        .ValueKey = "lvMaps.SortKey"
        lvMaps.SortKey = .Value
        Call SortColumn(frmCreateGame.lvMaps, .Value + 1)
        ShowListViewColumnHeaderSortIcon lvMaps
        
        'load last map and then find it in the list
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueKey = "MapDesc"
        If .Value <> vbNullString Then
            Dim itmX As ListItem
            Set itmX = lvMaps.FindItem(.Value, lvwText)
            If Not itmX Is Nothing Then
                i = itmX.Index
            End If
        End If
        
        'select it
        If lvMaps.ListItems.count >= i Then
            lvMaps.ListItems.Item(i).Selected = True
            lvMaps.SelectedItem.EnsureVisible
           Call lvMaps_ItemClick(lvMaps.ListItems.Item(i))
        End If
      
        .ValueKey = "chkHostPassword"
        chkHostPassword = Val(.Value)
        .ValueKey = "HostPassword"
        txtHostPassword = .Value
        txtHostPassword.Enabled = chkHostPassword 'enable if ticked
        
        If blnChkSync = True Then
            chkSync = vbChecked
        Else
            chkSync = vbUnchecked
        End If
        
    End With
End Sub

Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then cmdRefresh_Click
    If Shift = vbCtrlMask And KeyCode = vbKeyR Then
        KeyCode = 0
        cmdRefresh_Click
    End If
End Sub
