VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   7530
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   15060
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   15060
   Begin VB.Frame frameArray 
      BorderStyle     =   0  'None
      Caption         =   "Display"
      Height          =   2775
      Index           =   4
      Left            =   10560
      TabIndex        =   49
      Top             =   4440
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CheckBox chkMenu 
         Caption         =   "Hid&e &menu"
         Height          =   255
         Left            =   0
         TabIndex        =   60
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CheckBox chkPad 
         Caption         =   "Pad names"
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   1290
         Width           =   2655
      End
      Begin VB.CheckBox chkGameClear 
         Caption         =   "Show game removal messages"
         Height          =   255
         Left            =   0
         TabIndex        =   52
         Top             =   900
         Width           =   3015
      End
      Begin VB.CommandButton cmdURL 
         Caption         =   "&URL colour"
         Height          =   300
         Left            =   0
         TabIndex        =   59
         Top             =   2070
         Width           =   1215
      End
      Begin VB.CheckBox chkTimestamp 
         Caption         =   "T&imestamp messages"
         Height          =   255
         Left            =   0
         TabIndex        =   51
         Top             =   510
         Width           =   3015
      End
      Begin VB.CheckBox chkHighlight 
         Caption         =   "&Highlight alert words"
         Height          =   255
         Left            =   0
         TabIndex        =   50
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Frame frameArray 
      BorderStyle     =   0  'None
      Caption         =   "Startup"
      Height          =   2175
      Index           =   3
      Left            =   4320
      TabIndex        =   46
      Top             =   5040
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkConnectOnStartup 
         Caption         =   "Sign &In automatically"
         Height          =   255
         Left            =   0
         TabIndex        =   47
         Top             =   120
         Width           =   3015
      End
      Begin VB.CheckBox chkStartup 
         Caption         =   "&Run on startup"
         Height          =   255
         Left            =   0
         TabIndex        =   48
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Frame frameArray 
      BorderStyle     =   0  'None
      Caption         =   "System tray"
      Height          =   2415
      Index           =   2
      Left            =   120
      TabIndex        =   41
      Top             =   5040
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CheckBox chkMinToTray 
         Caption         =   "M&inimize to tray"
         Height          =   255
         Left            =   0
         TabIndex        =   45
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CheckBox chkCloseTray 
         Caption         =   "&Close to tray"
         Height          =   255
         Left            =   0
         TabIndex        =   44
         Top             =   840
         Width           =   3255
      End
      Begin VB.CheckBox chkStartTray 
         Caption         =   "Sta&rt in tray"
         Height          =   255
         Left            =   0
         TabIndex        =   43
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox chkTray 
         Caption         =   "A&lways in tray"
         Height          =   255
         Left            =   0
         TabIndex        =   42
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   58
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   57
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame frameArray 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   7095
      Begin VB.CheckBox chkVPN 
         Caption         =   "Force blank IP for VPN"
         Height          =   195
         Left            =   0
         TabIndex        =   61
         Top             =   2520
         Width           =   6975
      End
      Begin VB.Frame fraName 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         TabIndex        =   29
         Top             =   120
         Width           =   7095
         Begin VB.CheckBox chkAutoDownload 
            Caption         =   "Automatically &download and install maps"
            Height          =   195
            Left            =   0
            TabIndex        =   38
            Top             =   2040
            Width           =   6975
         End
         Begin VB.CheckBox chkShow 
            Caption         =   "&Show"
            Height          =   315
            Left            =   6240
            TabIndex        =   33
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtGTA2folder 
            Height          =   285
            Left            =   1320
            TabIndex        =   35
            Top             =   1560
            Width           =   5175
         End
         Begin VB.CommandButton cmdGTA2Folder 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6600
            TabIndex        =   37
            Top             =   1560
            Width           =   375
         End
         Begin VB.TextBox txtGTA2name 
            Height          =   285
            Left            =   1320
            TabIndex        =   30
            Top             =   120
            Width           =   5655
         End
         Begin VB.TextBox txtIRCPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   32
            Top             =   1080
            Width           =   4695
         End
         Begin VB.TextBox txtPreferedUsername 
            Height          =   285
            Left            =   1320
            TabIndex        =   31
            Top             =   600
            Width           =   5655
         End
         Begin VB.Label Label5 
            Caption         =   "&GTA2 name:"
            Height          =   375
            Left            =   0
            TabIndex        =   40
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lblFolder 
            Caption         =   "GTA2 &folder:"
            Height          =   255
            Left            =   0
            TabIndex        =   39
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "&Password:"
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "&IRC name:"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   600
            Width           =   1215
         End
      End
   End
   Begin VB.Frame frameArray 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   1
      Left            =   7800
      TabIndex        =   27
      ToolTipText     =   "Seperate alert words by spaces. Case insensitive."
      Top             =   480
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CheckBox chkSoundHosted 
         Height          =   255
         Left            =   5760
         TabIndex        =   56
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton cmdSoundHosted 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   25
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdPlaySoundHosted 
         Height          =   375
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdSoundLocation1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   5
         Top             =   450
         Width           =   375
      End
      Begin VB.CheckBox chkMuteAlertSound 
         Caption         =   "&Mute alert sounds while GTA2 is open"
         Height          =   195
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   4095
      End
      Begin VB.TextBox txtWordAlert 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         ToolTipText     =   "Seperate alert words by spaces. Case insensitive."
         Top             =   1440
         Width           =   4335
      End
      Begin VB.CheckBox chkSoundLocation1 
         Height          =   255
         Left            =   5760
         TabIndex        =   4
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdPlaySoundLocation1 
         Height          =   375
         Left            =   6600
         Picture         =   "frmOptions.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   450
         Width           =   375
      End
      Begin VB.CommandButton cmdPlaySoundJoin 
         Height          =   375
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2355
         Width           =   375
      End
      Begin VB.CommandButton cmdSoundJoin 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   23
         Top             =   2355
         Width           =   375
      End
      Begin VB.CheckBox chkSoundJoin 
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   2400
         Width           =   255
      End
      Begin VB.CommandButton cmdSoundPrivmsg 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   19
         Top             =   1875
         Width           =   375
      End
      Begin VB.CommandButton cmdSoundWordAlert 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   15
         Top             =   1395
         Width           =   375
      End
      Begin VB.CommandButton cmdSoundLocation2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   930
         Width           =   375
      End
      Begin VB.CommandButton cmdPlaySoundPrivmsg 
         Height          =   375
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1875
         Width           =   375
      End
      Begin VB.CommandButton cmdPlaySoundWordAlert 
         Height          =   375
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdPlaySoundLocation2 
         Height          =   375
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   930
         Width           =   375
      End
      Begin VB.CheckBox chkSoundPrivmsg 
         Height          =   255
         Left            =   5760
         TabIndex        =   18
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox chkSoundWordAlert 
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkSoundLocation2 
         Height          =   255
         Left            =   5760
         TabIndex        =   9
         Top             =   960
         Width           =   255
      End
      Begin MSComctlLib.ImageCombo imgcboLocation1 
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Top             =   450
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "imgcboLocation"
      End
      Begin MSComctlLib.ImageCombo imgcboLocation2 
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Top             =   930
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "imgcboLocation2"
      End
      Begin VB.Label lblSoundHosted 
         Caption         =   "Anyone &hosts a game:"
         Height          =   255
         Left            =   0
         TabIndex        =   55
         Top             =   2880
         Width           =   5655
      End
      Begin VB.Label lblSound 
         Caption         =   "Sound"
         Height          =   255
         Left            =   5640
         TabIndex        =   54
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "&Joining your game:"
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   2400
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "Pri&vate message:"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label lblNickAlert 
         Caption         =   "A&lert words:"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         ToolTipText     =   "Seperate alert words by spaces. Case insensitive."
         Top             =   1485
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Players from:"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Players from:"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   525
         Width           =   1335
      End
   End
   Begin MSComctlLib.TabStrip tabSettings 
      Height          =   4800
      Left            =   0
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8467
      MultiRow        =   -1  'True
      TabMinWidth     =   529
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Main"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Audio alerts"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Tray"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Startup"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Display"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
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
Dim strErrdesc As String
Dim strErrNum As String

Private Sub chkShow_Click()
    If chkShow.Value = vbChecked Then
        txtIRCPassword.PasswordChar = vbNullString
    Else
        txtIRCPassword.PasswordChar = "*"
    End If
End Sub

Private Sub cmdURL_Click()
    Dim strTemp As String
    strTemp = ColorView
    If strTemp <> vbNullString Then
        strLinkColor = strTemp
        frmGH.rtbTopic(1).Text = frmGH.rtbTopic(1).Text & vbNullString
    End If
End Sub

Private Sub form_load()

With cr
    frmOptions.Width = frameArray(0).Width + 500
    frmOptions.Height = frameArray(0).Height + 2000
    Dim frameTab As Frame
    For Each frameTab In frameArray
        frameTab.Left = frameArray(0).Left
        frameTab.Top = frameArray(0).Top
    Next
    
    tabSettings.Height = frmOptions.Height
    tabSettings.Width = frmOptions.Width

    'Load window size and position:
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
    .ValueKey = "SettingsTop"
    'Debug.Prin "frmOptions_Load()"; .Value;
    If .Value <> vbNullString And .Value >= -20000 And .Value <= 50000 Then Top = .Value
    .ValueKey = "SettingsLeft"
    'Debug.Prin .Value;
    If .Value <> vbNullString And .Value >= -20000 And .Value <= 50000 Then Left = .Value
    .ValueKey = "SettingsWidth"
    'Debug.Prin .Value;
    'If .Value <> vbNullString And .Value >= -20000 And .Value <= 50000 Then Width = .Value
    '.ValueKey = "SettingsHeight"
    'Debug.Prin .Value;
    'If .Value <> vbNullString And .Value >= 1000 And .Value <= 50000 Then Height = .Value
    .ValueKey = "SettingsWindowState"
    'Debug.Prin .Value
    If .Value = vbNullString Then .Value = vbMaximized
    If .Value = "0" Or .Value = "2" Then
        WindowState = .Value
    End If
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

With cr
    'Save window size and position:
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
    .ValueKey = "SettingsWindowState"
    .Value = WindowState
    If .Value <> "2" Then
        .ValueKey = "SettingsHeight"
        .ValueType = REG_SZ
        .Value = Height

        .ValueKey = "SettingsWidth"
        .ValueType = REG_SZ
        .Value = Width

        .ValueKey = "SettingsLeft"
        .ValueType = REG_SZ
        .Value = Left

        .ValueKey = "SettingsTop"
        .ValueType = REG_SZ
        .Value = Top
    End If
    
    '''FOCUS'''Call frmGH.giveChatFocus
End With

End Sub

Private Sub imgcboLocation1_Keypress(key As Integer)
    Call keypressCountrySearch(imgcboLocation1, key)
End Sub

Private Sub imgcboLocation2_Keypress(key As Integer)
    Call keypressCountrySearch(imgcboLocation2, key)
End Sub

'when a key is pushed, search the combobox for an item starting with that letter
Private Sub keypressCountrySearch(imagecombobox, key As Integer)
    Dim i As Integer
    Dim intTotalCount As Integer
    
    If imagecombobox.SelectedItem.Index = imagecombobox.ComboItems.count Then
        imagecombobox.SelectedItem = imagecombobox.ComboItems.Item(1)
    End If
    
    For i = imagecombobox.SelectedItem.Index + 1 To imagecombobox.ComboItems.count
        If Left$(imagecombobox.ComboItems.Item(i).Text, 1) = UCase$(Chr$(key)) Then
            imagecombobox.SelectedItem = imagecombobox.ComboItems.Item(i)
            Exit For
        Else
            If i = imagecombobox.ComboItems.count Then
                imagecombobox.SelectedItem = imagecombobox.ComboItems.Item(1)
                i = 1
            End If
        End If
        intTotalCount = intTotalCount + 1
        If intTotalCount > imagecombobox.ComboItems.count Then Exit Sub
    Next
End Sub

Private Sub cmdCancel_Click()
On Error GoTo oops:
    '''FOCUS'''Call frmGH.giveChatFocus
    Unload Me
    Exit Sub
oops:
    strErrdesc = Err.Description
    displaychat strDestTab, vbRed, "error unloading form: " & strErrdesc
End Sub

Private Sub cmdGTA2Folder_Click()
    Dim strTemp As String
    strTemp = BrowseFile(strGTA2path & TXT_GTA2EXE)
    If Len(strTemp) > Len(TXT_GTA2EXE) And Right$(strTemp, Len(TXT_GTA2EXE)) = TXT_GTA2EXE Then
        strGTA2path = Left$(strTemp, Len(strTemp) - Len(TXT_GTA2EXE))
        txtGTA2folder = strGTA2path
        frmOptions.Show
    End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo oops:
    'gta2 game hunter settings
    strPreferedNick = txtPreferedUsername
    If blnConnected = True Then
        If strNick <> strPreferedNick Then
            intNickservWaitTime = 0
            send "NICK " & strPreferedNick
        End If
        
        If txtIRCPassword <> vbNullString And txtIRCPassword <> strPassword Then send "NS IDENTIFY " & txtIRCPassword
    End If
    strGTA2path = txtGTA2folder
    'intServerNum = cboServer.ListIndex
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\DMA Design Ltd\GTA2\Network"
        .ValueKey = "PlayerName"
        .ValueType = REG_SZ
        .Value = txtGTA2name
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        
        '.ValueKey = "ServerNum"
        '.ValueType = REG_SZ
        '.Value = intServerNum
        
        .ValueKey = "Username"
        If txtPreferedUsername = vbNullString Then
            .DeleteValue
        Else
            .Value = txtPreferedUsername
        End If
        
        .ValueKey = "PlayerName"
        If txtGTA2name = vbNullString Then
            .DeleteValue
        Else
            .Value = txtGTA2name
        End If
        
        strPassword = txtIRCPassword
        .ValueKey = "EncodedPassword"
        If txtIRCPassword = vbNullString Then
            If .Value <> vbNullString Then .DeleteValue
        Else
            .Value = encode(txtIRCPassword)
        End If
        
'         .ValueKey = "Password"
'        strPassword = txtIRCPassword
'        If txtIRCPassword = vbNullString Then
'            If .Value <> vbNullString Then .DeleteValue
'        Else
'            .Value = strPassword
'        End If
       
        .ValueKey = "GTA2Folder"
        If strGTA2path = vbNullString Then
            .DeleteValue
        Else
            .Value = strGTA2path
        End If
        
        .ValueKey = "chkMuteAlertSound"
        .Value = chkMuteAlertSound
        
        .ValueKey = "chkAutoDownload"
        .Value = chkAutoDownload
        
        .ValueKey = "chkVPN"
        .Value = chkVPN
        
        'check all the current sound filenames, if they aren't blank then save their name in registry
        If strSoundLocation1 <> vbNullString Then
            .ValueKey = "SoundLocation1"
            .Value = strSoundLocation1
        End If
        
        If strSoundLocation2 <> vbNullString Then
            .ValueKey = "SoundLocation2"
            .Value = strSoundLocation2
        End If
        
        .ValueKey = "SoundWordAlert"
        If strSoundWordAlert = vbNullString Then
            .DeleteValue
        Else
            .Value = strSoundWordAlert
        End If
        
        .ValueKey = "SoundHosted"
        If strSoundHosted = vbNullString Then
            .DeleteValue
        Else
            .Value = strSoundHosted
        End If
        
        .ValueKey = "txtWordAlert"
        If txtWordAlert = vbNullString Then
            .DeleteValue
        Else
            .Value = txtWordAlert
        End If
        
        If strSoundPrivmsg <> vbNullString Then
            .ValueKey = "SoundPrivmsg"
            .Value = strSoundPrivmsg
        End If
        If strSoundJoin <> vbNullString Then
            .ValueKey = "SoundJoin"
            .Value = strSoundJoin
        End If
        If strSoundHosted <> vbNullString Then
            .ValueKey = "SoundHosted"
            .Value = strSoundHosted
        End If
        
        .ValueKey = "Location1"
        strLocation1 = Right$(imgcboLocation1.SelectedItem.Text, 2)
        .Value = strLocation1
        
        .ValueKey = "Location2"
        strLocation2 = Right$(imgcboLocation2.SelectedItem.Text, 2)
        
        .Value = strLocation2
                
        .ValueKey = "chkSoundLocation1"
        .Value = chkSoundLocation1
        
        .ValueKey = "chkSoundLocation2"
        .Value = chkSoundLocation2
        
        .ValueKey = "chkSoundWordAlert"
        .Value = chkSoundWordAlert
        
        .ValueKey = "chkSoundPrivmsg"
        .Value = chkSoundPrivmsg
        
        .ValueKey = "chkSoundJoin"
        .Value = chkSoundJoin
        
        .ValueKey = "chkSoundHosted"
        .Value = chkSoundHosted
        
        .ValueKey = "ConnectOnStartup"
        .Value = chkConnectOnStartup
        
        .ValueKey = "chkTray"
        .Value = chkTray
        
        .ValueKey = "chkCloseTray"
        .Value = chkCloseTray
        
        .ValueKey = "chkStartTray"
        .Value = chkStartTray
        
        .ValueKey = "chkHighlight"
        .Value = chkHighlight
        
        .ValueKey = "chkGameClear"
        .Value = chkGameClear
        
        .ValueKey = "chkPad"
        .Value = chkPad
        
        .ValueKey = "LinkColor"
        .Value = strLinkColor
        
        .ValueKey = "chkTimestamp"
        frmGH.mnuViewTimestamp.Checked = chkTimestamp
        .ValueType = REG_DWORD
        .Value = chkTimestamp
        blnchkTime = .Value
        
        .ValueKey = "chkMinToTray"
        .Value = chkMinToTray
        blnchkMinToTray = .Value
        
        .ValueKey = "chkMenu"
        .Value = chkMenu
        
        blnHighlight = chkHighlight
        blnchkGameClear = chkGameClear
        blnchkPad = chkPad
        frmGH.mnuViewHighlight.Checked = blnHighlight
        blnchkMuteAlertSound = chkMuteAlertSound
        blnchkAutoDownload = chkAutoDownload
        blnchkVPN = chkVPN
        strTxtWordAlert = Trim(txtWordAlert)
        Call AlertWords
        blnchkSoundLocation1 = chkSoundLocation1
        blnchkSoundLocation2 = chkSoundLocation2
        blnchkSoundWordAlert = chkSoundWordAlert
        blnchkSoundPrivmsg = chkSoundPrivmsg
        blnchkSoundJoin = chkSoundJoin
        blnchkSoundHosted = chkSoundHosted
        blnchkConnectOnStartup = chkConnectOnStartup
        blnchkTray = chkTray
        If blnchkTray = True Then Call Systray
        blnchkCloseTray = chkCloseTray
        blnchkStartTray = chkStartTray
        blnchkStartup = chkStartup
        
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
        .ValueKey = "gta2gh.exe"
        .ValueType = REG_SZ
        If chkStartup.Value = vbChecked Then
            .Value = App.Path & "\" & App.EXEName
        Else
            .DeleteValue
        End If
                
        Dim b As Boolean
        b = chkMenu.Value
        Dim i  As Integer
        frmGH.mnuFile.Visible = Not (b)
        frmGH.mnuEdit.Visible = Not (b)
        frmGH.mnuView.Visible = Not (b)
        frmGH.mnuTools.Visible = Not (b)
        frmGH.mnuHelp.Visible = Not (b)
        frmGH.Visible = False
        frmGH.Visible = True
        
    End With
    '''FOCUS'''Call frmGH.giveChatFocus
    Unload Me
    Exit Sub
oops:
strErrdesc = Err.Description
strErrNum = Erl
displaychat strChannel, vbRed, "If this is a registry related error then try logging in to Windows as administrator: " & strErrdesc
send "PRIVMSG " & gta2ghbot & " :Settings OK error: " & strErrdesc
End Sub


Private Sub cmdPlaySoundLocation1_Click()
PlaySound strSoundLocation1, ByVal 0&, SND_ASYNC
End Sub

Private Sub cmdPlaySoundLocation2_Click()
PlaySound strSoundLocation2, ByVal 0&, SND_ASYNC
End Sub

Private Sub cmdPlaySoundWordAlert_Click()
PlaySound strSoundWordAlert, ByVal 0&, SND_ASYNC
End Sub

Private Sub cmdPlaySoundPrivmsg_Click()
PlaySound strSoundPrivmsg, ByVal 0&, SND_ASYNC
End Sub

Private Sub cmdPlaySoundJoin_Click()
PlaySound strSoundJoin, ByVal 0&, SND_ASYNC
End Sub

Private Sub cmdPlaySoundHosted_Click()
PlaySound strSoundHosted, ByVal 0&, SND_ASYNC
End Sub

Private Sub cmdSoundLocation1_Click()
Dim strSound As String
strSound = BrowseFile(strSoundLocation1, True)
If strSound <> vbNullString Then strSoundLocation1 = strSound
Show
End Sub

Private Sub cmdSoundLocation2_Click()
Dim strSound As String
strSound = BrowseFile(strSoundLocation2, True)
If strSound <> vbNullString Then strSoundLocation2 = strSound
Show
End Sub

Private Sub cmdSoundWordAlert_Click()
Dim strSound As String
strSound = BrowseFile(strSoundWordAlert, True)
If strSound <> vbNullString Then strSoundWordAlert = strSound
Show
End Sub

Private Sub cmdSoundPrivmsg_Click()
Dim strSound As String
strSound = BrowseFile(strSoundPrivmsg, True)
If strSound <> vbNullString Then strSoundPrivmsg = strSound
Show
End Sub

Private Sub cmdSoundJoin_Click()
Dim strSound As String
strSound = BrowseFile(strSoundJoin, True)
If strSound <> vbNullString Then strSoundJoin = strSound
Show
End Sub

Private Sub cmdSoundHosted_Click()
Dim strSound As String
strSound = BrowseFile(strSoundHosted, True)
If strSound <> vbNullString Then strSoundHosted = strSound
Show
End Sub

Public Sub loadSettings() 'form_load()
On Error GoTo oops

Dim i As Integer

''Fill port list with ports
'For i = 2301 To 2400
'    lstPort.AddItem (i)
'Next i
'
'lstPort.ListIndex = 0

cmdPlaySoundLocation2.Picture = cmdPlaySoundLocation1.Picture
cmdPlaySoundWordAlert.Picture = cmdPlaySoundLocation1.Picture
cmdPlaySoundPrivmsg.Picture = cmdPlaySoundLocation1.Picture
cmdPlaySoundJoin.Picture = cmdPlaySoundLocation1.Picture
cmdPlaySoundHosted.Picture = cmdPlaySoundLocation1.Picture
    
Me.Icon = frmGH.Icon

'For i = 0 To UBound(strServer)
'    If strServer(i) <> vbNullString Then cboServer.list(i) = strServer(i)
'Next i

With cr
    .ClassKey = HKEY_CURRENT_USER
    
    .SectionKey = "Software\DMA Design Ltd\GTA2\Network"
    .ValueKey = "PlayerName"
    txtGTA2name = .Value
    .SectionKey = "SOFTWARE\GTA2 Game Hunter"
    
'    .ValueKey = "ServerNum"
'
'    If IsNumeric(.Value) Then
'        intServerNum = .Value
'    Else
'        intServerNum = 0
'    End If
'
'    'If intServerNum < 0 Or intServerNum + 1 > cboServer.ListCount Then intServerNum = 0
'
'    cboServer.ListIndex = intServerNum
    
    .ValueKey = "Username"
    If .Value = vbNullString Then
        txtPreferedUsername = vbNullString 'txtGTA2name
    Else
        txtPreferedUsername = .Value
    End If
  
    .ValueKey = "EncodedPassword"
    If .Value = vbNullString Then
        .ValueKey = "Password"
        txtIRCPassword = .Value
        strPassword = txtIRCPassword
    Else
        txtIRCPassword = encode(.Value)
        strPassword = txtIRCPassword
    End If
    
    .ValueKey = "chkMuteAlertSound"
    chkMuteAlertSound = .Value
    
    .ValueKey = "chkAutoDownload"
    If .Value = vbNullString Then
        chkAutoDownload.Value = vbChecked
        .ValueType = REG_DWORD
        .Value = 1
    End If
    
    chkAutoDownload = .Value
    
    .ValueKey = "chkVPN"
    If .Value = vbNullString Then
        chkVPN.Value = vbChecked
        .ValueType = REG_DWORD
        .Value = 1
    End If
    
    chkVPN = .Value
    
    .ValueKey = "SoundLocation1"
    strSoundLocation1 = .Value
    .ValueKey = "SoundLocation2"
    strSoundLocation2 = .Value
    .ValueKey = "SoundWordAlert"
    strSoundWordAlert = .Value
    
    .ValueKey = "txtWordAlert"
    If .Value = vbNullString Then
        txtWordAlert = txtPreferedUsername
    Else
        txtWordAlert = .Value
    End If
    
    .ValueKey = "SoundPrivmsg"
    strSoundPrivmsg = .Value
    .ValueKey = "SoundJoin"
    strSoundJoin = .Value
    .ValueKey = "SoundHosted"
    strSoundHosted = .Value
    
    'Try to find GTA2 folder:
    Call setGTA2path
    
'    If Exists(strGTA2path & TXT_GTA2EXE) = False Then
'
'        If Exists(PROGRAM_FILES & "\rockstar games\gta2\" & TXT_GTA2EXE) = True Then
'            strGTA2path = PROGRAM_FILES & "\rockstar games\gta2\"
'        Else
'            If Exists(PROGRAM_FILES & "\gta2\" & TXT_GTA2EXE) = True Then
'                strGTA2path = PROGRAM_FILES & "\gta2\"
'            Else
'                '[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\gta2.exe]
'                .ClassKey = HKEY_LOCAL_MACHINE
'                .SectionKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\gta2.exe\"
'                .ValueKey = vbNullString
'                If Len(.Value) > 8 Then
'                    strGTA2path = Mid$(.Value, 1, Len(.Value) - 8)
'                End If
'
'                'change classkey and sectionkey back to the location of GH settings
'                .ClassKey = HKEY_CURRENT_USER
'                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
'            End If
'        End If
'    End If
    
    If strGTA2path = "\" Or strGTA2path = "/" Or strGTA2path = "\\" Then strGTA2path = vbNullString

    If strGTA2path <> vbNullString Then
        .ValueKey = "GTA2folder"
        .ValueType = REG_SZ
        .Value = strGTA2path
    End If
    
    txtGTA2folder = strGTA2path
    
    .ValueKey = "Country"
    If .Value = vbNullString Then
        'do nothing
    Else
        'some country code is saved in registry
        
        For i = 0 To UBound(strCountries)
            If Right$(strCountries(i), 2) = .Value Then
                 strCountryCode = Right$(strCountries(i), 2)
                 intCountryIndex = i
                 Exit For
            End If
        Next
    End If
        
    .ValueKey = "chkSoundWordAlert"
    If .Value = vbNullString Then
        chkSoundWordAlert = vbChecked
    Else
        chkSoundWordAlert = .Value
    End If
    
    .ValueKey = "chkSoundPrivmsg"
    If .Value = vbNullString Then
        chkSoundPrivmsg = vbChecked
    Else
        chkSoundPrivmsg = .Value
    End If
    
    .ValueKey = "chkSoundJoin"
    If .Value = vbNullString Then
        chkSoundJoin = vbChecked
    Else
        chkSoundJoin = .Value
    End If
    
    .ValueKey = "chkSoundHosted"
    If .Value = vbNullString Then
        chkSoundHosted = vbChecked
    Else
        chkSoundHosted = .Value
    End If
    
    .ValueKey = "ConnectOnStartup"
    chkConnectOnStartup = .Value
    
    .ValueKey = "Hide"
    blnHidden = .Value
    
    .ValueKey = "chkTray"
    chkTray = .Value
    
    .ValueKey = "chkCloseTray"
    chkCloseTray = .Value
    
    .ValueKey = "chkStartTray"
    chkStartTray = .Value
    
    .ValueKey = "chkTimestamp"
    frmGH.mnuViewTimestamp.Checked = .Value
    chkTimestamp = .Value
    
    .ValueKey = "chkMinToTray"
    chkMinToTray = .Value
    
    .ValueKey = "chkGameClear"
    chkGameClear = .Value
    
    .ValueKey = "chkPad"
    If .Value = vbNullString Then
        chkPad.Value = vbChecked
    Else
        chkPad.Value = Val(.Value)
    End If
    
    'Hide tool menu
    .ValueKey = "chkMenu"
    chkMenu = .Value
    Dim b As Boolean
    b = chkMenu.Value
    frmGH.mnuFile.Visible = Not (b)
    frmGH.mnuEdit.Visible = Not (b)
    frmGH.mnuView.Visible = Not (b)
    frmGH.mnuTools.Visible = Not (b)
    frmGH.mnuHelp.Visible = Not (b)
    frmGH.Visible = False
    If lngMaster = 0 Then frmGH.Visible = True
    
    blnchkMuteAlertSound = chkMuteAlertSound
    blnchkAutoDownload = chkAutoDownload
    blnchkVPN = chkVPN
    strTxtWordAlert = Trim(txtWordAlert)
    Call AlertWords
    blnchkSoundLocation1 = chkSoundLocation1
    blnchkSoundLocation2 = chkSoundLocation2
    blnchkSoundWordAlert = chkSoundWordAlert
    blnchkSoundPrivmsg = chkSoundPrivmsg
    blnchkSoundJoin = chkSoundJoin
    blnchkSoundHosted = chkSoundHosted
    blnchkConnectOnStartup = chkConnectOnStartup
    blnchkTray = chkTray
    blnchkCloseTray = chkCloseTray
    blnchkStartTray = chkStartTray
    blnchkMinToTray = chkMinToTray
    
    If blnchkTray = True Then Call Systray
    
    .ValueKey = "chkHighlight"
    If .Value = vbNullString Then
        chkHighlight.Value = vbChecked
    Else
        chkHighlight.Value = Val(.Value)
    End If
    
    blnHighlight = chkHighlight
    frmGH.mnuViewHighlight.Checked = blnHighlight
    
    .ValueKey = "LinkColor"
    If .Value <> vbNullString Then strLinkColor = .Value
    
    blnchkGameClear = chkGameClear
    blnchkPad = chkPad
       
    Set imgcboLocation1.ImageList = frmGH.ImageList1
    Set imgcboLocation2.ImageList = frmGH.ImageList1

    For i = 1 To frmGH.ImageList1.ListImages.count
        imgcboLocation1.ComboItems.Add i, frmGH.ImageList1.ListImages(i).key, strCountries(i - 1), i, i
        imgcboLocation2.ComboItems.Add i, frmGH.ImageList1.ListImages(i).key, strCountries(i - 1), i, i
    Next i
      
    imgcboLocation1.Locked = True
    imgcboLocation2.Locked = True
    Set imgcboLocation1.SelectedItem = imgcboLocation1.GetFirstVisible
    Set imgcboLocation2.SelectedItem = imgcboLocation2.GetFirstVisible
    
    .ValueKey = "Location1"
    If .Value = vbNullString Then
        'imgcboLocation1.ComboItems.Item(imgCboCountry.SelectedItem.Index).Selected = True
    Else
        For i = 0 To UBound(strCountries)
            If Right$(strCountries(i), 2) = .Value Then
                imgcboLocation1.ComboItems.Item(i + 1).Selected = True
                strLocation1 = .Value
                Exit For
            End If
        Next
    End If
  
    .ValueKey = "Location2"
    For i = 0 To UBound(strCountries)
        If Right$(strCountries(i), 2) = .Value Then
            imgcboLocation2.ComboItems.Item(i + 1).Selected = True
            strLocation2 = .Value
            Exit For
        End If
    Next
          
    .ValueKey = "chkSoundLocation1"
    If .Value = vbNullString Then
        chkSoundLocation1 = vbChecked
    Else
        chkSoundLocation1 = .Value
    End If
    
    .ValueKey = "chkSoundLocation2"
    If .Value = vbNullString Then
        chkSoundLocation2 = vbChecked
    Else
        chkSoundLocation2 = .Value
    End If
    
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Run"
    .ValueKey = "gta2gh.exe"
    If .Value <> vbNullString Then chkStartup.Value = vbChecked
    blnchkStartup = chkStartup
    
End With
Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    strErrNum = Err.Number
    displaychat strDestTab, vbRed, "error with settings form: " & strErrdesc & " Line number:" & strErrLine & " Error number:" & strErrNum & " i = " & i
End Sub

Public Sub AlertWords()

On Error GoTo oops:
Dim i As Integer
Dim j As Integer

For i = 0 To UBound(strAlertWords)
    strAlertWords(i) = vbNullString
Next i

For i = 1 To Len(strTxtWordAlert)
    If Mid$(strTxtWordAlert, i, 1) <> " " Then strAlertWords(j) = strAlertWords(j) & Mid$(strTxtWordAlert, i, 1)
    If Mid$(strTxtWordAlert, i, 1) = " " Then j = j + 1
    If j = UBound(strAlertWords) Then Exit For
Next i
Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    strErrNum = Err.Number
    displaychat strDestTab, vbRed, "alertwords error: " & strErrdesc & " Line number:" & strErrLine & " Error number:" & strErrNum

End Sub

Private Sub tabSettings_Click()

Dim frameTab As Frame
For Each frameTab In frameArray
    If tabSettings.SelectedItem.Index <> frameTab.Index + 1 Then
        frameTab.Visible = False
    Else
        frameTab.Visible = True
    End If
Next

End Sub
