VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Game settings"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame fraLock 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtHostPassword 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
      Begin VB.CheckBox chkHostPassword 
         Caption         =   "&Lock Game"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblPassword 
         Caption         =   "&Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSettings"
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

Public Sub main()
On Error GoTo oops
    Me.Icon = frmGH.Icon
    Dim cr As New cRegistry
    frmSettings.Show vbModeless, frmGH
    If strPasswordProtectGame = "Yes" Then chkHostPassword.Value = vbChecked
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueKey = "HostPassword"
        txtHostPassword = .Value
    End With
   
Exit Sub

oops:
displaychat strChannel, vbRed, "Game settings fail"

End Sub

Private Sub chkHostPassword_Click()
    txtHostPassword.Enabled = chkHostPassword
    lblPassword.Enabled = chkHostPassword
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim cr As New cRegistry
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
Unload frmSettings
End Sub
