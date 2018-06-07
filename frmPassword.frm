VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game password?"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmPassword"
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

Dim strHostNick As String

Public Sub main(ByVal lstGames As String)
On Error Resume Next
    Me.Icon = frmGH.Icon
    strHostNick = lstGames
    frmPassword.Show vbModeless, frmGH
    '''FOCUS'''If txtPassword.Enabled = True Then Call giveFocus(txtPassword)
Exit Sub
End Sub

Private Sub cmdOK_Click()
On Error GoTo oops
    'strExecutableChecksum = calc_crc32(strGTA2path & TXT_GTA2EXE)
    'strMapChecksum = calc_crc32(strGTA2path & "data\" & strHostGMP)
    'strScriptChecksum = calc_crc32(strGTA2path & "data\" & strHostSCR)
    
    If txtPassword <> vbNullString Then
        send "NOTICE " & strHostNick & " :J" & strExecutableChecksum _
            & strMapChecksum & strScriptChecksum & strMMPfile & "/" & txtPassword
        frmPassword.Hide
    End If
Exit Sub

oops:
displaychat strChannel, vbRed, "FAIL!"
End Sub

Private Sub cmdCancel_Click()
    frmPassword.Hide
End Sub
