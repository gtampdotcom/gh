VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5775
   ClientLeft      =   2340
   ClientTop       =   2010
   ClientWidth     =   6210
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   3986.005
   ScaleMode       =   0  'User
   ScaleWidth      =   5831.513
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox rtbCredits 
      Height          =   3495
      Left            =   3720
      TabIndex        =   0
      Top             =   1680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   6165
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmAbout.frx":C8B7
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Respect"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      MouseIcon       =   "frmAbout.frx":C940
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "frmAbout"
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
Private Sub cmdOK_Click()
Unload Me
End Sub

Public Sub about()
    Show
    Dim i As Integer
    Dim strRespect As Variant
    Me.Icon = frmGH.Icon
    Me.Caption = frmGH.Caption

    strRespect = Array("[Code and design:", "Sektor", _
    " ", "[Donators:", "DAFE", "Kernel", "Wario5", "TommySprat", "DrSlony", "FrankCrank", "CubanPete", "Heri", _
    " ", "[Support:", "VERY-LAG-DUDE (discord)", "BenMillard (code)", "Elypter (code)", "Gustavob (testing)", "Razor (testing)", "Kamil (testing)", "irc.gtanet.com", _
    " ", "[Graphics:", "CubanPete", "famfamfam.com", "DMA Design", _
    " ", "[Disclaimer:", "The integrity of this product cannot be guaranteed for high voltage operation. The Zaibatsu Corporation reserves the right to change the specifications without notice. Conditions apply.")
    For i = 0 To UBound(strRespect)
        'Add strMsg to rtbHistory and apply color
        With rtbCredits
            .SelStart = Len(.Text)
            If InStr(strRespect(i), "[") Then
                .SelBold = True
                strRespect(i) = Replace(strRespect(i), "[", vbNullString)
            Else
                .SelBold = False
            End If
            .SelText = strRespect(i) & vbNewLine
        End With
    Next i
    rtbCredits.SelStart = 0
End Sub
