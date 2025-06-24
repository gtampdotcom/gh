VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGH 
   AutoRedraw      =   -1  'True
   Caption         =   "GH"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   10650
   DrawStyle       =   5  'Transparent
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmGH.frx":37C2
   ScaleHeight     =   5055
   ScaleWidth      =   10650
   Visible         =   0   'False
   Begin VB.Timer timStamp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   4320
   End
   Begin VB.PictureBox picGH 
      Height          =   375
      Left            =   5160
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   22
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSanta 
      Height          =   615
      Left            =   9720
      Picture         =   "frmGH.frx":3DEC
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picHalloween 
      Height          =   375
      Left            =   9840
      Picture         =   "frmGH.frx":4AB6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSlave 
      Height          =   375
      Left            =   8160
      TabIndex        =   19
      Text            =   "txtSlave"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picJoke 
      Height          =   375
      Left            =   9840
      Picture         =   "frmGH.frx":2C026
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdToolbar 
      Caption         =   "&Manager"
      Height          =   360
      Index           =   3
      Left            =   5280
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdX 
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   210
      Left            =   7680
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3480
      Width           =   375
   End
   Begin VB.PictureBox picTray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picHead 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Picture         =   "frmGH.frx":2C650
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdToolbar 
      Caption         =   "Cancel &Download"
      Height          =   360
      Index           =   4
      Left            =   6600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtNoGames 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Created games will be shown here"
      Top             =   720
      Visible         =   0   'False
      Width           =   2652
   End
   Begin MSWinsockLib.Winsock sckURL 
      Left            =   1200
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin MSComctlLib.ListView lvMMPlist 
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin RichTextLib.RichTextBox rtbTopic 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "#gta2gh"
      Top             =   1080
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"frmGH.frx":2CBEE
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
   Begin MSComctlLib.TabStrip tabIRC 
      Height          =   360
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   635
      Placement       =   1
      TabMinWidth     =   529
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "#gta2gh"
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
   Begin RichTextLib.RichTextBox rtbChatbox 
      Height          =   390
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Tag             =   "#gta2gh"
      Top             =   2880
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   688
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"frmGH.frx":2CC7C
   End
   Begin VB.Timer timUpdateMap 
      Interval        =   2000
      Left            =   1200
      Top             =   4440
   End
   Begin RichTextLib.RichTextBox rtbHistory 
      Height          =   1455
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Tag             =   "#gta2gh"
      Top             =   1440
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmGH.frx":2CD05
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdToolbar 
      Caption         =   "O&ptions"
      Height          =   360
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdToolbar 
      Caption         =   "Sign &In"
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sockIRC 
      Left            =   720
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdToolbar 
      Caption         =   "Create &Game"
      Height          =   360
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdToolbar 
      Caption         =   "Sign &Out"
      Enabled         =   0   'False
      Height          =   360
      Index           =   10
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer timStatus 
      Interval        =   1000
      Left            =   120
      Top             =   4440
   End
   Begin VB.Timer timTimeout 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1680
      Top             =   4440
   End
   Begin MSComctlLib.ImageList ImageListSortIconIndicator 
      Left            =   2280
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   7
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2CD7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2CE4E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   253
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2CF20
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2D017
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2D1A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2D32D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2D4BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2D649
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2D7D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2D961
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2DA71
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2DBFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2DCF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2DE7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2E004
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2E18B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2E314
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2E4AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2E62F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2E7BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2E938
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2EAC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2EC4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2EDD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2EF59
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2F0DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2F268
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2F3F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2F57F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2F6FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2F88F
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2FA17
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2FB9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2FD21
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":2FEB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":30040
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":301D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":30362
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":304E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":30668
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":307FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":30985
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":30B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":30CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":30E31
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":30FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3114A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":312D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3145E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":315E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3176F
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":318FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":31A7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":31BFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":31D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":31E8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":32015
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":32198
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":32324
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":324AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":32637
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":327C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":32946
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":32ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":32C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":32DED
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":32F79
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":330FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":33285
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3340C
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":33593
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":33718
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3389D
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":33A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":33BAD
            Key             =   ""
            Object.Tag             =   "EU"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":33C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":33E04
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":33F92
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":34120
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":342AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":34439
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":34530
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":346BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":34844
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":349CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":34B67
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":34CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":34E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":35007
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3518D
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3530F
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3549D
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":35629
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":357B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3593D
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":35ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":35C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":35DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":35F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":360EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":36271
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":363FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":36581
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":36717
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":368A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":36A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":36BB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":36D45
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":36ECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":37052
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":371DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":37361
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":374F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3767A
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":37771
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":378FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":379F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":37B7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":37D05
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":37E89
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3801B
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3819F
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3832E
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":384B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":38655
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":387DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3896C
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":38AF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":38C7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":38E07
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":38F8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3910E
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":39294
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":39421
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":395A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":39733
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":398C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":39A59
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":39BE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":39D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":39F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3A08B
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3A218
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3A3A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3A53D
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3A6D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3A854
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3A9EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3AB78
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3AD0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3AE90
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3B01B
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3B1A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3B332
            Key             =   ""
         EndProperty
         BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3B6FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3B88D
            Key             =   ""
         EndProperty
         BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3BA17
            Key             =   ""
         EndProperty
         BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3BBA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3BD2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3BEB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3C001
            Key             =   ""
         EndProperty
         BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3C185
            Key             =   ""
         EndProperty
         BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3C311
            Key             =   ""
         EndProperty
         BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3C408
            Key             =   ""
         EndProperty
         BeginProperty ListImage163 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3C590
            Key             =   ""
         EndProperty
         BeginProperty ListImage164 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3C71D
            Key             =   ""
         EndProperty
         BeginProperty ListImage165 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3C8A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage166 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3CA31
            Key             =   ""
         EndProperty
         BeginProperty ListImage167 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3CBC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage168 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3CD4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage169 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3CEE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage170 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3D06C
            Key             =   ""
         EndProperty
         BeginProperty ListImage171 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3D200
            Key             =   ""
         EndProperty
         BeginProperty ListImage172 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3D388
            Key             =   ""
         EndProperty
         BeginProperty ListImage173 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3D51D
            Key             =   ""
         EndProperty
         BeginProperty ListImage174 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3D6AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage175 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3D831
            Key             =   ""
         EndProperty
         BeginProperty ListImage176 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3D9BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage177 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3DB40
            Key             =   ""
         EndProperty
         BeginProperty ListImage178 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3DCC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage179 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3DE4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage180 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3DFD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage181 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3E15C
            Key             =   ""
         EndProperty
         BeginProperty ListImage182 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3E2E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage183 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3E46D
            Key             =   ""
         EndProperty
         BeginProperty ListImage184 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3E5FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage185 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3E782
            Key             =   ""
         EndProperty
         BeginProperty ListImage186 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3E90C
            Key             =   ""
         EndProperty
         BeginProperty ListImage187 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3EA93
            Key             =   ""
         EndProperty
         BeginProperty ListImage188 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3EC18
            Key             =   ""
         EndProperty
         BeginProperty ListImage189 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3ED9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage190 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3EF24
            Key             =   ""
         EndProperty
         BeginProperty ListImage191 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3F0B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage192 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3F1D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage193 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3F35F
            Key             =   ""
         EndProperty
         BeginProperty ListImage194 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3F4E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage195 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3F673
            Key             =   ""
         EndProperty
         BeginProperty ListImage196 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3F7FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage197 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3F98C
            Key             =   ""
         EndProperty
         BeginProperty ListImage198 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3FB14
            Key             =   ""
         EndProperty
         BeginProperty ListImage199 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3FC1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage200 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3FD9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage201 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":3FF26
            Key             =   ""
         EndProperty
         BeginProperty ListImage202 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":400AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage203 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":40234
            Key             =   ""
         EndProperty
         BeginProperty ListImage204 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":403B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage205 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":40543
            Key             =   ""
         EndProperty
         BeginProperty ListImage206 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":406D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage207 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4085E
            Key             =   ""
         EndProperty
         BeginProperty ListImage208 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":409E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage209 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":40B77
            Key             =   ""
         EndProperty
         BeginProperty ListImage210 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":40D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage211 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":40E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage212 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":41017
            Key             =   ""
         EndProperty
         BeginProperty ListImage213 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4119C
            Key             =   ""
         EndProperty
         BeginProperty ListImage214 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":41330
            Key             =   ""
         EndProperty
         BeginProperty ListImage215 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":414B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage216 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":41642
            Key             =   ""
         EndProperty
         BeginProperty ListImage217 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":417AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage218 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4192F
            Key             =   ""
         EndProperty
         BeginProperty ListImage219 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":41ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage220 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":41C3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage221 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":41DC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage222 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":41F4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage223 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":420D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage224 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4225B
            Key             =   ""
         EndProperty
         BeginProperty ListImage225 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":423EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage226 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":42576
            Key             =   ""
         EndProperty
         BeginProperty ListImage227 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4270B
            Key             =   ""
         EndProperty
         BeginProperty ListImage228 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4289E
            Key             =   ""
         EndProperty
         BeginProperty ListImage229 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":42A2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage230 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":42BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage231 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":42D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage232 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":42EC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage233 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4304A
            Key             =   ""
         EndProperty
         BeginProperty ListImage234 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":431CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage235 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":43353
            Key             =   ""
         EndProperty
         BeginProperty ListImage236 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":43473
            Key             =   ""
         EndProperty
         BeginProperty ListImage237 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":435FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage238 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4378D
            Key             =   ""
         EndProperty
         BeginProperty ListImage239 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4391E
            Key             =   ""
         EndProperty
         BeginProperty ListImage240 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":43AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage241 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":43C2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage242 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":43DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage243 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":43F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage244 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":440D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage245 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4425E
            Key             =   ""
         EndProperty
         BeginProperty ListImage246 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":443F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage247 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":44587
            Key             =   ""
         EndProperty
         BeginProperty ListImage248 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4470A
            Key             =   ""
         EndProperty
         BeginProperty ListImage249 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":4488A
            Key             =   ""
         EndProperty
         BeginProperty ListImage250 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":44A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage251 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":44B95
            Key             =   ""
         EndProperty
         BeginProperty ListImage252 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":44D97
            Key             =   ""
         EndProperty
         BeginProperty ListImage253 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGH.frx":44FF9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvGames 
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageListSortIconIndicator"
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "STRING"
         Text            =   "Games (0)"
         Object.Width           =   2091
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "STRING"
         Text            =   "Pass"
         Object.Width           =   926
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "STRING"
         Text            =   "CC"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "STRING"
         Text            =   "Map"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "STRING"
         Text            =   "GH"
         Object.Width           =   741
      EndProperty
   End
   Begin MSComctlLib.ListView lvPlayers 
      Height          =   2895
      Index           =   1
      Left            =   5520
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "#gta2gh"
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageListSortIconIndicator"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "STRING"
         Text            =   "Players (0)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "STRING"
         Text            =   "CC"
         Object.Width           =   688
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "STRING"
         Text            =   "Status"
         Object.Width           =   1111
      EndProperty
   End
   Begin VB.Label lblHide 
      Height          =   195
      Left            =   8400
      TabIndex        =   17
      Top             =   3480
      Width           =   525
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSignIn 
         Caption         =   "Sign &In"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileCreateGame 
         Caption         =   "Create &Game"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFileSignOut 
         Caption         =   "Sign Out"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "Select &All"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewTimestamp 
         Caption         =   "&Timestamp"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewGridlines 
         Caption         =   "&Gridlines"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewHighlight 
         Caption         =   "&Highlight alert words"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTheme 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewBMTheme 
         Caption         =   "&Standard Theme"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewDarkTheme 
         Caption         =   "&Dark Theme"
      End
      Begin VB.Menu mnuViewLightTheme 
         Caption         =   "&Light Theme"
      End
      Begin VB.Menu mnuViewClassicTheme 
         Caption         =   "&Classic Theme"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuToolsGTA2manager 
         Caption         =   "GTA2 &Manager..."
      End
      Begin VB.Menu mnuToolsIgnoreList 
         Caption         =   "&Ignore list..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpCommands 
         Caption         =   "&Keys, commands and links"
      End
      Begin VB.Menu mnuHelpPorts 
         Caption         =   "&Port forward help"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmGH"
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

Dim lii As LASTINPUTINFO

'Start: User Idle Time
Private Declare Function GetLastInputInfo Lib "user32" (plii As Any) As Long
Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type
'End: User Idle Time

'Mouse API and constants
'Private Const IDC_APPSTARTING = 32650&
Private Const IDC_HAND = 32649&
'Private Const IDC_ARROW = 32512&
'Private Const IDC_CROSS = 32515&
'Private Const IDC_IBEAM = 32513&
'Private Const IDC_ICON = 32641&
'Private Const IDC_NO = 32648&
'Private Const IDC_SIZE = 32640&
'Private Const IDC_SIZEALL = 32646&
'Private Const IDC_SIZENESW = 32643&
'Private Const IDC_SIZENS = 32645&
'Private Const IDC_SIZENWSE = 32642&
'Private Const IDC_SIZEWE = 32644&
'Private Const IDC_UPARROW = 32516&
'Private Const IDC_WAIT = 32514&
Private Declare Function LoadCursorLong Lib "user32" Alias "LoadCursorA" _
  (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" _
  (ByVal hCursor As Long) As Long

'Main Variables '''BenMillard'''
Dim blnCreateGameLoaded As Boolean 'true once CreateGame form has been displayed once
Dim lngLobby As Long
Dim blnReconnect As Boolean 'true if GH should reconnect
Dim intReconnect As Integer
Dim intSecondsWaited As Integer
Dim blnConnectClick As Boolean
Dim blnDisconnect As Boolean 'true if disconnect() has been called
Dim intPrevWinState As Integer 'Stores the windowstate before GH is sent to tray
Dim blnCalculatedGTA2checksum 'true if GTA2.exe CRC32 checksum has been calculated, false if you close GTA2
Dim intPreviousMapIndex As Integer 'GTA2 registry key
Dim strPreviousMapFile As String 'GMP file
Dim strPreviousScriptFile As String 'SCR file
Dim SortedArray As New cSortArray 'Stores map descriptions and sorts them like GTA2
Dim strHostCommentLastJoined As String 'The comment for the game the player last joined
Dim m_hwndEdit As Long
Dim IDL As Long, aPath As String 'used for special folders like program_files
Dim lCRC32 As Long 'stores file CRC
Dim blnItemInList As Boolean 'true if an item is found in a list
Dim strOldState As String
Dim cr As New cRegistry

'Toolbar Button Identifiers: '''BenMillard'''
Enum eToolbarButtons
    BTN_SIGN_IN = 0
    BTN_CREATE = 1
    BTN_OPTIONS = 2
    BTN_MANAGER = 3
    BTN_CANCEL = 4
    BTN_SIGN_OUT = 10 '''to be obsoleted by combining with BTN_SIGN_IN
End Enum

'''FOCUS'''
Dim btnMouse As Long 'which button was clicked (on lvPlayers)
Dim tabSelected As Long 'the current tabIRC.SelectedTab.Index (before an event changes it)

'Command Line '''BenMillard'''
Private Declare Function CommandLineToArgv Lib "shell32" Alias "CommandLineToArgvW" ( _
    ByVal lpCmdLine As Long, pNumArgs As Integer) As Long
Private Declare Function GlobalFree Lib "Kernel32" ( _
    ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
    pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function SysAllocString Lib "oleaut32" (ByVal pwsz As Long) As Long

Private Declare Function LocalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Any) As Long

'API call to get the XP effects:
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Private Sub cmdCancel_Click()
    blnCancel = True
End Sub

Private Sub cmdX_Click()
    Call form_KeyDown(vbKeyW, vbCtrlMask)
End Sub

Public Function URLdecshort(ByRef Text As String) As String
    Dim strArray() As String, lngA As Long
    strArray = Split(Replace(Text, "+", " "), "%")
    For lngA = 1 To UBound(strArray)
        strArray(lngA) = Chr$("&H" & Left$(strArray(lngA), 2)) & Mid$(strArray(lngA), 3)
    Next lngA
    URLdecshort = Join(strArray, vbNullString)
End Function

Public Function URLencshort(ByRef Text As String) As String
    Dim lngA As Long, strChar As String
    For lngA = 1 To Len(Text)
        strChar = Mid$(Text, lngA, 1)
        If strChar Like "[A-Za-z0-9]" Then
        ElseIf strChar = " " Then
            strChar = "+"
        Else
            strChar = "%" & Right$("0" & Hex$(Asc(strChar)), 2)
        End If
        URLencshort = URLencshort & strChar
    Next lngA
End Function

Public Function sUniPtrZToVBString(lStrptr As Long) As String
' Convert 'pointer to (wide-)null-terminated Unicode string' to a 'VB string Value '
    sUniPtrZToVBString = vbNullString
    lStrptr = SysAllocString(lStrptr)
    CopyMemory ByVal VarPtr(sUniPtrZToVBString), ByVal VarPtr(lStrptr), 4
End Function

Private Function ParseCommandLine() As String()
' Parse the current command line and return an Argv() array of strings
On Error GoTo oops

Dim lpArgv As Long, lpStr As Long
Dim lArgc As Integer
Dim i As Integer
Dim sArgs() As String
Dim decoded As String
    ' Parse the command line, breaking it up into separate options/parameters
    decoded = command$()
'        print (decoded)
    If Right(decoded, 1) = "/" Then decoded = Mid(decoded, 1, Len(decoded) - 1)
'        print (decoded)
    lpArgv = CommandLineToArgv(StrPtr(URLdecshort(decoded)), lArgc)
    ' Resize our array to hold the component strings
    ReDim sArgs(lArgc - 1)
    For i = 0 To lArgc - 1
        ' Get the address of the Unicode text for the i'th component
        CopyMemory lpStr, ByVal lpArgv + 4 * i, 4
        ' Convert this to a proper VB string value, and store it in out array
        sArgs(i) = sUniPtrZToVBString(lpStr)
    Next i
    ' Return the string array to our caller
    ParseCommandLine = sArgs
    ' Free the memory allocated by CommandLineToArgvW
    GlobalFree lpArgv
    Exit Function

oops:
    Call ErrorHandler("ParseCommandLine", Err.Description, Erl)

End Function

Public Sub form_load()
On Error GoTo oops

picGH = Me.Icon
picTray = Me.Icon

'MsgBox "main"
'Set colors
strServerColor = 32768
strCTCPcolor = 8388736
strTextColor = vbBlack
strConnectionColor = 8388736   'QBColor(Val("&H" & 5)) purple
strHelpColor = 8388736
strActionColor = 8388736
strTopicColor = vbBlack
strBannedColor = vbRed
strGHColor = 32896             'QBColor(Val("&H" & 6)) tan
strPrivateMessageColor = 128   'QBColor(Val("&H" & 4)) brown
strLinkColor = vbRed

bln98 = getVersion 'true if Windows 98

frmGH.Caption = "Game Hunter v" & TXT_GHVER

''Detect special folders
'  aPath = Space$(MAX_PATH)
'
'  Dim p
'  For p = 1 To 200
'  If SHGetSpecialFolderLocation(hwnd, p, IDL) = 0 Then
'        If SHGetPathFromIDList(IDL, aPath) Then
'            PROGRAM_FILES = Left$(aPath, InStr(aPath, vbNullChar) - 1)
'            Debug.Print p & " " & Left$(aPath, InStr(aPath, vbNullChar) - 1)
'        End If
'        LocalFree IDL
'  End If
'  Next p
'  Exit Sub

'Store Windows folder in WINDOWS_FOLDER global string
aPath = Space$(MAX_PATH)
If SHGetSpecialFolderLocation(hwnd, 36, IDL) = 0 Then
      If SHGetPathFromIDList(IDL, aPath) Then
          WINDOWS_FOLDER = Left$(aPath, InStr(aPath, vbNullChar) - 1)
      End If
      LocalFree IDL
End If

'Store Programs Files folder in PROGRAM_FILES global string
aPath = Space$(MAX_PATH)
If SHGetSpecialFolderLocation(hwnd, 38, IDL) = 0 Then
      If SHGetPathFromIDList(IDL, aPath) Then
          PROGRAM_FILES = Left$(aPath, InStr(aPath, vbNullChar) - 1)
      End If
      LocalFree IDL
End If

'Store documents folder in DOCUMENTS global string
aPath = Space$(MAX_PATH)
If SHGetSpecialFolderLocation(hwnd, 5, IDL) = 0 Then
      If SHGetPathFromIDList(IDL, aPath) Then
          DOCUMENTS = Left$(aPath, InStr(aPath, vbNullChar) - 1)
      End If
      LocalFree IDL
End If


'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
'''EnableAutoURLDetection rtbHistory(1)
'''EnableAutoURLDetection rtbTopic(1)

strPasswordProtectGame = "No"
intNickservWaitTime = -1
strChannel = "#gta2gh"
'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
'lvPlayers(1).ToolTipText = strChannel
lvPlayers(1).Tag = strChannel
strDestTab = vbNullString

'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
With rtbTopic(1)
    .SelStart = 0
    .SelLength = 666
    .SelText = "https://GTAMP.com https://GTAMP.com/forum"
    .Tag = strChannel
    '.ToolTipText = strChannel
End With

'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
m_hwndEdit = rtbHistory(tabIRC.SelectedItem.Index).hwnd

' No shortcut key text until the context menu is shown & hidden
SetEditMenuItemText False

Randomize

With cr
    'Force debug tab, blood, debug keys and min_frame_rate
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Software\DMA Design Ltd\GTA2\Debug\"
    'set blnChkSync to GTA2 do_sync_check value
    .ValueKey = "do_sync_check"
    If .Value = vbNullString Then
        blnChkSync = False
    Else
        blnChkSync = True
    End If
    .ValueKey = "bob_debug_display"
    .ValueType = REG_DWORD
    .Value = 0
    .ValueKey = "do_blood"
    .ValueType = REG_DWORD
    .Value = 0
    .ValueKey = "do_debug_keys"
    .ValueType = REG_DWORD
    .Value = 0
    .SectionKey = "Software\DMA Design Ltd\GTA2\screen\"
    .ValueKey = "min_frame_rate"
    .ValueType = REG_DWORD
    .Value = 1
    .ValueKey = "gamma"
    If .Value = 0 Then .Value = 15

    'IRC settings:
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter"
    'Load IRC username from registry
    .ValueKey = "Username"
    strPreferedNick = .Value
    'Load timestamp display setting
    .ValueKey = "chkTimestamp"
    If .Value = vbNullString Then
        blnchkTime = True
        .ValueType = REG_DWORD
        .Value = 1
    Else
        blnchkTime = .Value
    End If
    
End With

strNick = Left$(strPreferedNick, 16)

blnDisconnectClick = True

'Set country codes:
strCountries = Array("Somewhere - ??", _
"Afghanistan - AF", "Aland Islands - AW", "Albania - AL", "Algeria - DZ", "American Samoa - AS", "Andorra - AD", "Angola - AO", "Anguilla - AI", "Antarctica - AQ", "Antigua and Barbuda - AG", "Argentina - AR", "Armenia - AM", "Aruba - AW", "Australia - AU", "Austria - AT", "Azerbaijan - AZ", "Bahamas - BS", "Bahrain - BH", "Barbados - BB", "Bangladesh - BD", "Belarus - BY", "Belgium - BE", "Belize - BZ", "Benin - BJ", "Bermuda - BM", "Bahamas - BS", "Bhutan - BT", "Botswana - BW", "Bolivia - BO", "Bosnia and Herzegovina - BA", "Bouvet Island - BV", "Brazil - BR", "British Indian Ocean Territory - IO", "Brunei Darussalam - BN", "Bulgaria - BG", "Burkina Faso - BF", "Burundi - BI", "Cambodia - KH", "Cameroon - CM", "Canada - CA", "Cape Verde - CV", "Cayman Islands - KY", "Central African Republic - CF", "Chad - TD", "Chile - CL", "China - CN", "Christmas Island - CX", "Cocos (Keeling) Islands - CC", "Colombia - CO", "Comoros - KM", "Congo - CG", "Congo, Democratic Republic - CD", "Cook Islands - CK", _
"Costa Rica - CR", "Cote D'Ivoire (Ivory Coast) - CI", "Croatia (Hrvatska) - HR", "Cuba - CU", "Cyprus - CY", "Czech Republic - CZ", "Czechoslovakia (former) - CS", "Denmark - DK", "Djibouti - DJ", "Dominica - DM", "Dominican Republic - DO", "Ecuador - EC", "Egypt - EG", "El Salvador - SV", "Equatorial Guinea - GQ", "Eritrea - ER", "Estonia - EE", "Ethiopia - ET", "European Union - EU", "Falkland Islands (Malvinas) - FK", "Faroe Islands - FO", "Fiji - FJ", "Finland - FI", "France - FR", "France, Metropolitan - FX", "French Guiana - GF", "French Polynesia - PF", "French Southern Territories - TF", "F.Y.R.O.M. (Macedonia) - MK", "Gabon - GA", "Gambia - GM", "Georgia - GE", "Germany - DE", "Ghana - GH", "Gibraltar - GI", "Greece - GR", "Greenland - GL", "Grenada - GD", "Guadeloupe - GP", "Guam - GU", "Guatemala - GT", "Guernsey - GF", "Guinea - GN", "Guinea-Bissau - GW", "Guyana - GY", "Haiti - HT", "Heard and McDonald Islands - HM", "Honduras - HN", "Hong Kong - HK", "Hungary - HU", _
"Iceland - IS", "India - IN", "Indonesia - ID", "Iran - IR", "Iraq - IQ", "Ireland - IE", "Israel - IL", "Isle of Man - IM", "Italy - IT", "Jersey - JE", "Jamaica - JM", "Japan - JP", "Jordan - JO", "Kazakhstan - KZ", "Kenya - KE", "Kiribati - KI", "Korea (North) - KP", "Korea (South) - KR", "Kuwait - KW", "Kyrgyzstan - KG", "Laos - LA", "Latvia - LV", "Lebanon - LB", "Liechtenstein - LI", "Liberia - LR", "Libya - LY", "Lesotho - LS", "Lithuania - LT", "Luxembourg - LU", "Macau - MO", "Madagascar - MG", "Malawi - MW", "Malaysia - MY", "Maldives - MV", "Mali - ML", "Malta - MT", "Marshall Islands - MH", "Martinique - MQ", "Mauritania - MR", "Mauritius - MU", "Mayotte - YT", "Mexico - MX", "Micronesia - FM", "Monaco - MC", "Moldova - MD", "Morocco - MA", "Mongolia - MN", "Montenegro - ME", _
"Montserrat - MS", "Mozambique - MZ", "Myanmar - MM", "Namibia - NA", "Nauru - NR", "Nepal - NP", "Netherlands - NL", "Netherlands Antilles - AN", "Neutral Zone - NT", "New Caledonia - NC", "New Zealand - NZ", "Nicaragua - NI", _
"Niger - NE", "Nigeria - NG", "Niue - NU", "Norfolk Island - NF", "Northern Mariana Islands - MP", "Norway - NO", "Oman - OM", "Pakistan - PK", "Palau - PW", "Palestinian Territory, Occupied - PS", "Panama - PA", "Papua New Guinea - PG", "Paraguay - PY", "Peru - PE", "Philippines - PH", "Pitcairn - PN", "Poland - PL", "Portugal - PT", "Puerto Rico - PR", "Qatar - QA", "Reunion - RE", "Romania - RO", "Russian Federation - RU", "Rwanda - RW", "S. Georgia and S. Sandwich Isls. - GS", "Saint Kitts and Nevis - KN", "Saint Lucia - LC", "Saint Vincent & the Grenadines - VC", "Samoa - WS", "San Marino - SM", "Sao Tome and Principe - ST", "Saudi Arabia - SA", "Senegal - SN", "Serbia - RS", "Seychelles - SC", "Sierra Leone - SL", "Singapore - SG", "Slovenia - SI", "Slovak Republic - SK", "Solomon Islands - SB", "Somalia - SO", "South Africa - ZA", "Spain - ES", "Sri Lanka - LK", "St. Helena - SH", "St. Pierre and Miquelon - PM", "Sudan - SD", "Suriname - SR", "Svalbard & Jan Mayen Islands - SJ", _
"Swaziland - SZ", "Sweden - SE", "Switzerland - CH", "Syria - SY", "Taiwan - TW", "Tajikistan - TJ", "Tanzania - TZ", "Thailand - TH", "Timor-Leste - TL", "Togo - TG", "Tokelau - TK", "Tonga - TO", "Trinidad and Tobago - TT", "Tunisia - TN", "Turkey - TR", "Turkmenistan - TM", "Turks and Caicos Islands - TC", "Tuvalu - TV", "Uganda - UG", "Ukraine - UA", "United Arab Emirates - AE", "United Kingdom - UK", "United States - US", "US Minor Outlying Islands - UM", "Uruguay - UY", "Uzbekistan - UZ", "Vanuatu - VU", "Vatican City State (Holy See) - VA", "Venezuela - VE", "Viet Nam - VN", "Virgin Islands (British) - VG", "Virgin Islands (U.S.) - VI", "Wallis and Futuna Islands - WF", "Western Sahara - EH", "Yemen - YE", "Zambia - ZM", "Zimbabwe - ZW", "Icecream - IC", "C64 - 64", "Hedgewars - HW")
    
'890   strEurope = Array( _
'      "Albania - AL", "Andorra - AD", "Austria - AT", "Belarus - BY", "Belgium - BE", "Bosnia and Herzegovina - BA", "Bulgaria - BG", "Burkina Faso - BF", _
'      "Croatia (Hrvatska) - HR", "Cyprus - CY", "Czech Republic - CZ", "Czechoslovakia (former) - CS", "Denmark - DK", "Estonia - EE", "Ethiopia - ET", "European Union - EU", "Falkland Islands (Malvinas) - FK", "Faroe Islands - FO", "Finland - FI", "France - FR", "France, Metropolitan - FX", "F.Y.R.O.M. (Macedonia) - MK", "Germany - DE", "Ghana - GH", "Greece - GR", "Greenland - GL", "Hungary - HU", _
'      "Iceland - IS", "Ireland - IE", "Isle of Man - IM", "Italy - IT", "Latvia - LV", "Liechtenstein - LI", "Lithuania - LT", "Luxembourg - LU", "Malta - MT", "Mauritius - MU", "Monaco - MC", "Moldova - MD", "Morocco - MA", "Montenegro - ME", _
'      "Netherlands - NL", _
'      "Norway - NO", "Poland - PL", "Portugal - PT", "Romania - RO", "San Marino - SM", "Serbia - RS", "Slovenia - SI", "Slovak Republic - SK", "Spain - ES", _
'      "Sweden - SE", "Switzerland - CH", "Ukraine - UA", "United Kingdom - UK", "Vatican City State (Holy See) - VA")
        
Dim i As Integer
    
'Used for finding a country code position to help insert a flag in the imagelist

'For i = 0 To UBound(strCountries)
'If Right$(strCountries(i), 2) = "ME" Then MsgBox i + 1
'Next i

Dim strTemp As String
strTemp = CopyURLToRAM("http://gtamp.com/server.txt")

If Left(strTemp, 2) = "s=" Then
    strTemp = Mid$(strTemp, 3, 666)
    Dim serverArray() As String
    serverArray = Split(strTemp)
    strServer(0) = serverArray(0)
    strPort = serverArray(1)
Else
    strServer(0) = "irc.gtanet.com"
    strPort = 6667
End If

'displaychat strChannel, strGHColor, strServer(0) & " " & strPort

'displaychat strChannel, strGHColor, "Latest GH version: " & CopyURLToRAM("http://gtamp.com/version.txt")
'strServer(0) = "127.0.0.1"

'ShellExecute hwnd, "runas", "sc stop upnphost", "", App.Path, vbNormalFocus
'ShellExecute hwnd, "runas", "sc config upnphost start= disabled", "", App.Path, vbNormalFocus

Dim strSystem As String, lngRet As Long
strSystem = Space(255)
lngRet = GetSystemDirectory(strSystem, 255)
strSystem = Left$(strSystem, lngRet)

'displaychat strChannel, strTextColor, "Trying to close upnphost: " & GetCommandOutput(strSystem & "\sc stop upnphost", True, False, True)
'displaychat strChannel, strTextColor, "Trying to disable upnphost startup: " & GetCommandOutput(strSystem & "\sc config upnphost start= disabled", True, False, True)

'displaychat strChannel, strTextColor, "Trying to close SSDPSRV: " & GetCommandOutput(strSystem & "\sc stop SSDPSRV", True, False, True)
'displaychat strChannel, strTextColor, "Trying to disable SSDPSRV startup: " & GetCommandOutput(strSystem & "\sc config SSDPSRV start= disabled", True, False, True)

Debug.Print GetCommandOutput(strSystem & "\sc stop SSDPSRV", True, False, True)
Debug.Print GetCommandOutput(strSystem & "\sc config SSDPSRV start= disabled", True, False, True)


With cr
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter"
    
    '''Window size and position (most visible settings should be first):
    .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
    .ValueKey = "Top"
    If .Value <> vbNullString And .Value >= -20000 And .Value <= 50000 Then Top = .Value
    .ValueKey = "Left"
    If .Value <> vbNullString And .Value >= -20000 And .Value <= 50000 Then Left = .Value
    .ValueKey = "Width"
    If .Value <> vbNullString And .Value >= -20000 And .Value <= 50000 Then Width = .Value
    .ValueKey = "Height"
    If .Value <> vbNullString And .Value >= 1000 And .Value <= 50000 Then Height = .Value
    .ValueKey = "WindowState"
    If .Value = vbNullString Then .Value = vbMaximized
    If .Value = "0" Or .Value = "2" Then
        WindowState = .Value
    End If
    
    'Create Game password:
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter"
    .ValueKey = "Password"
    If .Value <> vbNullString Then
        strPassword = .Value
    Else
        strPassword = "x"
    End If
    
    'Play sounds:
    .ValueKey = "Sound"
    strWave = .Value
    
    'Set sorting arrows and options on main ListViews:
    lvGames(0).ColumnHeaderIcons = ImageListSortIconIndicator 'imagelist1 '''
    lvPlayers(1).ColumnHeaderIcons = ImageListSortIconIndicator  '''
        
    'Players list sorting options:
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
    .ValueKey = "lvPlayers(1).SortOrder"
    lvPlayers(1).SortOrder = .Value
    .ValueKey = "lvPlayers(1).SortKey"
    lvPlayers(1).SortKey = .Value
    Call SortColumn(frmGH.lvPlayers(1), .Value + 1)
    
    'Games list sorting options:
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
    .ValueKey = "lvGames(0).SortOrder"
    lvGames(0).SortOrder = .Value
    .ValueKey = "lvGames(0).SortKey"
    lvGames(0).SortKey = .Value
    Call SortColumn(frmGH.lvGames(0), .Value + 1)
    
End With 'with Cr


'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
tabIRC.Top = Me.ScaleHeight - rtbChatbox(1).Height
rtbChatbox(1).Top = tabIRC.Top - tabIRC.Height
rtbChatbox(1).Tag = strChannel
'rtbChatbox(1).ToolTipText = strChannel
rtbHistory(1).Top = rtbTopic(1).Top + rtbTopic(1).Height + 20
rtbHistory(1).Height = rtbChatbox(1).Top - (rtbTopic(1).Top + rtbChatbox(1).Height + tabIRC.Height) + 90
rtbHistory(1).Tag = strChannel
'rtbHistory(1).ToolTipText = strChannel

'  .ClassKey = HKEY_CURRENT_USER
'  .SectionKey = "Software\DMA Design Ltd\GTA2\Option"
'  .ValueKey = "Language"
'  strLanguage = .Value

'Load settings for Settings window:
Call frmOptions.loadSettings

'Start in the systray if Run at startup is ticked
If blnchkStartTray = True Then
  Call Systray
  'Hide
End If
'Load and apply theming system:
With cr
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
    .ValueKey = "Theme"
    intTheme = .Value
    Select Case intTheme
        Case 1
            mnuViewBMTheme_Click
            Call loadViewSettings
        Case 0, 2
            mnuViewDarkTheme_Click
            Call loadViewSettings
        Case 3
            mnuViewLightTheme_Click
            Call loadViewSettings
        Case 4
            mnuViewClassicTheme_Click
            Call loadViewSettings
    End Select
End With

If command <> vbNullString Then
    Dim argv() As String
    Dim strMap As String
    Dim doExit As Boolean
    Dim doJoin As Boolean
    
    argv = ParseCommandLine()
    For i = 0 To UBound(argv())
      Select Case argv(i)
        Case "/s", "-s" 'Specify Server
            If UBound(argv()) >= i + 1 And (Left(argv(i + 1), 1) <> "/") And (Left(argv(i + 1), 1) <> "-") Then
                strServer(0) = argv(i + 1)
                i = i + 1
            End If

        Case "/e", "-e" 'Join/Enter IP or Name
            strIPAddress = argv(i + 1)
            'name joining not implemented yet
            If UBound(argv()) >= i + 1 And (Left(argv(i + 1), 1) <> "/") And (Left(argv(i + 1), 1) <> "-") Then
                If argv(i + 1) = "" Then
                strIPAddress = "127.0.0.1"
                'if valid name
                'if valid game id
                'if valid ip
                'else dont join
                Else: strIPAddress = argv(i + 1)
                End If
                i = i + 1
            End If
            doJoin = True

        Case "/c", "-c" 'Open Create Game dialog (optional: pre select map)
            strMap = ""
            If UBound(argv()) >= i + 1 And (Left(argv(i + 1), 1) <> "/") And (Left(argv(i + 1), 1) <> "-") Then
                strMap = argv(i + 1)
                i = i + 1
            End If

            If strMap <> "" Then
            With cr
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                .ValueKey = "MapDesc"
                .ValueType = REG_SZ
                .Value = strMap
            End With
            End If

            cmdHost_Click

        Case "/h", "-h" 'Host Game with Map
            If UBound(argv()) >= i + 1 And (Left(argv(i + 1), 1) <> "/") And (Left(argv(i), 1 + 1) <> "-") Then
                strMap = argv(i + 1)
                i = i + 1
            End If

            'unlike in /c the maps is not applyed successfully here

            If strMap <> "" Then
            With cr
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                .ValueKey = "MapDesc"
                .ValueType = REG_SZ
                .Value = strMap
            End With
            End If

            blnHosted = True
            If frmGH.PreHost = False Then Exit Sub
            Call frmGH.Host

            Call SendMessage(lngHandleStart, BM_CLICK, 0, 0)
            'active = Activate 'error


        Case "/p", "-p" 'Singleplayer
            If UBound(argv()) >= i + 1 And (Left(argv(i + 1), 1) <> "/") And (Left(argv(i), 1 + 1) <> "-") Then
                strMap = argv(i + 1)
                i = i + 1
            End If

            If strMap <> "" Then
            With cr
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                .ValueKey = "MapDesc"
                .ValueType = REG_SZ
                .Value = strMap
            End With
            End If

            frmCreateGame.cmdPlayAlone_Click 'no effect
            'active = Activate 'error

'        Case "/d", "-d" 'Download Maps
'            Do While UBound(argv()) >= i + 1 And (Left(argv(i + 1), 1) <> "/") And (Left(argv(i + 1), 1) <> "-")
'                'Call CopyURLToFile(argv(i + 1), GetTmpPath & "gta2map.7z")
'                i = i + 1
'            Loop
'
        Case "/j", "-j" 'Join Channel
            Do While UBound(argv()) >= i + 1 And (Left(argv(i + 1), 1) <> "/") And (Left(argv(i + 1), 1) <> "-")
                'join channels not implemented yet
                i = i + 1
            Loop

        Case "/q", "-q" 'Exit gh
            doExit = True

        Case "/i", "-i" 'Don't open new instance
            If App.PrevInstance = True Then
                'pass parameters not implemented
                doExit = True
            End If

        Case "/l", "-l" 'No Sign in
            blnchkConnectOnStartup = False

        Case "/h", "-h", "--h", "/?", "-?", "/help", "-help", "--help"
            blnchkConnectOnStartup = False
            Call displayCommands

      End Select
    Next i
    
    If doExit = True Then cmdExit_Click
    'If doJoin = True Then lvGames_click (0)
End If

'displaychat strDestTab, strGHColor, "Total GTA2 Lobby Time: " & Duration(Val(readINI("Statistics", "GTA2 Lobby Time", DOCUMENTS & "\gta2gh.ini")), 2)
'displaychat strDestTab, strGHColor, "Total GTA2 Running Time: " & Duration(Val(readINI("Statistics", "GTA2 Running Time", DOCUMENTS & "\gta2gh.ini")), 2)
                
If blnchkConnectOnStartup = True Then cmdToolbar_Click (BTN_SIGN_IN)

'Dim SharingMgr, EachConnection, ConnectionProps, item
'    Set SharingMgr = CreateObject("HNetCfg.HNetShare.1")  '1

'    displaychat strDestTab, vbRed, "The following network adapters were detected:"
'    For Each item In SharingMgr.EnumEveryConnection          '2
'        Set EachConnection = SharingMgr.INetSharingConfigurationForINetConnection(item) '3
'        Set ConnectionProps = SharingMgr.NetConnectionProps(item)   '4
'        displaychat strDestTab, vbRed, vbNullString & ConnectionProps.Name
'        'If EachConnection.InternetFirewallEnabled = True Then
'        '    displaychat strDestTab, vbRed, "Windows firewall is enabled: " & ConnectionProps.Name
'        'Else
'        '    displaychat strDestTab, vbRed, "Windows firewall is disabled: " & ConnectionProps.Name
'        'End If
'    Next
    
    Exit Sub
oops:
    'If Err.Number <> 429 Then
    '    If blnchkConnectOnStartup = True Then cmdToolbar_Click (BTN_SIGN_IN)
    'Else
        Print "Error during startup:" & Err.Description & " " & Err.Number & " Line: " & Erl
        displaychat strDestTab, vbRed, "Error during startup:" & Err.Description & " " & Err.Number & " Line: " & Erl
        
    'End If
End Sub

'Public Sub getCC()
'On Error GoTo oops:
'    'sckURL.Close
'    'sckURL.Connect "api.hostip.info", 80
'    'sckURL.Connect "geoloc.daiguo.com", 80
'    'http://api.hostip.info/country.php
'    'sckURL.Connect "gtamp.com", 80
'    'sckURL.Connect "www.maxmind.com"  'www.maxmind.com/app/mylocation
'    'sckURL.Connect "127.0.0.1", 80
'    Exit Sub
'oops:
'displaychat strChannel, strGHColor, "Error connecting to " & TXT_GEOSITE
'End Sub

Private Sub cmdExit_Click()
    'save settings to registry on exit
    Save_Click
    Call saveChannels
    With cr
        .SectionKey = "SOFTWARE\GTA2 Game Hunter\display"
        .ValueKey = "Theme"
        .Value = intTheme
        .ValueKey = "lvGames(0).SortKey"
        .Value = lvGames(0).SortKey
        .ValueKey = "lvGames(0).SortOrder"
        .Value = lvGames(0).SortOrder
        .ValueKey = "lvPlayers(1).SortKey"
        .Value = lvPlayers(1).SortKey
        .ValueKey = "lvPlayers(1).SortOrder"
        .Value = lvPlayers(1).SortOrder

        ''''save the Widths of all columns in ListView lvGames(0) (left)
        '''(Removed 2010-10-28.)
        
        'Save window size and position:
        .ValueKey = "WindowState"
        .Value = WindowState
        If .Value <> "2" Then
            .ValueKey = "Height"
            .ValueType = REG_SZ
            .Value = Height
            
            .ValueKey = "Width"
            .ValueType = REG_SZ
            .Value = Width
            
            .ValueKey = "Left"
            .ValueType = REG_SZ
            .Value = Left
            
            .ValueKey = "Top"
            .ValueType = REG_SZ
            .Value = Top
        End If
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\DMA Design Ltd\GTA2\Debug\"
        .ValueKey = "bob_debug_display"
        .ValueType = REG_DWORD
        .Value = 0
        .ValueKey = "do_blood"
        .ValueType = REG_DWORD
        .Value = 0
        .ValueKey = "do_debug_keys"
        .ValueType = REG_DWORD
        .Value = 0
        
        .ValueKey = "skip_frontend"
        .DeleteValue
        
        .ValueKey = "skip_mission"
        If .Value <> vbNullString Then .DeleteValue

        .ValueKey = "play_replay"
        If .Value <> vbNullString Then .DeleteValue
        
        .ValueKey = "do_sync_check"
        .ValueType = REG_DWORD
        If blnChkSync = True Then
            .Value = 0
        Else
            .DeleteValue
        End If
    End With
    RemoveSystray 'removes the icon from the systray

    End
End Sub

Public Function PreHost() As Boolean
On Error GoTo oops
    
    PreHost = True
    
    If Exists(strGTA2path & TXT_GTA2EXE) = False Then
        Dim strTemp As String
        strTemp = BrowseFile(strGTA2path & TXT_GTA2EXE)
        If strTemp = vbNullString Or Len(strTemp) < Len(TXT_GTA2EXE) Then
            PreHost = False
            Exit Function
        Else
            strGTA2path = Left$(strTemp, Len(strTemp) - Len(TXT_GTA2EXE))
        End If
    End If
    
    If DetectGTA2version = False Then Exit Function
    
    Call MoveMMPfiles 'moves MMP files from tempMMP to data and then removes tempMMP folder
    AddDescriptionAndFileToListView 'lvMMPlist
    
    'add all map descriptions from lvMMPlist to cSortArray (sorted array)
    Dim i As Integer
    Set SortedArray = Nothing
    
    For i = 1 To frmGH.lvMMPlist.ListItems.count
        SortedArray.AddItem LCase(Trim(frmGH.lvMMPlist.ListItems.Item(i).ListSubItems(4))) 'add map description to array (trim and convert to lcase)
    Next i
    
    'clear current IP address from registry to stop GTA2 from trying to scan anyone
    'before the Create Game button is clicked
    strIPAddress = "127.0.0.1" 'vbNullString
    Dim cr As New cRegistry
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\DMA Design Ltd\GTA2\Network\"
        .ValueKey = "TCPIPAddrStringd"
        If LenB(.Value) > 0 Then .DeleteValue
        .ValueKey = "TCPIPAddrStrings"
        If LenB(.Value) > 0 Then .DeleteValue
    End With
    
    Save_Click
Exit Function
oops:
strErrLine = Erl
strErrdesc = Err.Description
displaychat strDestTab, strTextColor, "Error during PreHost: " & strErrdesc & ", Line: " & strErrLine
send "PRIVMSG " & gta2ghbot & " :Error during PreHost function: " & strErrdesc & " " & strErrLine
End Function

Public Sub cmdHost_Click()
    Dim strTemp As String
    Static lngStaticCount As Long
    Dim lngCount As Long
    
    Call setGTA2path
    If AudioFileCheck = False Then Exit Sub
    
    'If the PlayerCount line is missing from the Tiny Town mmp file then it's a new MMP file
    'and we can delete the obsolete files.
    If readINI("MapFiles", "PlayerCount", strGTA2path & "data\mp1-6p.mmp") = vbNullString Then
        'I really should put all these filenames in an array
        Call modFileKill(strGTA2path & "data\downtown-2p.mmp")
        Call modFileKill(strGTA2path & "data\downtown-3p.mmp")
        Call modFileKill(strGTA2path & "data\downtown-4p.mmp")
        Call modFileKill(strGTA2path & "data\downtown-5p.mmp")
        Call modFileKill(strGTA2path & "data\downtown-2P.scr")
        Call modFileKill(strGTA2path & "data\downtown-3P.scr")
        Call modFileKill(strGTA2path & "data\downtown-4P.scr")
        Call modFileKill(strGTA2path & "data\downtown-5P.scr")
        Call modFileKill(strGTA2path & "data\industrial-2p.mmp")
        Call modFileKill(strGTA2path & "data\industrial-3p.mmp")
        Call modFileKill(strGTA2path & "data\industrial-4p.mmp")
        Call modFileKill(strGTA2path & "data\industrial-5p.mmp")
        Call modFileKill(strGTA2path & "data\Industrial-2P.scr")
        Call modFileKill(strGTA2path & "data\Industrial-3P.scr")
        Call modFileKill(strGTA2path & "data\Industrial-4P.scr")
        Call modFileKill(strGTA2path & "data\Industrial-5P.scr")
        Call modFileKill(strGTA2path & "data\mp1-2p.mmp")
        Call modFileKill(strGTA2path & "data\mp1-3p.mmp")
        Call modFileKill(strGTA2path & "data\mp1-4p.mmp")
        Call modFileKill(strGTA2path & "data\mp1-5p.mmp")
        Call modFileKill(strGTA2path & "data\MP1-2P.scr")
        Call modFileKill(strGTA2path & "data\MP1-3P.scr")
        Call modFileKill(strGTA2path & "data\MP1-4P.scr")
        Call modFileKill(strGTA2path & "data\MP1-5P.scr")
        Call modFileKill(strGTA2path & "data\mp2-2p.mmp")
        Call modFileKill(strGTA2path & "data\mp2-3p.mmp")
        Call modFileKill(strGTA2path & "data\mp2-4p.mmp")
        Call modFileKill(strGTA2path & "data\mp2-5p.mmp")
        Call modFileKill(strGTA2path & "data\MP2-2P.scr")
        Call modFileKill(strGTA2path & "data\MP2-3P.scr")
        Call modFileKill(strGTA2path & "data\MP2-4P.scr")
        Call modFileKill(strGTA2path & "data\MP2-5P.scr")
        Call modFileKill(strGTA2path & "data\mp5-2p.mmp")
        Call modFileKill(strGTA2path & "data\mp5-3p.mmp")
        Call modFileKill(strGTA2path & "data\mp5-4p.mmp")
        Call modFileKill(strGTA2path & "data\mp5-5p.mmp")
        Call modFileKill(strGTA2path & "data\MP5-2P.scr")
        Call modFileKill(strGTA2path & "data\MP5-3P.scr")
        Call modFileKill(strGTA2path & "data\MP5-4P.scr")
        Call modFileKill(strGTA2path & "data\MP5-5P.scr")
        Call modFileKill(strGTA2path & "data\res-2p.mmp")
        Call modFileKill(strGTA2path & "data\res-3p.mmp")
        Call modFileKill(strGTA2path & "data\res-4p.mmp")
        Call modFileKill(strGTA2path & "data\res-5p.mmp")
        Call modFileKill(strGTA2path & "data\res-2p.scr")
        Call modFileKill(strGTA2path & "data\res-3p.scr")
        Call modFileKill(strGTA2path & "data\res-4p.scr")
        Call modFileKill(strGTA2path & "data\res-5p.scr")
    End If
    
    If blnCreateGameLoaded = False Then
        If PreHost = False Then Exit Sub
        frmCreateGame.loadDisplaySettings
        frmCreateGame.loadFilterSettings
        blnCreateGameLoaded = True
    Else
        'Count the number of MMP files
        strTemp = Dir(strGTA2path & "data\*.mmp", vbHidden + vbNormal + vbSystem + vbReadOnly + vbArchive)
        Do While strTemp <> vbNullString
            strTemp = Dir
            lngCount = lngCount + 1
        Loop
        
        'If the number of MMP files has changed from last time then refresh the list
        If lngCount <> lngStaticCount Then
            Call frmCreateGame.cmdRefresh_Click
            lngStaticCount = lngCount
        End If
    End If
    
    frmCreateGame.txtFind.Text = vbNullString
    Call frmCreateGame.loadSortSettings
    frmCreateGame.Show
   
    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error hosting: " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :Error hosting: " & strErrdesc & " Line: " & strErrLine
End Sub

Public Sub Host()
On Error GoTo oops:
    Dim strReplay As String
    Call AudioFileCheck
    If IsGTA2running() = True Then Exit Sub
    If DetectGTA2version = False Then Exit Sub
    Call FindProcess("dplaysvr.exe", True) 'Find and kill process
  
    displaychat strChannel, strGHColor, "Launching GTA2"
    modMkDir strGTA2path & "test"
    
    If blnPlayReplay = True Then
        If Exists(strGTA2path & "test\replay.rep") Then strReplay = " -r"
    End If
    
    blnInGame = False
    
    With cr
        .ValueKey = "do_sync_check"
        .ValueType = REG_DWORD
        If blnChkSync = True Then
            .Value = 0
        Else
            .DeleteValue
        End If
    End With
    
    lngPID = shellandwait(strGTA2path & TXT_GTA2EXE & " -n -c" & strReplay, strGTA2path)
    
    blnHosted = True
    
    If blnConnected = False Then
        displaychat strDestTab, strGHColor, "Sign in to publically advertise your game."
    End If
    
    Call SetPlayerName
    
Exit Sub
oops:
 strErrLine = Erl
 strErrdesc = Err.Description
 displaychat strDestTab, strTextColor, "Error during host function: " & strErrdesc & ", Line: " & strErrLine
 send "PRIVMSG " & gta2ghbot & " :Error during host function: " & strErrdesc & " " & strErrLine
End Sub

Public Sub SetPlayerName()
On Error GoTo oops
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\DMA Design Ltd\GTA2\Network"
        .ValueKey = "PlayerName"
        If .Value = vbNullString Then
            .ValueType = REG_SZ
            .Value = strNick
        End If
    End With
Exit Sub

oops:
    Call ErrorHandler("setPlayerName", Err.Description, Erl)
End Sub

Public Function IsGTA2running() As Boolean
    IsGTA2running = False
    Call FindProcess(TXT_GTA2EXE, True) 'Find and kill process
End Function

Public Sub cmdJoin_Click()
    On Error GoTo oops
    
    If Exists(strGTA2path & TXT_GTA2EXE) = False Then Exit Sub
    If DetectGTA2version = False Then Exit Sub
    Call Save_Click
    
    Call SetPlayerName 'if playername is blank then change it to IRC nick
    modMkDir strGTA2path & "test"
    If blnPlayReplay = True Then
        blnPlayReplay = False
        Dim strPlayReplay As String
        strPlayReplay = " -r"
    End If
    displaychat strChannel, strGHColor, "Launching " & strGTA2path & TXT_GTA2EXE ' & " -n -j" & strPlayReplay
    blnInGame = False
    lngPID = shellandwait(strGTA2path & TXT_GTA2EXE & " -n -j" & strPlayReplay, strGTA2path)
Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
If Err.Number = 5 Then
    displaychat strDestTab, strTextColor, "Unable to launch GTA2."
    send "PRIVMSG " & gta2ghbot & " :Launch failure"
Else
    displaychat strDestTab, strTextColor, "Error joining: " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :Error joining: " & strErrdesc & " " & strErrLine
End If

End Sub

Private Sub Save_Click()
    Dim i As Integer
    Dim j As Integer
    On Error GoTo oops
    Dim byteIPAddress() As Byte

    With cr
        'gta2 game hunter settings
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueKey = "GTA2folder"
        .ValueType = REG_SZ
        If Len(strGTA2path) > 0 Then
            .Value = strGTA2path
        End If
        .ValueKey = "GHfolder"
        .ValueType = REG_SZ
        .Value = App.Path

        'HKEY_CURRENT_USER\Control Panel\Accessibility\StickyKeys
        'displaychat strDestTab, vbRed, "Writing registry key HKEY_CURRENT_USER\Control Panel\Accessibility\StickyKeys\Flags: new value = 26"

        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Control Panel\Accessibility\StickyKeys"
        .ValueKey = "Flags"
        .ValueType = REG_SZ
        .Value = 26

        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Software\DMA Design Ltd\GTA2\Network"

        'if game time limit < 5 then set it to unlimited
        .ValueKey = "game_time_limit"
        If .Value < 5 Then
            .ValueType = REG_DWORD
            .Value = 0
        End If

        'if no frag limit
        .ValueKey = "f_limit"
        If LenB(.Value) = 0 Then
            .ValueType = REG_DWORD
            .Value = 10
        End If

        'if no protocol is selected then select 0
        .ValueKey = "protocol_selected"
        If LenB(.Value) = 0 Then
            .ValueType = REG_DWORD
            .Value = 0
        End If

        'store IP address in byte array
        byteIPAddress = strIPAddress & vbNullChar & vbNullChar 'add two nulls to the end
        ' Store IP Address in registy as hex
        .ValueKey = "TCPIPAddrStringd"
        .ValueType = REG_BINARY
        .Value = byteIPAddress() 'add IP address array to registry

        .ValueKey = "TCPIPAddrStrings" 'this key stores the length of the IP
        .ValueType = REG_DWORD
        .Value = Len(strIPAddress) * 2 + 2 '*2 and then add 2, to include nulls in length

        Dim eek() As Byte
        .ValueKey = "UseConnectiond"

        Dim strUseConnectiond As String
        j = 0

        byteIPAddress = strIPAddress
        strUseConnectiond = "60f518132c91d0119daa00a0c90a43cb04000000FF000000c016d907afe0cf119c4e00a0c905425e10000000e05ee9367785cf11960c0080c7534e82a03232e6bf9dd0119cc100a0c905425e" & UBound(byteIPAddress)
        If Len(strIPAddress) < 51 Then
          strUseConnectiond = strUseConnectiond & ".."
        Else
          strUseConnectiond = strUseConnectiond & "."
        End If
        ReDim eek(Len(strUseConnectiond) / 2 + Len(strIPAddress) * 2 + 5) As Byte
        For i = 1 To 152 Step 2
            eek(j) = CByte("&H" & Mid(strUseConnectiond, i, 2))
            j = j + 1
        Next i

        eek(20) = CByte("&H" & Hex(Len(strUseConnectiond) / 2 + Len(strIPAddress) * 2 + 4))
        eek(76) = CByte("&H" & Hex(Len(strIPAddress) * 2 + 2))

        For i = 0 To UBound(byteIPAddress)
            eek(80 + i) = byteIPAddress(i)
        Next i

        .ValueType = REG_BINARY
        .Value = eek()

        'UseConnectionS
        .ValueKey = "UseConnections" 'this key stores the length of the IP
        .ValueType = REG_DWORD
        i = eek(20)
        .Value = i
        Erase eek()
        strUseConnectiond = vbNullString

        ReDim eek(15) As Byte
        .ValueKey = "UseProtocold"
        If LenB(.Value) = 0 Then 'if the data is empty
            .ValueType = REG_BINARY
            .ValueKey = "UseProtocold"
            Dim UseProtocold As String
            j = 0
            UseProtocold = "e05ee9367785cf11960c0080c7534e82"
            For i = 1 To 32 Step 2
                eek(j) = CByte("&H" & Mid(UseProtocold, i, 2))
                j = j + 1
            Next i
            .Value = eek()
            Erase eek()
            ReDim eek(0) As Byte
            UseProtocold = vbNullString
        End If

        .ValueKey = "UseProtocols"
        If LenB(.Value) = 0 Then 'if the data is empty
            .ValueType = REG_DWORD
            .Value = 16
        End If
        
    End With
    
    Exit Sub
oops:
    strErrLine = Erl
    strErrdesc = Err.Description
    Call MsgBox("GTA2 needs write access to HKEY_CURRENT_USER\Software\DMA Design Ltd\GTA2" _
                    & vbNewLine & "Try logging in as administrator to fix the problem, you know the password right?" _
                    & vbNewLine & "Error writing to registry: " & strErrdesc & " - Line: " & strErrLine _
                    , vbExclamation, "GTA2 Game Hunter - error writing to registry")
End Sub

Private Sub mnuHelpPorts_Click()
Call displayPortHelp

'Call QueryAdaptersAddresses

'Debug.Print "IP return by GetAdaptersAddresses()" & vbCrLf & String(32, "-")
'Debug.Print LocalIP(True)
    
'Debug.Print "IP return by GetIpAddrTable()" & vbCrLf & String(32, "-")
'Debug.Print LocalIP(False)
    
End Sub

Private Sub mnuToolsIgnoreList_Click()
    Call ShellExecute(Me.hwnd, "Open", DOCUMENTS & "\gta2gh_ignore_list.txt", vbNullString, DOCUMENTS & "\gta2gh_ignore_list.txt", vbNormalFocus)
End Sub

Private Sub mnuViewGridlines_Click()
On Error GoTo oops
    Dim i As Integer
    mnuViewGridlines.Checked = Not mnuViewGridlines.Checked
    
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueKey = "Gridlines"
        .ValueType = REG_DWORD
        If mnuViewGridlines.Checked = True Then
            .Value = 1
        Else
            .Value = 0
        End If
            
        frmGH.lvGames(0).GridLines = mnuViewGridlines.Checked
        
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        For i = 1 To frmGH.lvPlayers.UBound
            frmGH.lvPlayers(i).GridLines = mnuViewGridlines.Checked
        Next
        
    End With
Exit Sub

oops:
    Call ErrorHandler("mnuViewGridlines", Err.Description, Erl)

End Sub

Private Sub mnuViewTimestamp_Click()
    blnchkTime = Not blnchkTime
    mnuViewTimestamp.Checked = blnchkTime
    
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueKey = "chkTimestamp"
        .ValueType = REG_DWORD
        If blnchkTime = True Then
            .Value = 1
        Else
            .Value = 0
        End If
    End With
End Sub

Private Sub mnuFileExit_Click()
    cmdExit_Click
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.about
End Sub

Private Sub cmdChat_Click()
On Error GoTo oops
   
'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
rtbChatbox(1).SelColor = strTextColor

Dim i As Integer
Dim x As Long
Dim y As Long
Dim intChatbox As Integer

intChatbox = -1

'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
For i = 1 To rtbChatbox.UBound
    If LCase(rtbChatbox(i).Tag) = LCase(tabIRC.SelectedItem) Then
        intChatbox = i
        Exit For
    End If
Next

If intChatbox < 1 Then Exit Sub
If rtbChatbox(intChatbox).Text = vbNullString Then Exit Sub

' /msg
Dim strChatMsg As String
Dim strChatCommand As String 'cmdX is all chars up to the first space
Dim strChatParam As String 'The first chat parameter, usually a channel
Dim strChatParam2 As String
Dim strChatExcludingCommand As String 'Chat message excluding "/command "
Dim bytFirstSpaceLoc As Byte
'Dim intFirstSpaceLoc As Integer
'Dim intFirstNotSpaceLoc As Integer
'Dim intSecondSpaceLoc As Integer

strChatMsg = Trim(rtbChatbox(intChatbox).Text)  'remove all leading and trailing spaces from msg and store in strChatMsg

If strChatMsg = vbNullString Then Exit Sub    'if there's no message exit the sub

If Left$(strChatMsg, 1) <> "/" Then 'if the message doesn't begin with a / then send it to the channel
    send "PRIVMSG " & tabIRC.SelectedItem.Caption & " :" & strChatMsg 'send the message to the active channel/nick
    If blnConnected Then
        Dim strPaddedNick As String
        strPaddedNick = strNick
        If blnchkPad And intTheme > 1 Then
            If Len(strNick) < 10 Then strPaddedNick = String(10 - Len(strNick), " ") & strNick
        End If
        displaychat tabIRC.SelectedItem.Caption, strTextColor, "<" & strPaddedNick & "> " & strChatMsg   'display the message"
    End If
Else
    'Msg is a command because it begins with a /
    bytFirstSpaceLoc = InStr(1, strChatMsg, " ")  'Location of the first space character in the msg
    If bytFirstSpaceLoc = 0 Then
        strChatCommand = strChatMsg
    Else
        strChatExcludingCommand = Mid$(strChatMsg, bytFirstSpaceLoc + 1, 666)
        strChatCommand = Left$(strChatMsg, bytFirstSpaceLoc - 1) 'cmdX is all chars up to the first space
    End If
    
    strChatCommand = UCase$(Right$(strChatCommand, Len(strChatCommand) - 1)) 'Strip off the / and convert to uppercase

    'The parameters start when there are no more space characters
    If bytFirstSpaceLoc > 0 Then
        Do Until Mid$(strChatMsg, bytFirstSpaceLoc, 1) <> " " Or bytFirstSpaceLoc = 255
            bytFirstSpaceLoc = bytFirstSpaceLoc + 1
        Loop
        
        intLinePosition = bytFirstSpaceLoc
        Call AddChar(intLinePosition, strChatMsg, strChatParam)
    End If

Select Case strChatCommand
  Case "MV"
        Dim b As Boolean
        b = frmGH.mnuView.Visible
        frmGH.mnuFile.Visible = Not (b)
        frmGH.mnuEdit.Visible = Not (b)
        frmGH.mnuView.Visible = Not (b)
        frmGH.mnuTools.Visible = Not (b)
        frmGH.mnuHelp.Visible = Not (b)
  Case "K", "KICK"
        If InStr(strChatMsg, "#") Then
            send "KICK " & strChatExcludingCommand
        Else
            If Left$(tabIRC.SelectedItem.Caption, 1) = "#" Then
                send "KICK " & tabIRC.SelectedItem.Caption & Right$(strChatMsg, Len(strChatMsg) - Len(strChatCommand) - 1)
            Else
                send "KICK " & strChannel & Right$(strChatMsg, Len(strChatMsg) - Len(strChatCommand) - 1)
            End If
        End If
  Case "RAWW"
      send Right$(strChatMsg, Len(strChatMsg) - 6)
  Case "ME", "ACTION"
      If strChatParam = vbNullString Then Exit Sub
      send "PRIVMSG " & tabIRC.SelectedItem.Caption & " :" & "ACTION " & Right$(strChatMsg, Len(strChatMsg) - bytFirstSpaceLoc + 1) & ""
      displaychat tabIRC.SelectedItem.Caption, strActionColor, strNick & " " & Right$(strChatMsg, Len(strChatMsg) - bytFirstSpaceLoc + 1)
  Case "CLEAR", "CLS"
      rtbHistory(intChatbox).Text = vbNullString
  Case "DNS", "D"
      displaychat strDestTab, strConnectionColor, "Trying to resolve " & strChatParam
      Dim strIPFromHostName As String
      strIPFromHostName = GetIPFromHostName(strChatParam)
      If strIPFromHostName = vbNullString Then
          displaychat strDestTab, strConnectionColor, "Unable to resolve " & strChatParam
      Else
          displaychat strDestTab, strConnectionColor, "Resolved " & strChatParam & " to " & strIPFromHostName
      End If
  Case "ICMP"
        Dim replyInfo As ICMP_ECHO_REPLY
        Dim replyData As Long
        displaychat strDestTab, strConnectionColor, "ping " & strChatParam & " - please wait"
        replyData = ping(strChatParam, replyInfo)
        If replyData = 0 Then
            displaychat strDestTab, strConnectionColor, EvaluatePingResponse(replyData) & " in " & replyInfo.RoundTripTime & " ms"
  Else:
            displaychat strDestTab, strConnectionColor, EvaluatePingResponse(replyData)
        End If
  Case "AWAY", "A"
        Call toggleAwayStatus(strChatExcludingCommand)
  Case "BACK", "B"
        Call back
        Call changeStatus(strStatusMsg)
  Case "MSG", "M"
      If bytFirstSpaceLoc = 0 Then Exit Sub
      
      Do Until Mid$(strChatMsg, intLinePosition) <> " " Or bytFirstSpaceLoc = 255
          intLinePosition = intLinePosition + 1
      Loop
      If strChatParam = vbNullString Then Exit Sub
      If Len(strChatMsg) + 1 = intLinePosition Then Exit Sub
      strChatParam2 = Trim(Right$(strChatMsg, Len(strChatMsg) - intLinePosition))
        'only send a message if the user is on the same channel as you
        If InStr(LCase(strChatParam), "serv") = False Then
            If onYourChannel(strChatParam) = False Then
                displaychat tabIRC.SelectedItem.Caption, strGHColor, "You must be in the same channel as " & strChatParam & " to send them a message."
                Exit Sub
            End If
        End If
      
      send "PRIVMSG " & strChatParam & " :" & strChatParam2
      displaychat tabIRC.SelectedItem.Caption, strTextColor, "-> *" & strChatParam & "* " & strChatParam2
  Case "QUERY", "Q"
      displaychat strChatParam, strTextColor, vbNullString
  Case "HOST", "IP"
      displaychat strDestTab, strConnectionColor, "Your external host name is: " & strExternalHostName
  Case "VPN"
      blnchkVPN = Not (blnchkVPN)
      With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueKey = "chkVPN"
        .Value = blnchkVPN
      End With
      displaychat strDestTab, strTextColor, "VPN mode switched to: " & blnchkVPN
  Case "QUIT", "EXIT"
      send "QUIT :" & Mid$(strChatMsg, 7, 255)
      'blnDONOTCHANGESERVER = True
      blnDisconnectClick = True
      Disconnect
  Case "IGNORE", "I"
      If strChatParam = vbNullString Then
        displaychat tabIRC.SelectedItem.Caption, strGHColor, "Ignore list: file://" & DOCUMENTS & "\gta2gh_ignore_list.txt"
      Else
        displaychat tabIRC.SelectedItem.Caption, strGHColor, "Added " & strChatParam & " to ignore list: file://" & DOCUMENTS & "\gta2gh_ignore_list.txt"
        Call WriteINI("Ignore", strChatParam, "True", DOCUMENTS & "\gta2gh_ignore_list.txt")
      End If
  Case "HIDE"
      blnHidden = Not blnHidden
      
      If blnHidden = True Then
          With cr
              .ClassKey = HKEY_CURRENT_USER
              .SectionKey = "SOFTWARE\GTA2 Game Hunter"
              .ValueKey = "Hide"
              .Value = 1
          End With

          send "hs on"
          displaychat strDestTab, strGHColor, "Host hidden"
      Else
          With cr
              .ClassKey = HKEY_CURRENT_USER
              .SectionKey = "SOFTWARE\GTA2 Game Hunter"
              .ValueKey = "Hide"
              .Value = 0
          End With
          
          send "hs off"
          displaychat strDestTab, strGHColor, "Host visible"
      End If
  'Case "FONT", "F"
  '   Call displayDialogFont("History")
  'Case "CC"
  '    Call updateCountry(strChatParam, "GH",intDestTab)
  'Case "L"
  '    displaychat strDestTab, strGHColor, "Listening on port " & frmProbeGTA2.Socket(0).LocalPort
  Case "T", "TERMINATE"
      Call FindProcess(TXT_GTA2EXE, True) 'Find and kill process
      Call FindProcess("dplaysvr.exe", True) 'Find and kill process
  Case "KILL", "PSKILL", "TASKKILL"
      If strChatParam = vbNullString Then
        displaychat strChannel, strGHColor, "/kill needs an executable name"
      Else
        Call FindProcess(strChatParam, True)
      End If
  Case "GET"
       Call CopyURLToFile(strChatParam, GetTmpPath & "gta2map.7z")
       'Call CopyURLToFile("http://gtamp.com/maps/test ing.7z", GetTmpPath & "gta2map.7z")
      'displaychat strDestTab, strGHColor, GetCommandOutput(App.Path & "\mydown.exe -x http://127.0.0.1/gta2.7z", , True)
  Case "WHOIS", "W", "WI", "WII"
      'name is sent twice so it works on players who are on different linked servers (IRC oddity)
      send "WHOIS " & strChatParam & " " & strChatParam
  Case "/", "?", "H", "HELP", "NOTICE", "SETNAME"
      Call displayCommands
  Case "J", "JOIN"
      send "JOIN " & strChatExcludingCommand
  Case "P", "PART"
      If strChatExcludingCommand = vbNullString Then strChatExcludingCommand = tabIRC.SelectedItem
      If strChatExcludingCommand <> strChannel Then
        For i = 1 To tabIRC.Tabs.count
            If tabIRC.Tabs(i).Caption = strChatExcludingCommand Then
                tabIRC.Tabs(i).Selected = True
                send "PART " & strChatExcludingCommand
                Call form_KeyDown(vbKeyW, vbCtrlMask)
                Exit For
            End If
        Next i
      End If
  Case "HOP", "CYCLE"
        send "CYCLE " & tabIRC.SelectedItem.Caption & " " & strChatParam
  Case "E"
      blnHosted = True
      If frmGH.PreHost = False Then Exit Sub
      Call frmGH.Host 'ehost
  Case "C"
      blnHosted = True
      If frmGH.PreHost = False Then Exit Sub
      Call frmGH.Host
  Case "R" 'Set GH to custom resoltion eg: /r 640x480
      i = InStr(LCase(strChatParam), "x")
      If i = 0 Then
        Call winRes(640, 480)
      Else
        If Val(Mid$(strChatParam, 1, i - 1)) < 31536000 Then x = Val(Mid$(strChatParam, 1, i - 1))
        If Val(Mid$(strChatParam, i + 1, 12)) < 31536000 Then y = Val(Mid$(strChatParam, i + 1, 12))
        Call winRes(x, y)
      End If
'  Case "CRCALL"
'        Dim oFileSystem As New FileSystemObject
'        Dim oFolder As Folder
'        Dim oCurrentFile As File
'        Dim oFileColl As Files
'
'        Set oFolder = oFileSystem.GetFolder(strGTA2path & "data\")
'        Set oFileColl = oFolder.Files
'
'
'        If oFileSystem.FolderExists(oFolder) = False Then Exit Sub
'
'        'Move all files in gta2\data\tempMMP to gta2\data
'        If oFileColl.Count > 0 Then
'            For Each oCurrentFile In oFileColl
'                Open "c:\temp\" & LCase(oCurrentFile.Name) & ".crc" For Output As #1
'                Print #1, calc_crc32(strGTA2path & "data\" & oCurrentFile.Name)
'                Close #1
'            Next
'        End If
'
'        Set oFileSystem = Nothing
'        Set oFolder = Nothing
'        Set oFileColl = Nothing
'        Set oCurrentFile = Nothing
'
  Case "CRC32", "CRC"
      If strChatParam = vbNullString Then Exit Sub
      strChatParam2 = calc_crc32(strChatExcludingCommand)
      If strChatParam2 = "00000000" Then
        displaychat strDestTab, strGHColor, "File not Found " & strChatExcludingCommand
      Else
        displaychat strDestTab, strGHColor, strChatExcludingCommand & " CRC32 " & strChatParam2
      End If
  Case "RUN"
      If strChatParam = vbNullString Then Exit Sub
      If ShellExecute(Me.hwnd, "Open", strChatExcludingCommand, vbNullString, Mid$(strChatExcludingCommand, 1, InStrRev(strChatExcludingCommand, "\")), vbNormalFocus) = 2 Then
        displaychat strDestTab, strGHColor, "Failed to run " & strChatExcludingCommand
      End If
  Case "RAWHIDE"
        displaychat strChannel, strGHColor, strRaw
'  Case "CTCP"
'      i = InStr(1, strChatExcludingCommand, " ")
'      If i Then
'            strChatParam2 = Mid$(strChatExcludingCommand, i + 1, 666)
'            send "PRIVMSG " & strChatParam & " " & Chr$(1) & UCase$(strChatParam2) & Chr$(1)
'      End If
  Case "TOPIC"
        If strChatParam = vbNullString Then Exit Sub
        If Left$(strChatParam, 1) = "#" Then
            If InStr(strChatExcludingCommand, " ") Then
                send "TOPIC " & strChatExcludingCommand
            Else
                Exit Sub
            End If
        Else
            send "TOPIC " & tabIRC.SelectedItem.Caption & "  " & strChatExcludingCommand
        End If
  Case "SERVER"
        strServer(0) = strChatParam
        strChatParam2 = Trim(Right$(strChatMsg, Len(strChatMsg) - intLinePosition))
        strPort = strChatParam2
        cmdDisconnectClick
        cmdToolbar_Click (BTN_SIGN_IN)
  Case Else
      send Mid$(strChatMsg, 2, 666)
End Select

End If

rtbChatbox(intChatbox).Text = vbNullString    'clear the field
'Add code to scroll to bottom if you send a message

Exit Sub

oops:
strErrdesc = Err.Description
strErrLine = Erl
displaychat strDestTab, vbRed, "Error during Chat_Click(): " & strErrdesc & " - Line: " & strErrLine
send "PRIVMSG " & gta2ghbot & " :Error during Chat_Click(): " & strErrdesc & " - Line: " & strErrLine
End Sub

Private Sub displayPortHelp()

strDestTab = frmGH.tabIRC.SelectedItem 'display in current tab
displaychat strDestTab, strHelpColor, vbNullString
Call setGTA2path
displaychat strDestTab, strGHColor, "GTA2 path: " & strGTA2path

Call GetMACs_AdaptInfo

End Sub

Private Sub displayCommands()

strDestTab = frmGH.tabIRC.SelectedItem 'display in current tab
displaychat strDestTab, strHelpColor, vbNullString
displaychat strDestTab, strHelpColor, "Commands, controls and help:"
displaychat strDestTab, strHelpColor, "Private message: /msg GTA2Guy What's the password?"
displaychat strDestTab, strHelpColor, "Set your status to away/back: /away or /back"
displaychat strDestTab, strHelpColor, "Say you are doing something: /me shot the food!"
displaychat strDestTab, strHelpColor, "View player details: /wi GTA2Guy"
displaychat strDestTab, strHelpColor, "Get an IP from a hostname: /dns GTAMP.com"
displaychat strDestTab, strHelpColor, "Ping an IP through ICMP: /icmp 192.168.1.17"
displaychat strDestTab, strHelpColor, "Clear the currently selected chat history: /clear or /cls"
displaychat strDestTab, strHelpColor, "Quickly create a game: /c"
displaychat strDestTab, strHelpColor, "Set GH window resolution: /r 640x480"
displaychat strDestTab, strHelpColor, "Exit with a quit message: /quit I regret nothing!"
displaychat strDestTab, strHelpColor, "Leave channel: ctrl+w or click the X on the bottom right"
displaychat strDestTab, strHelpColor, "Rejoin channel: /hop or /cycle"
displaychat strDestTab, strHelpColor, "Close all processes named gta2.exe and dplaysvr.exe: /t"
displaychat strDestTab, strHelpColor, "Close all processes named notepad: /kill notepad"
displaychat strDestTab, strHelpColor, "Display CRC32 of a file: /crc gta2.exe"
displaychat strDestTab, strHelpColor, "ctrl+w to hide the current tab"
displaychat strDestTab, strHelpColor, "ctrl+F4 to hide all tabs"
displaychat strDestTab, strHelpColor, "ctrl/alt + a number from 1 to 9 switches chat tab from 1 to 9"
displaychat strDestTab, strHelpColor, "alt+right or ctrl+tab switches chat tab to the right"
displaychat strDestTab, strHelpColor, "alt+left or ctrl+shift+tab switches chat tab to the left"
displaychat strDestTab, strHelpColor, "Ignore all messages from a player: /i playername or /ignore playername"
displaychat strDestTab, strHelpColor, "F10 to chat in GTA2"
displaychat strDestTab, strHelpColor, vbNullString
displaychat strDestTab, strHelpColor, "Command line arguments of gta2gh.exe: (also works with / instead of -)"
displaychat strDestTab, strHelpColor, "[-s IP]: Let GH join a different IRC server"
displaychat strDestTab, strHelpColor, "[-e [host]]: Join. Automatically joins a game. No argument: join lan game, ID: join game #x, IP: join IP, username: join user"
displaychat strDestTab, strHelpColor, "[-c [map]]: Open Create Game dialog and select a map if specified"
displaychat strDestTab, strHelpColor, "[-h [map]]: Host last played map if no map is specified"
displaychat strDestTab, strHelpColor, "[-p [map]]: Play specified map in single player mode"
displaychat strDestTab, strHelpColor, "[-d [[url1] [url2] [url3] [...]]: Install specified maps"
displaychat strDestTab, strHelpColor, "[-j [[channel1] [channel2] [channel3] [...]]: Install specified maps"
displaychat strDestTab, strHelpColor, "[-q]: exit GH (after doing tasks)"
displaychat strDestTab, strHelpColor, "[-i]: don't open a new instance"
displaychat strDestTab, strHelpColor, "[-l]: do not sign in"
displaychat strDestTab, strHelpColor, "[-?]: display this help"
displaychat strDestTab, strHelpColor, vbNullString
displaychat strDestTab, strHelpColor, "Temporary storage folder: " & GetTmpPath
'displaychat strDestTab, strHelpColor, vbNullString
displaychat strDestTab, strHelpColor, "https://gtamp.com/GTA2/changelog.txt https://gtamp.com/GTA2/todo.txt http://gtamp.com/gh"
End Sub

Public Sub cmdDisconnectClick()
    blnPrivmsg = False
    Call saveChannels
    If sockIRC.State = sckConnected Then send "QUIT :signed out" 'send the quit message
    displaychat strDestTab, strConnectionColor, "Disconnected"
    blnConnected = False
    blnLogin = False
    sockIRC.Close
    blnDisconnectClick = True
    Disconnect
End Sub

Public Sub Disconnect()
    lvPlayers(1).ListItems.Clear
    lvGames(0).ListItems.Clear
    timTimeout.Enabled = False
    blnLogin = False
    cmdToolbar(BTN_SIGN_IN).Enabled = True
    mnuFileSignIn.Enabled = True
End Sub

Private Sub list_of_games(ByVal lstGames As String, ByVal intGameListIndex As Integer) '- join click
On Error GoTo oops:
    Dim i As Integer
    Dim strHostMap As String
    Dim strHostMMP As String
    Dim strHostGMP As String
    Dim strHostSTY As String
    Dim strHostSCR As String
    Dim strHostGHver As String
    
    If tabIRC.SelectedItem.Index <> 1 Then tabIRC.Tabs.Item(1).Selected = True
    
    If lvGames(0).ListItems.count >= intGameListIndex Then
      If lvGames(0).ListItems.Item(intGameListIndex).ListSubItems.count = 7 Then
        strHostCommentLastJoined = lstGames & ": " & lvGames(0).ListItems.Item(intGameListIndex).ListSubItems.Item(7)
      End If
    Else
      displaychat strChannel, strGHColor, "Try again."
    End If

    'check if host is in user list
    blnItemInList = False
    
    For i = 1 To lvPlayers(1).ListItems.count
        Dim strFirstCharName As String
        strFirstCharName = Left$(lvPlayers(1).ListItems.Item(i), 1)
        
        Select Case strFirstCharName
'            Case "@", "+", "~", "&", "%"
'                If lstGames = Mid$(lvPlayers(1).ListItems.Item(i), 2, 255) Then
'                    blnItemInList = True
'                    Exit For
'                End If
            Case Else
                If lstGames = lvPlayers(1).ListItems.Item(i) Then
                    blnItemInList = True
                    Exit For
                End If
        End Select
    Next i

    Call AudioFileCheck

    If lstGames = strNick Then
        frmSettings.main
        Exit Sub
    End If
    
    If lngPID <> 0 And FindProcess(TXT_GTA2EXE, , lngPID) = True Then
        displaychat strChannel, strGHColor, "Close GTA2 before trying to join a game."
        Exit Sub
    End If
   
    Call DetectGTA2version

    If intGameListIndex > lvGames(0).ListItems.count Then
      displaychat strChannel, strTextColor, "Unable to join: Game number " & intGameListIndex & " does not exist.  Try again."
      Exit Sub
    End If

    strMMPfile = lvGames(0).ListItems.Item(intGameListIndex).ListSubItems(3).ToolTipText
    If strMMPfile = vbNullString Then
        displaychat strChannel, strGHColor, "Unable to detect the map filename used by " & lstGames
        Exit Sub
    End If
    
    strHostMMP = lvGames(0).ListItems.Item(intGameListIndex).ListSubItems(3).ToolTipText & ".mmp"
    strHostMap = lvGames(0).ListItems.Item(intGameListIndex).ListSubItems(3).Text
    
    If lvGames(0).ListItems.Item(intGameListIndex).ListSubItems(4).ToolTipText = "Play Replay" Then
        blnPlayReplay = True
    Else
        blnPlayReplay = False
    End If
    
    If Exists(strGTA2path & "data\" & strHostMMP) = False Then
        If blnchkAutoDownload = True Then
            displaychat strChannel, strGHColor, "You don't have " & strMMPfile & ". Attempting to download."
            Call CopyURLToFile("https://gtamp.com/maps/" & LCase(strMMPfile), GetTmpPath & "gta2map.7z")
        Else
            displaychat strChannel, strGHColor, "Try to download map this map from https://gtamp.com/maps/" & LCase(strMMPfile) & " or https://gtamp.com/mapscript/maplist/download.php?mmp=" & Replace(LCase(strMMPfile), " ", "%20") & ".mmp"
        End If
        Exit Sub
    End If
        
    Dim strMMPfullpath As String
    strMMPfullpath = strGTA2path & "data\" & strHostMMP
    With lvGames(0).ListItems.Item(intGameListIndex).ListSubItems(3)
        .Text = readINI("MapFiles", "Description", strMMPfullpath)
    End With
    strHostGMP = readINI("MapFiles", "GMPFile", strMMPfullpath)
    strHostSTY = readINI("MapFiles", "STYFile", strMMPfullpath)
    strHostSCR = readINI("MapFiles", "SCRFile", strMMPfullpath)
    strExecutableChecksum = calc_crc32(strGTA2path & TXT_GTA2EXE)
    strMapChecksum = calc_crc32(strGTA2path & "data\" & strHostGMP)
    strScriptChecksum = calc_crc32(strGTA2path & "data\" & strHostSCR)
  
    strHostNick = lstGames
    
    strHostGHver = lvGames(0).ListItems.Item(intGameListIndex).ListSubItems(4)
    If strHostGHver < "1.513" Then strMMPfile = vbNullString
    
    'Ask for password if game is locked
    If lvGames(0).ListItems.Item(intGameListIndex).ListSubItems(1) = "Yes" Then
        If strHostGHver < "1.513" Then
            displaychat strChannel, strGHColor, "Ask the host to update GH."
        Else
            frmPassword.main (lstGames)
        End If
    Else
        send "NOTICE " & lstGames & " " & "J" & strExecutableChecksum & strMapChecksum & _
            strScriptChecksum & strMMPfile
        displaychat strChannel, strGHColor, "Asking to join " & lstGames
    End If
    
    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error while trying to join: " & strErrdesc & " - Line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :Error while trying to join: " & strErrdesc & " - Line: " & strErrLine
End Sub



'****************************************************************
' lvGH_ColumnClick
' Called when a column header is clicked on - sorts the data in
' that column
'----------------------------------------------------------------
Private Sub lvGames_ColumnClick(Index As Integer, ByVal ColumnHeader As _
                                    MSComctlLib.ColumnHeader)
Call SortLV(frmGH.lvGames(Index), ColumnHeader)
End Sub

Private Sub lvPlayers_ColumnClick(Index As Integer, ByVal ColumnHeader As _
                                    MSComctlLib.ColumnHeader)
Dim i As Integer
'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
For i = 1 To lvPlayers.UBound
    Call SortLV(frmGH.lvPlayers(i), ColumnHeader)
Next i
End Sub

'''FOCUS'''
Private Sub lvGames_GotFocus(Index As Integer)
Call giveChatFocus 'return to chat box, then allow button to do it's thing
End Sub

'''FOCUS'''
Private Sub lvPlayers_GotFocus(Index As Integer)
Call giveChatFocus 'return to chat box, then allow button to do it's thing
End Sub

'''BenMillard''' Changed from _MouseUp to _MouseDown
Private Sub lvPlayers_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
btnMouse = Button 'makes Button available to ItemClick
'Debug.Print "lvPlayers_MouseDown (btnMouse = " & btnMouse & ")"
End Sub

'''BenMillard''' Players list from channels other #gta2gh just won't got away, after clicking a player.
'''(Without moving the mouse during the click.)
Private Sub lvPlayers_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'Debug.Print "lvPlayers_MouseUp(" & Index & "). Visible = " & lvPlayers(Index).Visible
'''lvPlayers(Index).Visible = False 'IS IGNORED!!!
'''Debug.Print "BONKERS!! lvPlayers_MouseUp(" & Index & "). Visible = " & lvPlayers(Index).Visible
End Sub

'''BenMillard''' Trying out the 'correct' method.
Private Sub lvPlayers_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)

On Error GoTo oops

'''FOCUS'''
''''BenMillard''' confirmed _ItemClick fires twice for selected item due to a bug in Common Controls SP5:
'http://support.microsoft.com/kb/257495
'
'Workaround is to raise a flag in the _MouseDown. I re-use btnMouse for that purpose.
'http://www.bigresource.com/Tracker/Track-vb-ETDcxbAg25/
If btnMouse = 0 Then Exit Sub 'was not a genuine click
'''Continue with the subprocedure only if it was a genuine click.

'Debug.Print "lvPlayers_ItemClick (btnMouse = " & btnMouse & ") on: "; Index; Item.Text

Dim strPlayerName As String
Dim intrtbHistory As Integer: intrtbHistory = 1 '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection

'Right-clicked on a player?
If btnMouse = vbRightButton Then
    'Request some info about the player:
    With lvPlayers(Index)
        If .ListItems.count = 0 Then Exit Sub 'if the list is empty then you clicked nothing
        send "WHOIS " & .SelectedItem.Text & " " & .SelectedItem.Text
    End With
    
    '''FOCUS''' Removed from here; handled by _GotFocus instead.
    btnMouse = 0 '''FOCUS''' reset from _MouseDown
    Exit Sub
End If

'''BenMillard''' removed commented-out code: GetAsyncKeyState

'Is the list empty?
If lvPlayers(Index).ListItems.count = 0 Then Exit Sub 'you clicked nothing

'Strip operator codes from player name. There are some other codes used on certain IRC networks.
'I should parse the PREFIX IRC server response to get the real operator codes
strPlayerName = lvPlayers(Index).SelectedItem.Text
Select Case Left$(strPlayerName, 1)
    Case "@", "+", "~", "&", "%"
        strPlayerName = Mid$(strPlayerName, 2, 255)
End Select

'''Did the player click themselves?
If strPlayerName = strNick Then
    '''Toggle status
    Call toggleAwayStatus
    btnMouse = 0 '''FOCUS''' reset from _MouseDown
    Exit Sub '''done!
End If

'Try to select this chat, creating any controls which don't yet exist:
Call ShowChatByPlayerName(strPlayerName)

'''BenMillard''' created ShowChatByPlayerName to centralise and replace the below.
'''BenMillard''' has removed the now obsolete hideWindows and refactored all the code below this.

Exit Sub

oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Player list error: " & Err.Description & " Line: " & strErrLine
    btnMouse = 0 '''reset
End Sub

'''BenMillard''' Changed from _MouseUp to _MouseDown
Private Sub lvGames_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
btnMouse = Button 'makes Button available to ItemClick
End Sub

'''BenMillard''' Trying out the 'correct' method.
'''Private Sub lvGames_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
'''Debug.Print "lvPlayers_ItemClick (btnMouse = " & btnMouse & ") on: ", Index, Item.Text

'Sektor changed back to MouseUp because ItemClick fires twice when running gta2gh.exe
Private Sub lvGames_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo oops
    
    If btnMouse = vbRightButton Then
        With lvGames(Index)
            If .ListItems.count = 0 Then Exit Sub 'if the list is empty then you clicked nothing
            send "WHOIS " & .SelectedItem.Text & " " & .SelectedItem.Text
        End With
        Exit Sub
    End If
    
    If btnMouse = vbLeftButton Then
        If lvGames(Index).ListItems.count = 0 Then Exit Sub 'if the list is empty then you clicked nothing
    
        Call findGTA2
    
        Call DetectGTA2version
    
        If Val(strGTA2version) < 11 Then
            Exit Sub
        End If
        
        With lvGames(Index)
            Call list_of_games(.SelectedItem.Text, .SelectedItem.Index)
        End With
    End If

Exit Sub

oops:
    Call ErrorHandler("lvGames_itemClick", Err.Description, Erl)
End Sub

Private Sub mnuViewHighlight_Click()
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueType = REG_DWORD
        .ValueKey = "chkHighlight"
        
        If mnuViewHighlight.Checked = True Then
            mnuViewHighlight.Checked = False
            .Value = 0
        Else
            mnuViewHighlight.Checked = True
            .Value = 1
        End If
        
        blnHighlight = .Value
    End With
End Sub




'''BenMillard''' refactored tabIRC code


Private Sub tabIRC_GotFocus()
'''Debug.Print "tabIRC_GotFocus", tabIRC.SelectedItem.Index

'Give focus to visible chat box:
Call giveChatFocus

End Sub

'''BenMillard''' wrote this to support multi-line tabs at the top left of a window.
'Here the y-axis checks are removed, so it supports single-row tabs on the left at any vertical position.
'Switch tab immediately:
Private Sub tabIRC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'''Debug.Print "tabIRC_MouseDown"; tabIRC.SelectedItem.Index; x; y

On Error GoTo oops

'Variables:
Dim i As Long

'User left-clicked current tab?
If Button = vbLeftButton Then
    With tabIRC.SelectedItem
        If (.Left < x) And (x < (.Left + .Width)) Then
            'And (.Top < y) And (y < (.Top + .Height)) Then
            Exit Sub
        End If
    End With
End If

'Find which tab was clicked, like .HitTest:
For i = 1 To tabIRC.Tabs.count
    With tabIRC.Tabs(i)
        '''Debug.Print vbTab; x; .Left & " to " & .Left + .Width; (.Left < x) And (x < (.Left + .Width))
        '''Debug.Print vbTab; y; .Top & " to " & .Top + .Height; (.Top < y) And (y < (.Top + .Height))
        If (.Left < x) And (x < (.Left + .Width)) Then
            'And (.Top < y) And (y < (.Top + .Height)) Then
            'Did user right-click the tab?
            If Button = vbRightButton Then
                'Request some information about the current player:
                If Left$(tabIRC.Tabs(i).Caption, 1) <> "#" Then 'is not a channel
                    send "WHOIS " & tabIRC.Tabs(i).Caption & " " & tabIRC.Tabs(i).Caption
                End If
            End If
            
            'Always switch to the tab and update controls:
            Call ShowChatByPlayerName(tabIRC.Tabs(i).Caption)
            Exit For 'we found the tab, so stop looping
        End If
    End With
Next

Exit Sub

oops:
    Call ErrorHandler("tabIRC_MouseDown", Err.Description, Erl)
End Sub

Private Sub tabIRC_Click()
'''Debug.Print "tabIRC_Click: "; tabSelected; tabIRC.SelectedItem.Index; tabIRC.SelectedItem.Caption
'Debug.Print tabIRC.Tabs(tabSelected).Caption

'Selection has changed?
If tabIRC.SelectedItem.Index <> tabSelected Then
    Call ShowChatByPlayerName(tabIRC.SelectedItem.Caption)
End If

End Sub

'Form_KeyDown simply sets .SelectedTab.Index, so relies on _Click updating controls.
'Form_KeyDown deletes the tab but retains the chat history controls.
'Clicking on a tab expects tabIRC_MouseDown to set .SelectedTab and update the controls immediately.
'Avoids double-selecting a tab where possible, as it can flicker slightly.
'displaychat() has been patched, roughly. Needs lots of refactoring now.

'''FOCUS''' Chat includes the following controls:
'Topic    rtbTopic(n)
'History  rtbHistory(n)
'Chat Box rtbChatbox(n)
'Tab      tabIRC.Tabs(n)
'
'When a chat is closed, the tab is removed. The other controls remain, so chat history can be resumed.
'When a tab is removed, all the .Tabs(n) index values are updated by VB6. So they won't match the control array Index values.
'So we must match the player name from .Tabs(n).Caption with the .Tag for the other controls.
'When chatting to a player for the first time, we have to create the other controls - there won't be a match.
'modCommand.processCommand calls this once, after you /JOIN a channel and receive the /names list: Case "353"
Public Sub ShowChatByPlayerName(PlayerName As String)
'Debug.Print "ShowChatByPlayerName: ", PlayerName

'Variables:
Dim i As Long
Dim ChatControlsID As Long 'chat history found here while looping, or created here
Dim ChatTabID As Integer 'chat tab found here while looping, or created here

On Error GoTo oops

'Chat is fully available AND already selected:
If UCase(tabIRC.SelectedItem.Caption) = UCase(PlayerName) Then Exit Sub

'Is chat already available from a tab?
For i = 1 To tabIRC.Tabs.count
    '''Debug.Print vbTab & "tabIRC.Tabs(" & i & "):", tabIRC.Tabs(i).Caption
    If UCase(tabIRC.Tabs(i).Caption) = UCase(PlayerName) Then
        tabIRC.Tabs(i).Caption = PlayerName 'update the caption just in case the casing is different
        ChatTabID = i
        Exit For 'found it, stop looping
    End If
Next i

'Chat has no tab. Create it:
If ChatTabID = 0 Then
    '(Copied from lvPlayers_ItemClick)
    tabIRC.Tabs.Add , , PlayerName 'start using the key as well?
    ChatTabID = tabIRC.Tabs.count 'new index
End If
Debug.Print vbTab & "ChatTabID:", ChatTabID

'Is chat history available?
For i = 1 To rtbHistory.UBound
    '''Debug.Print vbTab & "rtbHistory(" & i & "):", rtbHistory(i).Tag
    If UCase(rtbHistory(i).Tag) = UCase(PlayerName) Then
        ChatControlsID = i
        Exit For 'found it, stop looping
    End If
Next i

'Chat has no tab and there is no chat history. Create a new chat:
If ChatControlsID = 0 Then
    ChatControlsID = rtbHistory.UBound + 1 'new index
    Load rtbTopic(ChatControlsID)
    Load rtbHistory(ChatControlsID)
    Load rtbChatbox(ChatControlsID)
    rtbTopic(ChatControlsID).Text = TXT_PRIVATE & PlayerName
    rtbTopic(ChatControlsID).Tag = PlayerName
    rtbTopic(ChatControlsID).Locked = True
    '''BenMillard''' says remove commented-out code: 'rtbTopic(rtbTopic.UBound).ToolTipText = strPlayerName
    rtbHistory(ChatControlsID).Text = vbNullString
    rtbHistory(ChatControlsID).Tag = PlayerName
    '''BenMillard''' says remove commented-out code: 'rtbHistory(ChatControlsID).ToolTipText = strPlayerName
    '''BenMillard''' says remove commented-out code: '''EnableAutoURLDetection rtbHistory(ChatControlsID)
    
    With rtbChatbox(ChatControlsID)
        .Tag = PlayerName
        '''BenMillard''' says remove commented-out code: '.ToolTipText = strPlayerName
        .Text = vbNullString
        .SelStart = 0
        .SelLength = 500
        .SelColor = strTextColor
        .SelLength = 0
    End With
End If
Debug.Print vbTab & "ChatControlsID:", ChatControlsID

'Chat is now ready to be selected:
Call SelectTab(ChatTabID, , ChatControlsID)

Exit Sub

oops:
    Call ErrorHandler("showChatByPlayerName", Err.Description, Erl)
End Sub

''''FOCUS''' View the controls for the selected tab.
'Also, this updated the .SelectedItem on _MouseDown. Otherwise it only updates after _Click.
Private Sub SelectTab(Index As Integer, Optional IsCyclic = False, Optional ChatControlsID As Long = 0)
'Debug.Print "SelectTab(" & Index & "), switching from:", tabIRC.SelectedItem.Index ', IsCyclic, ChatControlsID
'Uses tabSelected to see which tab is visually selected.
'IsCycling can be used to centralise Next/Previous tab code.
'ChatControlsID helps link the chat controls to the tab we want to select.

On Error GoTo oops

'Variables:
Dim i As Long

'Trying to select a tab which doesn't exist?
If Index < 1 Then 'too low:
    If IsCyclic Then
        Index = tabIRC.Tabs.count 'last
    Else
        Index = 1 'first
    End If
ElseIf Index > tabIRC.Tabs.count Then 'too high:
    If IsCyclic Then
        Index = 1 'first
    Else
        Index = tabIRC.Tabs.count 'last
    End If
End If

'Tab is already selected?
If Index = tabSelected Then Exit Sub

'Now we know which tab should be selected:
tabSelected = Index
Set tabIRC.SelectedItem = tabIRC.Tabs(tabSelected)
'Debug.Print vbTab & "SelectTab will select: ", Index

'Find which chat controls should be displayed:
If ChatControlsID < 1 Then 'hasn't been set
    'Search for the chat controls:
    For i = 1 To rtbTopic.UBound
        If rtbHistory(i).Tag = tabIRC.Tabs(tabSelected).Caption Then
            ChatControlsID = i
            Exit For 'found it, stop looping
        End If
    Next i
End If

'Loop through all controls related to tabs, only showing the relevant set:
'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
For i = 1 To rtbTopic.UBound
    rtbTopic(i).Visible = (i = ChatControlsID)
    rtbHistory(i).Visible = (i = ChatControlsID)
    rtbChatbox(i).Visible = (i = ChatControlsID)
    '''Debug.Print vbTab & "rtbTopic(" & i & ").Tag = " & rtbTopic(i).Tag
Next i

'Show the corresponding Players list:
Call ShowPlayerList(tabSelected, ChatControlsID)

'Allow user to close any tab, but not the first one:
cmdX.Enabled = Not (tabIRC.Tabs(tabSelected).Index = 1)

'Remove highlight from the tab which is now selected:
tabIRC.Tabs(tabSelected).HighLighted = False

'Give focus to visible chat box:
Call giveChatFocus

'Redraw controls to be the correct size:
Call Form_Resize
Debug.Print vbNewLine '''seperate subsequent events

Exit Sub

oops:
    Call ErrorHandler("selectTab", Err.Description, Erl)
End Sub

''''FOCUS''' Show the list of players for this chat or channel.
'SelectTab() is the only place this is called from, at the moment.
Sub ShowPlayerList(tabSelected As Long, ChatControlsID As Long)
'Debug.Print "ShowPlayerList()"

On Error GoTo oops

'Variables:
Dim i As Long
Dim PlayerListID As Long 'control array Index for the corresponding Players list

'No tab specified? Let's find it ourselves:
tabSelected = tabIRC.SelectedItem.Index

'Is the chat for a channel or private with a player?
If Left$(tabIRC.Tabs(tabSelected).Caption, 1) = "#" Then
    'It is a channel:
    PlayerListID = getPlayerLV(tabIRC.Tabs(tabSelected).Caption)
    'Debug.Print vbTab & "Viewing channel: ", PlayerListID, tabIRC.Tabs(tabSelected).Caption
    rtbTopic(ChatControlsID).Locked = False '''unlock to allow editing (server checks permission)
Else
    'Not a channel, therefore it's a private chat:
    PlayerListID = 1 'show the default #gta2gh players list
    '                      ^-- Remove this line to hide player list in private chat.
    rtbTopic(ChatControlsID).Text = TXT_PRIVATE & tabIRC.SelectedItem.Caption
    
    '''BenMillard''' moved this loop to other side of Else.
    'Select the corresponding Player in the Players list:
    For i = 1 To lvPlayers(1).ListItems.count
        If UCase(tabIRC.Tabs(tabSelected).Caption) = UCase(lvPlayers(1).ListItems(i).Text) Then
            lvPlayers(1).ListItems.Item(i).Selected = True
            Exit For
        End If
    Next i
End If

'Show only the corresponding Players list for the selected channel:
If PlayerListID > 0 Then 'we found the Players list
    For i = 1 To lvPlayers.UBound
        'Debug.Print vbTab & vbTab & i, lvPlayers(i).Tag, (i = PlayerListID)
        lvPlayers(i).Visible = (i = PlayerListID)
        If Not (i = PlayerListID) Then 'invisible control
            lvPlayers(i).ZOrder 1 'Send to Back, so it can't be seen even if it becomes visible again!
        End If
    Next i
End If

Exit Sub

oops:
    Call ErrorHandler("showPlayerList", Err.Description, Erl)
End Sub
'
'
'
'''BenMillard''' refactored tabIRC code ends.



'Networking

'Private Sub sckURL_Connect()
'On Error Resume Next
'    'If strFailedCountryIP = vbNullString Then
'      'http://api.hostip.info/country.php
'      'sckURL.SendData "GET /country.php HTTP/1.0" & vbCrLf
'      'sckURL.SendData "Host: api.hostip.info" & vbCrLf
'      'sckURL.SendData vbCrLf
'
'      sckURL.SendData "GET /?self HTTP/1.0" & vbCrLf
'      sckURL.SendData "Host: geoloc.daiguo.com" & vbCrLf
'      sckURL.SendData vbCrLf
'      '    Else
'      '      sckURL.SendData "GET /?ip=" & strFailedCountryIP & " HTTP/1.0" & vbCrLf
'      '      sckURL.SendData "Host: geoloc.daiguo.com" & vbCrLf
'      '      sckURL.SendData vbCrLf
'    'End If
'End Sub


Public Sub getCC()
On Error GoTo oops
    Dim strTemp As String
    Dim i As Integer
    Dim strTempCC As String
    
    'If sckURL.State = sckConnected Then sckURL.GetData strTemp, vbString, 666
    
    strTemp = CopyURLToRAM(TXT_GEOSITE)
    
    'search the response for 1; and the next two characters should be the country code
    i = InStr(strTemp, "1;")
    
    If i = 0 Then
        displaychat strDestTab, strConnectionColor, TXT_COUNTRY_DETECTION_FAILED
        'sckURL.Close
        blnCountryDetectFail = True
        Exit Sub
    End If
        
    If i > 0 Then strTempCC = Mid$(strTemp, i + 2, 2)
      
    'sckURL.Close
      
    Select Case strTempCC
        Case "GB"
            strTempCC = "UK"
        Case "EU"
            blnCountryDetectFail = True
    End Select
    
    For i = 0 To UBound(strCountries)
        If strTempCC = Right$(strCountries(i), 2) Then
            If blnCountryDetectFail = False Then
                displaychat strDestTab, strConnectionColor, vbNullString & "Country detected as " & Left$(strCountries(i), Len(strCountries(i)) - 5)
            End If
            
            strCountryCode = strTempCC
            intCountryIndex = i
            
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
    Exit Sub
oops:
    displaychat strChannel, strGHColor, "getCC error"
End Sub

'Private Sub sckURL_DataArrival(ByVal BytesTotal As Long)
'On Error GoTo oops
'    Dim strTemp As String
'    Dim i As Integer
'    Dim strTempCC As String
'
'    If sckURL.State = sckConnected Then sckURL.GetData strTemp, vbString, 666
'
'    strTemp = CopyURLToRAM(TXT_GEOSITE)
'
'    'search the response for 1; and the next two characters should be the country code
'    i = InStr(strTemp, "1;")
'
'    If i = 0 Then
'        displaychat strDestTab, strConnectionColor, TXT_COUNTRY_DETECTION_FAILED
'        sckURL.Close
'        blnCountryDetectFail = True
'        Exit Sub
'    End If
'
'    If i > 0 Then strTempCC = Mid$(strTemp, i + 2, 2)
'
'    sckURL.Close
'
'    Select Case strTempCC
'        Case "GB"
'            strTempCC = "UK"
'        Case "EU"
'            blnCountryDetectFail = True
'    End Select
'
'    For i = 0 To UBound(strCountries)
'        If strTempCC = Right$(strCountries(i), 2) Then
'            If blnCountryDetectFail = False Then
'                displaychat strDestTab, strConnectionColor, vbNullString & "Country detected as " & Left$(strCountries(i), Len(strCountries(i)) - 5)
'            End If
'
'            strCountryCode = strTempCC
'            intCountryIndex = i
'
'            With cr
'                .ClassKey = HKEY_CURRENT_USER
'                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
'                .ValueType = REG_SZ
'                .ValueKey = "Country"
'                .Value = strCountryCode
'            End With
'            Exit For
'        End If
'    Next i
'    Exit Sub
'oops:
'    displaychat strChannel, strGHColor, "sckURL_DataArrival error"
'End Sub

'Private Sub sckURL_error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    displaychat strDestTab, strConnectionColor, TXT_COUNTRY_DETECTION_FAILED & ": " & Description
'    displaychat strDestTab, strConnectionColor, "Failed to connect to: " & TXT_GEOSITE
'    cmdToolbar(BTN_SIGN_IN).Enabled = True
'    mnuFileSignIn.Enabled = True
'    blnCountryDetectFail = True
'End Sub

Private Sub sockIRC_Connect()   'as soon as we're connected to the server:
    Dim strIdent As String
    blnConnected = True    'set connected to true (cancel the timeout procedure)
    blnConnectClick = False
    
    If strPassword = vbNullString Then strPassword = "x"
    send "PASS " & strPassword
    send "NICK " & strPreferedNick 'send the nick message
    'USER <username> <hostname> <servername> <real name>
    If strCountryCode = vbNullString Then strCountryCode = "??"
    Call GetMACs_AdaptInfo
    strIdent = strMacAddress
    If strPreferedNick = "Sektor" Then strIdent = "admin"
    If strIdent = vbNullString Then strIdent = strPreferedNick
    send "USER " & strIdent & " " & sockIRC.LocalIP & _
        " GTA2GameHunter :GH" & TXT_GHVER & strCountryCode
End Sub

Private Sub sockIRC_error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'If Description = "Connection is forcefully rejected" Then
        Select Case Number
            Case "11001", "11004"
                displaychat strDestTab, strConnectionColor, "Failed to resolve host: " & strServer(0)
            Case "10061"
                displaychat strDestTab, strConnectionColor, "Unable to connect to server"
            Case Else
                displaychat strDestTab, strConnectionColor, Number & " " & Description
        End Select
        blnDisconnectClick = False
        timTimeout.Enabled = False
        sockIRC.Close
        'Call changeServer
        Call Disconnect
        If intServerNum = 1 Then
            blnReconnect = True
        Else
            Call Reconnect
        End If
    'End If
End Sub

Private Sub sockIRC_DataArrival(ByVal BytesTotal As Long)
On Error GoTo oops
    Dim x As Long
    For x& = 1 To BytesTotal    'get every byte we received, but only one at a time
        Dim strTemp As String   'this variable will be used to store one byte of data
        If sockIRC.State = sckConnected Then sockIRC.GetData strTemp, vbString, 1  'get 1 byte out of the data stream and store it in strTemp
        If strTemp = Chr$(10) Then    'if we received a newline character (this is the end of the message)
            processCommand  'process the entire command
            strData = vbNullString      'clear the strData
        End If
        'bug
        If strTemp <> vbNullString Then
            If Not (Asc(strTemp) = 10 Or Asc(strTemp) = 13) Then strData = strData & strTemp
        End If
            'if we received a newline character or a carriage return, dont add them to the message
    Next

    Exit Sub

oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, strTextColor, "error during data arrival: " & strErrLine & " " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :DA" & strErrLine & " " & Erl
    cmdDisconnectClick
End Sub


Private Sub timStamp_Timer()
    Dim strText As String
    Dim strReplace As String
    Dim lngLen As Long
    'if GTA2 isn't hosting then use the join history handle
    If blnReadyForJoiners = False Then lngHandleHistory = lngHandleJoinHistory
    'Read GTA2 chat history into strText
    lngLen = SendMessage(lngHandleHistory, WM_GETTEXTLENGTH, 0&, 0&) + 1
    strText = Space(lngLen)
    SendMessage lngHandleHistory, WM_GETTEXT, lngLen, ByVal strText
    'If there's no history then do nothing
    If strText <> Chr$(0) And strText <> vbNullString Then
        Dim i As Long
        'Search the history for a line break
        i = InStrRev(strText, vbCrLf)
        If i Then
            'Check if there is a timestamp (just checks if first char is a number)
            'If IsNumeric(Mid$(strText, i + 2, 1)) = False Then
            'Korean timestamp doesn't start with a number
            'Changed to check for a colon with a number next to it
            If IsNumeric(Mid$(strText, InStr(i, strText, ":") + 1, 1)) = False Then
                'There's no timestamp, so add one
                strReplace = Replace(strText, vbCrLf, vbCrLf & Time & " ", i)
                strText = Left$(strText, i) & strReplace
                'Scroll to the of the history
                SendMessage lngHandleHistory, WM_SETTEXT, 0, ByVal strText
                SendMessage lngHandleHistory, WM_VSCROLL, SB_BOTTOM, ByVal 0&
            End If
        Else
            'No line break, so we can just replace the whole strText with time & strText
            'This only occurs for the first chat line
            'If IsNumeric(Left$(strText, 1)) = False Then
            If IsNumeric(Mid$(strText, InStr(strText, ":") + 1, 1)) = False Then
                strText = Time & " " & strText
                SendMessage lngHandleHistory, WM_SETTEXT, 0, ByVal strText
                SendMessage lngHandleHistory, WM_VSCROLL, SB_BOTTOM, ByVal 0&
            End If
        End If
    End If
End Sub

'Timers



Private Sub timTimeout_Timer()
    If Not (blnConnected) Then
        displaychat strDestTab, strConnectionColor, "The connection to the server timed out"
        blnDisconnectClick = False
        timTimeout.Enabled = False
        sockIRC.Close
        'Call changeServer
        Call Disconnect
        Reconnect
    End If
End Sub

'Public Sub changeServer()
'    If InStr(strServer(0), "127.0.0.1") Then Exit Sub
'    If blnDONOTCHANGESERVER = True Then
'        blnDONOTCHANGESERVER = False
'        Exit Sub
'    End If
'    intServerNum = intServerNum + 1
'    If intServerNum > UBound(strServer) Then intServerNum = 0
'    If strServer(intServerNum) = vbNullString Then intServerNum = 0
'    displaychat strDestTab, strConnectionColor, "Server changed to: " & strServer(intServerNum)
'End Sub

'keeps checking to see if host has changed map by reading map_index
Private Sub timUpdateMap_Timer()
Dim i As Long

On Error GoTo oops

If blnReadyForJoiners = False Then Exit Sub

'Read map index from registry
With cr
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "Software\DMA Design Ltd\GTA2\Network"
    .ValueKey = "map_index"
    i = .Value
End With

If i >= SortedArray.count Then i = 0

If intPreviousMapIndex <> i Then
    If SortedArray.count > 0 Then strGTA2MapDesc = SortedArray.SortedItem(i)
    intPreviousMapIndex = i
End If

Call AdvertiseHostedGame

Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, strTextColor, "error during timUpdateMap: " & strErrLine & " " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :error during timUpdateMap: " & strErrLine & " " & strErrdesc

End Sub

Private Sub rtbTopic_Change(Index As Integer)

On Error GoTo oops:

''Remove any styling from Drag & Drop:
Dim strTemp As String, lngSelStart As Long, lngSelLength As Long 'store state

'Restore the text without its formatting:
With rtbTopic(Index)
    'Store:
    lngSelStart = .SelStart
    lngSelLength = .SelLength
    strTemp = .Text

    'Clear:
    .TextRTF = vbNullString

    'Restore:
    .Text = strTemp
    .SelStart = 0
    .SelLength = 9999
    .SelColor = strTextColor
    .SelUnderline = False '''BenMillard''' added this to remove underline bug
    .SelStart = lngSelStart
    .SelLength = lngSelLength
    
    'highlight URLs
    Dim strArray() As String
    Dim i As Integer
    Dim j As Integer
    Dim strClean As String
    
    strArray = Split(strTemp)
    
    For i = 0 To UBound(strArray)
        
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
        
        j = j + Len(strArray(i)) + 1
    Next
    
    .SelStart = lngSelStart
    
End With

Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, strTextColor, "error during topic change: " & strErrLine & " " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :topic_change: " & strErrLine & " " & strErrdesc
End Sub

Private Sub rtbChatbox_change(Index As Integer)
''Remove any styling from Drag & Drop:
Dim strTemp As String, lngSelStart As Long, lngSelLength As Long 'store state
'Restore the text without its formatting:
With rtbChatbox(Index)
    'Store:
    lngSelStart = .SelStart
    lngSelLength = .SelLength
    strTemp = .Text

    'Clear:
    .TextRTF = vbNullString

    'Restore:
    .Text = strTemp
    .SelStart = 0
    .SelLength = 9999
    .SelColor = strTextColor
    .SelStart = lngSelStart
    .SelLength = lngSelLength
    
End With

End Sub

'''FOCUS'''
'Private Sub rtbChatbox_LostFocus(Index As Integer)
''Debug.Print "LostFocus: rtbChatbox"; Index
''Debug.Print " NowFocus: "; ActiveControl.Name
''DoEvents
''Debug.Print "GiveFocus: rtbChatbox"; Index
''rtbChatbox(Index).SetFocus
'End Sub
'
'Private Sub rtbChatbox_GotFocus(Index As Integer)
''Debug.Print " GotFocus: rtbChatbox"; Index
'End Sub

Private Sub rtbChatbox_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case "13" 'enter
            If Left$(tabIRC.SelectedItem.Caption, 1) <> "#" Then
                'only send a message if the user is on the same channel as you
                If InStr(LCase(tabIRC.SelectedItem.Caption), "serv") = False Then
                    If onYourChannel(tabIRC.SelectedItem.Caption) = False Then
                        displaychat tabIRC.SelectedItem.Caption, strGHColor, "You must be in the same channel as " & tabIRC.SelectedItem.Caption & " to send them a message."
                        Exit Sub
                    End If
                End If
            End If
            cmdChat_Click
            KeyAscii = 0
    End Select
'Messages sent to gtanet can only be around 450 characters
'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
'        'If Len(rtbChatbox(1).Text) > 450 Then KeyAscii = 0
End Sub

Private Sub rtbTopic_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case KeyAscii
        Case "13" 'enter
            send "TOPIC " & rtbTopic(Index).Tag & " :" & rtbTopic(Index).Text
            KeyAscii = 0
    End Select
'Messages sent to gtanet can only be around 450 characters
'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
'        'If Len(rtbChatbox(1).Text) > 450 Then KeyAscii = 0
End Sub


Private Sub Form_Resize()
  On Error GoTo oops:
    Dim i As Integer
    Dim j As Integer
    
    If WindowState <> vbMinimized Then intPrevWinState = WindowState
    
    Dim intPlayerList As Integer
    
    Select Case frmGH.WindowState
        Case vbNormal
            If frmGH.Height < 5000 Then frmGH.Height = 5000
            If frmGH.Width < 8000 Then frmGH.Width = 8000
        Case vbMinimized
            If blnchkMinToTray = True Then
                Hide
                Call Systray
                Exit Sub
            End If
    End Select
 
    If frmGH.Width > 7000 And frmGH.Height > 4000 Then
        If WindowState = vbMaximized Then
            cmdX.Left = ScaleWidth - 370
        Else
            cmdX.Left = ScaleWidth - 370
        End If
        
        '''Set control positions and sizes:
        Dim intButtonOffset As Integer
        Dim intBannerSpace As Integer
        'Dim bannerStyle As Integer
        
        '''BenMillard''' removed commented-out code for bannerStyle
        
        tabIRC.Top = ScaleHeight - tabIRC.Height '- 30 'usually 15 twips per pixel
        
        For i = 1 To tabIRC.Tabs.count
            j = j + tabIRC.Tabs(i).Width
        Next
        
        If j >= tabIRC.Width Then
            tabIRC.Width = ScaleWidth - 500
        Else
            tabIRC.Width = ScaleWidth
        End If
        lblHide.Top = ScaleHeight - tabIRC.Height '- 30 'usually 15 twips per pixel
        lblHide.Height = tabIRC.Height
        lblHide.Width = 900
        lblHide.Left = ScaleWidth - 900
                
        'need to use j instead of i because i gets a 1 added to it???
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        j = 1
        For i = 1 To rtbTopic.UBound
            If LCase(tabIRC.SelectedItem.Caption) = LCase(rtbTopic(i).Tag) Then
                j = i
                Exit For
            End If
        Next
        
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        If Left$(tabIRC.SelectedItem.Caption, 1) = "#" Then
            intPlayerList = getPlayerLV(tabIRC.SelectedItem.Caption)
            If intPlayerList = -1 Then Exit Sub
        Else
            intPlayerList = 1
        End If
        
        'Toolbar buttons:
        cmdToolbar(BTN_SIGN_IN).Left = 0 + intButtonOffset
        cmdToolbar(BTN_CREATE).Left = 1320 + intButtonOffset
        cmdToolbar(BTN_OPTIONS).Left = 2640 + intButtonOffset
        cmdToolbar(BTN_SIGN_OUT).Left = 3960 + intButtonOffset
        'cmdEnter.Left = 5280 + intButtonOffset 'Elypter's code
        cmdToolbar(BTN_MANAGER).Left = 5280 + intButtonOffset
        cmdToolbar(BTN_CANCEL).Left = 6600 + intButtonOffset
        
        'Elypter's code for top banner commented out:
        'cmdToolbar(BTN_SIGN_IN).Top = 0 + intBannerSpace
        'cmdToolbar(BTN_CREATE).Top = 0 + intBannerSpace
        'cmdToolbar(BTN_OPTIONS).Top = 0 + intBannerSpace
        'cmdToolbar(BTN_SIGN_OUT).Top = 0 + intBannerSpace
        'cmdEnter.Top = 0 + intBannerSpace 'Elypter's code
        'cmdToolbar(BTN_MANAGER).Top = 0 + intBannerSpace
        'cmdToolbar(BTN_CANCEL).Top = 0 + intBannerSpace
        lvGames(0).Top = 360 '+ intBannerSpace
        lvPlayers(intPlayerList).Top = 360 '+ intBannerSpace
        
        'Players list:
        lvPlayers(intPlayerList).Left = ScaleWidth - lvPlayers(intPlayerList).Width
        lvPlayers(intPlayerList).Height = ScaleHeight - 1140 - intBannerSpace
        If lvPlayers(intPlayerList).ListItems.count Then lvPlayers(intPlayerList).SelectedItem.EnsureVisible
        
        'Games list:
        lvGames(0).Width = lvPlayers(intPlayerList).Left + 10
        
        'Chat controls:
        rtbTopic(j).Top = lvGames(0).Top + lvGames(0).Height
        rtbTopic(j).Width = lvGames(0).Width
        rtbChatbox(j).Top = tabIRC.Top - rtbChatbox(j).Height + 30 'hide top border of tabbed area
        rtbChatbox(j).Width = ScaleWidth
        rtbHistory(j).Top = rtbTopic(j).Top + rtbTopic(j).Height
        rtbHistory(j).Width = lvGames(0).Width
        rtbHistory(j).Height = rtbChatbox(j).Top - rtbHistory(j).Top
        
'        'Hide player list in private tab
'        Else
'            lvGames(0).Width = ScaleWidth
'            rtbTopic(j).Width = ScaleWidth
'            rtbChatbox(j).Top = tabIRC.Top - rtbChatbox(j).Height + 30 'hide top border of tabbed area
'            rtbChatbox(j).Width = ScaleWidth
'            rtbHistory(j).Top = rtbTopic(j).Top + rtbTopic(j).Height
'            rtbHistory(j).Width = ScaleWidth
'            rtbHistory(j).Height = rtbChatbox(j).Top - rtbHistory(j).Top
'        End If
    End If
    
    If lvGames(0).ListItems.count Then lvGames(0).SelectedItem.EnsureVisible
    
    'cmdP.Top = tabIRC.Tabs(1).Top + 30 'Elypter's code
    'cmdP.Height = 260 'Elypter's code
    
    cmdX.Top = tabIRC.Tabs(1).Top + 30
    cmdX.Height = 260
    
    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "error during resize: " & strErrdesc & " , line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :error during resize: " & strErrdesc & " " & strErrLine
End Sub

Private Sub AutoSizeGames()
On Error GoTo oops
'''Sets the Width of columns in lvGames
'''See "columns.txt" for facts about why and how much padding is needed.
'''Turns out be really fiddly but this sets column Widths perfectly!
'''Print "AutoSizeColumns(" & Index & ")"

'Variables:
Dim sngWidth As Single, sngHeight As Single 'total width of columns, total height of items
Dim strHeader As String, lngValueStart As Long, lngValueEnd As Long 'get count value from 1st column header

With lvGames(0)
    'Store count displayed by 1st column header:
    strHeader = .ColumnHeaders.Item(1).Text
    lngValueStart = InStr(1, strHeader, "(") + 1
    lngValueEnd = InStr(1, strHeader, ")")
    '''Debug.Prin lngValueStart; lngValueEnd
    If lngValueStart And lngValueEnd Then 'has a value:
        'Is value different from actual count?
        If Mid$(strHeader, lngValueStart, lngValueEnd - lngValueStart) <> .ListItems.count Then
            'Update column header:
            .ColumnHeaders.Item(1).Text = "Games ("
            .ColumnHeaders.Item(1).Text = .ColumnHeaders.Item(1).Text & .ListItems.count & ")" 'count
            ShowListViewColumnHeaderSortIcon lvGames(0) 'restore sorting arrow
        End If
    End If
    
    '''Actual sizing is abstracted to modSizeLV.AutoSizeLV
    sngWidth = AutoSizeLV(lvGames(0))
    
    'Games list changes height to suit number of games
    'No games?
    'txtNoGames.Visible = (.ListItems.Count < 1)
    
    'Less than 3 games?
    If .ListItems.count < 3 Then
        sngHeight = 735 '''840 'default
    Else
        'Headers, scrollbar and breathing room:
        sngHeight = (210 * .ListItems.count) + 255 + 60
        '210 for normal row; 255 for headers row; 60 for 3D borders
    End If
    
    'Horizontal scrollbar?
    sngWidth = sngWidth + 60 'add outer borders and padding
    If sngWidth > .Width Then 'horizontal scrollbar
        sngHeight = sngHeight + 255 'space for scrollbar
    End If
    
    'Apply height only if it will be different:
    If .Height <> sngHeight Then 'heights are always Round() due to above
        .Height = sngHeight
        Call Form_Resize 'adapt other controls
    End If

End With

Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, strTextColor, "AutoSizeColumns error: Line " & strErrLine & " " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :AutoSizeColumns error: Line " & strErrLine & " " & strErrdesc
End Sub

Private Sub AutoSizePlayers()
On Error GoTo oops
'''Sets the Width of columns in lvPlayers

'Variables:
Dim sngWidth As Single 'total width of columns
Dim strHeader As String, lngValueStart As Long, lngValueEnd As Long 'get count value from 1st column header
Dim strChannelCaption As String
Dim intPlayerList As String

strChannelCaption = frmGH.tabIRC.SelectedItem.Caption
intPlayerList = getPlayerLV(strChannelCaption)
If intPlayerList = -1 Then Exit Sub

With lvPlayers(intPlayerList)
    'Store count displayed by 1st column header:
    strHeader = .ColumnHeaders.Item(1).Text
    lngValueStart = InStr(1, strHeader, "(") + 1
    lngValueEnd = InStr(1, strHeader, ")")
    '''Debug.Prin lngValueStart; lngValueEnd
    If lngValueStart And lngValueEnd Then 'has a value:
        'Is value different from actual count?
        If Mid$(strHeader, lngValueStart, lngValueEnd - lngValueStart) <> .ListItems.count Then
            'Update column header:
            .ColumnHeaders.Item(1).Text = "Players ("
            .ColumnHeaders.Item(1).Text = .ColumnHeaders.Item(1).Text & .ListItems.count & ")" 'count
            ShowListViewColumnHeaderSortIcon lvPlayers(intPlayerList) 'restore sorting arrow
        End If
    End If

    '''Actual sizing is abstracted to modSizeLV.AutoSizeLV
    sngWidth = AutoSizeLV(lvPlayers(intPlayerList))
  
    'Players list. If it changes width, other controls adapt to it.
    'Is showing a vertical scrollbar?
    ''''Rebug.Print "Players list: "; .Height; .ListItems.Count; (.ListItems.Count * 210) + 255 + 60
    If .Height < (.ListItems.count * 210) + 255 + 60 Then 'more items than can be shown:
        sngWidth = sngWidth + 255 'typical scrollbar width
    End If
    
    'Apply width only if it will be different:
    sngWidth = sngWidth + 60 'add borders
    If CLng(.Width) <> CLng(sngWidth) Then
        .Width = sngWidth
        Call Form_Resize 'adapt other controls
    End If

End With

Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, strTextColor, "AutoSizePlayers error: Line " & strErrLine & " " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :AutoSizePlayers error: Line " & strErrLine & " " & strErrdesc
End Sub

Private Sub timStatus_Timer()
    On Error GoTo oops
    
    Static intTimerError As Integer
    
    Dim strState As String
    Dim i As Integer
    Dim blnSpecialDay As Boolean
      
    
    'April Fool's day
    If Month(Now) = 4 And Day(Now) = 1 Then
        If frmGH.Icon <> picJoke.Picture Then frmGH.Icon = picJoke.Picture
        If blnAprilFools = False Then
            blnAprilFools = True
            frmGH.Caption = "Game Hunter v" & Int((Rnd * 99) + 1) & "." & Int((Rnd * 99999) + 1)
        End If
        blnSpecialDay = True
    End If
    
    'GH 1.0 2005-01-24
    If Month(Now) = 1 And Day(Now) = 24 Then
        frmGH.Caption = "Game Hunter v" & TXT_GHVER & " - GH is " & Year(Now) - 2005 & " years old and you're still using it!"
        If frmGH.Icon <> frmAbout.cmdOK.MouseIcon Then frmGH.Icon = frmAbout.cmdOK.MouseIcon
        blnSpecialDay = True
    End If
    
    'Sektor's birthday 1979-07-18
    If Month(Now) = 7 And Day(Now) = 18 Then
        frmGH.Caption = "Game Hunter v" & TXT_GHVER & " - Sektor was created " & Year(Now) - 1979 & " years ago!"
        If frmGH.Icon <> frmAbout.cmdOK.MouseIcon Then frmGH.Icon = frmAbout.cmdOK.MouseIcon
        blnSpecialDay = True
    End If
    
    'GTA2 release anniversary 1999-10-22
    If Month(Now) = 10 And Day(Now) = 22 Then
        frmGH.Caption = "Game Hunter v" & TXT_GHVER & " - GTA2 is " & Year(Now) - 1999 & " years old!"
        'If frmGH.Icon <> picHead.Picture Then frmGH.Icon = picHead.Picture
        If frmGH.Icon <> frmAbout.cmdOK.MouseIcon Then frmGH.Icon = frmAbout.cmdOK.MouseIcon
        blnSpecialDay = True
    End If
    
    'Happy Halloween
    If Month(Now) = 10 And Day(Now) = 31 Then
        frmGH.Caption = "Game Hunter v" & TXT_GHVER & " - Happy Halloween!"
        If frmGH.Icon <> picHalloween.Picture Then frmGH.Icon = picHalloween.Picture
        blnSpecialDay = True
    End If
    
    'Happy Holidays
    If Month(Now) = 12 And Day(Now) = 25 Then
        If frmGH.Icon <> picSanta.Picture Then frmGH.Icon = picSanta.Picture
        frmGH.Caption = "Game Hunter v" & TXT_GHVER & " - Happy Holidays!"
        blnSpecialDay = True
    End If
    
    If blnSpecialDay = False Then
        If frmGH.Icon <> picGH.Picture Then frmGH.Icon = picGH.Picture
    End If

    If blnReconnect = True Then
        intReconnect = intReconnect + 1
        If intReconnect > 3 Then
            intReconnect = 0
            Call Reconnect
            blnReconnect = False
        End If
    End If
    
    If blnReadyForJoiners = False Then Call RemovePlayersGameFromList
    
    'This is used to join the channel after 5 seconds if there is no message from nickserv about the nick being in use
'        If intNickservWaitTime <> -1 Then
'            intNickservWaitTime = intNickservWaitTime + 1
'           If intNickservWaitTime = 5 Then
'               intNickservWaitTime = -1
'               send "JOIN " & strChannel & " " & strKey
'           End If
'       End If
    
    'Some sounds shouldn't play when GH has focus
'    If GetActiveWindow <> 0 Then
'        blnFocus = True
'    Else
'        blnFocus = False
'    End If
    
    If Visible = False And blnSystray = True Then Call Systray 'redraw the systray just in case explorer crashed

    If blnDisconnectClick = False Then
        If intTimeSinceLastServerData > 700 Then
            displaychat strDestTab, vbRed, "Server timed out.  Reconnecting..."
            intTimeSinceLastServerData = 0
            Reconnect
        Else
            intTimeSinceLastServerData = intTimeSinceLastServerData + 1
            Select Case intTimeSinceLastServerData
                Case 200, 400, 600
                    'displaychat strDestTab, strGHColor, "No activity from server for " & intTimeSinceLastServerData & " seconds. Sending connection test command."
                    send "userhost " & strNick
            End Select
        End If
    End If
        
        '''Autosize:
        'AutoSizegames and AutoSizePlayers make the GH banner image flicker
        If Visible = True Then
            Call AutoSizeGames
            If Left(tabIRC.SelectedItem.Caption, 1) = "#" Then Call AutoSizePlayers
        End If
        
        If blnLogin = True Then
            
            'User idle time http://www.vbforums.com/showthread.php?t=516757 yosef_mreh
            'Won't work on older than Win2k
            lii.cbSize = Len(lii)
            Call GetLastInputInfo(lii)
            Dim blnTickCount As Boolean
            
             If (GetTickCount - lii.dwTime) / 1000 > 300 Then blnTickCount = True
         
            If blnTickCount = True Then
                If strStatusMsg <> "=AFK" Then
                    changeStatus "=AFK"
                    send "NOTICE " & strChannel & " S" & strStatusMsg
                End If
            Else
                If strStatusMsg = "=AFK" Then
                    changeStatus vbNullString
                    send "NOTICE " & strChannel & " S"
                End If
                
                If FindProcess("hedgewars") = True Then
                    If strStatusMsg <> "=HW" Then
                        changeStatus "=HW"
                        send "NOTICE " & strChannel & " S" & strStatusMsg
                    End If
                Else
                    If strStatusMsg = "=HW" Then
                        changeStatus vbNullString
                        send "NOTICE " & strChannel & " S"
                    End If
                End If
                
            End If
        End If
        
        'is GTA2 running?
        If lngPID <> 0 And FindProcess(TXT_GTA2EXE, , lngPID) = True Then
            If blnInGame = False Then
                Call EnumWindows(AddressOf EnumWindowsProc, 0)
            End If
        Else
            lngWindowHandle = 0
            blnInGame = False
            blnLobby = False
            frmGH.timStamp.Enabled = False
        End If
       
        If blnLobby = False And blnInGame = False Then
            'If strStatusMsg <> "A" And InStr(strStatusMsg, "=") = 0 And strStatusMsg <> vbNullString Then
            
            Select Case strStatusMsg
                Case "A", "=AFK", vbNullString, "=HW"
                    'do nothing
                Case Else
                    send "NOTICE " & strChannel & " S"
                    changeStatus (vbNullString)
            End Select
            
            'If strStatusMsg <> "A" And strStatusMsg <> "=AFK" And strStatusMsg <> vbNullString Then
             
        End If

        'Time GTA2 has been in lobby or in game
        If blnLobby = True Then
            lngLobby = lngLobby + 1
            
'       Activate/focus the GTA2 chatbox control on startup 'obsolete, GTA2 takes care of that now
'            If blnReadyForJoiners Then 'Activate hosting chatbox
'                If lngLobby < 3 Then Call SendMessage(lngHandleChat, WM_ACTIVATE, 1, 0)
'            ElseIf lngLobby > 0 Then 'Activate joining chatbox
'                Call SendMessage(lngHandleJoinChat, WM_ACTIVATE, 1, 0)
'            End If
            
            If strStatusMsg = vbNullString And strNickLastJoined <> vbNullString Then
                For i = 1 To frmGH.lvPlayers(1).ListItems.count
                    If frmGH.lvPlayers(1).ListItems.Item(i).Text = strNick Then
                        strStatusMsg = "=" & strNickLastJoined
                        frmGH.lvPlayers(1).ListItems.Item(i).ListSubItems(2) = Left$(strNickLastJoined, 12)
                        frmGH.lvPlayers(1).ListItems.Item(i).ListSubItems(2).ToolTipText = strNickLastJoined
                        send "NOTICE " & strChannel & " S=" & strNickLastJoined
                        strNickLastJoined = vbNullString
                        Exit For
                    End If
                Next i
            End If
        Else
            If lngLobby <> 0 Then
                displaychat strChannel, strGHColor, "GTA2 lobby was open for " & Duration(lngLobby, 2)
                lngLobby = lngLobby + Val(readINI("Statistics", "GTA2 Lobby Time", DOCUMENTS & "\gta2gh.ini"))
                Call WriteINI("Statistics", "GTA2 Lobby Time", lngLobby, DOCUMENTS & "\gta2gh.ini")
                lngLobby = 0
            End If
            
            strPreviousScriptFile = vbNullString
            strPreviousMapFile = vbNullString
            blnReadyForJoiners = False
            blnCalculatedGTA2checksum = False
            
            If blnHosted = True Then
                'set status to GTA2
                If strStatusMsg <> "2" And strStatusMsg <> "A" And strStatusMsg <> "=HW" And blnInGame = True Then
                    strStatusMsg = "2"
                    Call changeStatus(strStatusMsg)
                    send "NOTICE " & strChannel & " :S2"
                End If
                
                Call RemovePlayersGameFromList
                intSecondsWaited = intSecondsWaited + 1

                'if you hosted and GTA2 has been closed for 5 seconds then remove your game
                If intSecondsWaited > 10 Then
                   intSecondsWaited = 0
                   blnHosted = False
                   Call RemovePlayersGameFromList
                End If
            End If
            
            If blnInGame = True Then
                lngGTA2RunningTime = lngGTA2RunningTime + 1
'                If lngGTA2RunningTime > 31536000 Then
'                    lngGTA2RunningTime = 1
'                    MsgBox "WASTE-OF-POWER-BONUS! You left GTA2 open for an entire year!"
'                End If
            Else
                If lngGTA2RunningTime <> 0 Then
                    displaychat strChannel, strGHColor, "GTA2 was in game for " & Duration(lngGTA2RunningTime, 2)
                    lngGTA2RunningTime = lngGTA2RunningTime + Val(readINI("Statistics", "GTA2 Running Time", DOCUMENTS & "\gta2gh.ini"))
                    Call WriteINI("Statistics", "GTA2 Running Time", lngGTA2RunningTime, DOCUMENTS & "\gta2gh.ini")
                    lngGTA2RunningTime = 0
                    blnPlayReplay = False
                End If
            End If
        End If
    
    If sockIRC.State = sckClosed Then
        cmdToolbar(BTN_SIGN_IN).Enabled = True
        mnuFileSignIn.Enabled = True
    End If
    
    If blnDisconnectClick = False Then
        Select Case sockIRC.State
            Case sckClosed
                strState = "Disconnected"
                Disconnect
                'If blnConnectClick = False Then
                'cmdToolbar(BTN_SIGN_IN).Enabled = True
                'mnuFileSignIn.Enabled = True
                'End If
                blnReconnect = True
            Case sckOpen
                strState = "Open"
            Case sckListening
                strState = "Listening"
            Case sckConnectionPending
                strState = "Connection pending"
            Case sckResolvingHost
                strState = "Resolving host"
            Case sckHostResolved
                strState = "Host resolved"
            Case sckConnecting
                strState = "Connecting"
            Case sckConnected
                strState = "Connected"
                blnDisconnect = False
            Case sckClosing
                strState = "Closing"
                cmdToolbar(BTN_SIGN_OUT).Enabled = False
                mnuFileSignOut.Enabled = False
                sockIRC.Close
            Case sckError
                strState = "The server could be down or the host may not be resolving."
                displaychat strDestTab, strConnectionColor, strState
                Call Disconnect
                cmdToolbar(BTN_SIGN_IN).Enabled = True
                mnuFileSignIn.Enabled = True
                sockIRC.Close
                blnDisconnectClick = False
                blnReconnect = True
        End Select
        'txtStats = "Useless statistics:" & vbNewline & "Bytes received: " & dblBytesReceived & vbNewline & "Bytes sent: " & dblBytesSent
        'If strState <> strOldState And strState <> "Connected" And strState <> "Connecting" Then displaychat strDestTab, strConnectionColor, "* " & strState & " *"
        strOldState = strState
    End If
    
    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    
    If intTimerError = 0 Then
        displaychat strDestTab, strTextColor, "Network status timer error: Line " & strErrLine & " " & strErrdesc
        send "PRIVMSG " & gta2ghbot & " :Network status timer error: Line " & strErrLine & " " & strErrdesc & " GetTickCount:" & GetTickCount() & " " & lii.dwTime
        intTimerError = 1
    End If
End Sub

Public Sub Reconnect()
    If blnDisconnectClick = False Then
        If intServerNum = 1 Then
            intServerNum = 0
        Else
            intServerNum = 1
        End If
        cmdToolbar_Click (BTN_SIGN_IN)
    End If
End Sub

Private Sub Form_Initialize()
On Error Resume Next

Dim i As Long
Dim strString As String

If command = vbNullString Then
    'frmGH.Visible = True
Else
    Dim argv() As String
    Dim strMap As String
    Dim doExit As Boolean
    Dim doJoin As Boolean
    
    strString = Replace(command, "%20", " ")

    i = InStr(strString, "-m")
    If i Then
        lngMaster = Mid$(strString, i + 3, (InStr(i + 3, strString, " ")) - (i + 3))
    End If

    i = InStr(strString, "-d")
    If i Then
        frmGH.Visible = False
        If lngMaster = 0 Then lngMaster = 1
        Call CopyURLToFile(Mid$(strString, i + 3, 666), GetTmpPath & "gta2map.7z")
    Else
        'Call frmGH.main
    End If
End If

'Const SWP_SHOWWINDOW = &H40
'Const SWP_NOMOVE = &H2
'Const SWP_NOSIZE = &H1
'only allow one copy of GH to run at a time (can easily be fooled by renaming the exe)
'If App.PrevInstance = True Then 'Or FindProcess("gta2gh.exe", True) Then
''    Select Case MsgBox("Found another gta2gh.exe process. Should I stop it? ", vbYesNo Or vbQuestion Or vbDefaultButton1, App.Title)
''        Case vbYes
''            FindProcess "gta2gh.exe", True
''        Case vbNo
'            Dim lngWindow As Long
'            lngWindow = FindWindow(vbNullString, "GTA2 Game Hunter v" & TXT_GHVER)
'            SetWindowPos lngWindow, 0, 0, 0, 0, 0, SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE
'            Call OpenIcon(lngWindow)
'            SetForegroundWindow lngWindow
'            End
''    End Select
'End If
'Use the operating system style for XP effects where available:
InitCommonControls 'run the API call

'''BenMillard''' set initial tab index:
tabSelected = 1

End Sub

Public Function AudioFileCheck() As Boolean
On Error GoTo oops
    AudioFileCheck = True
    
    If Exists(strGTA2path & "data\audio\wil.raw") = False Then
        If strGTA2path = vbNullString Then
            displaychat strDestTab, strGHColor, "Where is GTA2 really installed?"
        Else
            displaychat strDestTab, strGHColor, strGTA2path & "data\audio\wil.raw is missing. This isn't a valid GTA2 folder!"
        End If
        
        Dim strTemp As String
        strTemp = BrowseFile(strGTA2path & TXT_GTA2EXE)
        If Len(strTemp) > Len(TXT_GTA2EXE) Then
            strGTA2path = Left$(strTemp, Len(strTemp) - Len(TXT_GTA2EXE))
        End If
        If strTemp = vbNullString Then
            AudioFileCheck = False 'cancel button pushed?
        Else
            If Exists(strGTA2path & "data\audio\wil.raw") = True Then
                With cr
                    .ClassKey = HKEY_CURRENT_USER
                    .SectionKey = "SOFTWARE\GTA2 Game Hunter"
                    .ValueType = REG_SZ
                    .ValueKey = "GTA2Folder"
                    .Value = strGTA2path
                End With
            End If
        End If
    End If
    
    Exit Function
oops:
    strErrdesc = Err.Description
    displaychat strDestTab, strTextColor, "Error during audio file check: " & strErrdesc
End Function

Private Sub rtbChatbox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Show right-click context menu
  If Button = vbRightButton Then
    m_hwndEdit = rtbChatbox(Index).hwnd
    SetEditMenuItemText True
    ' Execution stops here until context menu is hidden.
    ' (TrackPopupMenu() behaves identically!)
    PopupMenu mnuEdit   ' Invokes mnuEdit_Click()
    SetEditMenuItemText False
  End If
End Sub

Private Sub rtbTopic_mouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If isURL(RichWordOver(rtbTopic(Index), x, y)) Then
        SetMouseCursor IDC_HAND
    End If
End Sub

Private Sub rtbHistory_mouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If isURL(RichWordOver(rtbHistory(Index), x, y)) Then
        SetMouseCursor IDC_HAND
    End If

'rebug.Print RichWordOver(rtbHistory(Index), x, y)
'If isURL(RichWordOver(rtbHistory(Index), x, y)) Then
    'Screen.MouseIcon = Me.Icon
    'Screen.MousePointer = 1
    'Screen.MousePointer = 99
'Else
    'Screen.MousePointer = 0
'End If

End Sub

Private Sub rtbHistory_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
On Error GoTo oops
   
Dim strChar As String

'Call giveChatFocus History doesn't need focus but I need to improve the way text is copied

'Show right-click context menu
If Button = vbRightButton Then
    m_hwndEdit = rtbHistory(Index).hwnd
    SetEditMenuItemText True
    ' Execution stops here until context menu is hidden.
    ' (TrackPopupMenu() behaves identically!)
    PopupMenu mnuEdit   ' Invokes mnuEdit_Click()
    SetEditMenuItemText False
    Exit Sub
End If

'This is for Internet URL linking
If Button = vbLeftButton Then  'left click
    If rtbHistory(Index).SelLength > 0 Then Exit Sub
    If rtbHistory(Index).SelStart = Len(rtbHistory(Index).Text) Then Exit Sub
    strChar = Mid$(rtbHistory(Index).Text, rtbHistory(Index).SelStart + 1, 1)
    Select Case strChar
        Case vbCr, vbTab
            Exit Sub
    End Select

    Dim txt As String
    txt = RichWordOver(rtbHistory(Index), x, y)
    Call urlCheck(txt)
End If
 
Exit Sub

oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "History URL error: " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :History URL error: " & strErrdesc & " Line: " & strErrLine
End Sub

Private Sub urlCheck(txt As String)
On Error GoTo oops
Dim ret As Long
Dim strLow As String 'lowercase version of txt
strLow = LCase$(txt)

If isURL(strLow) Then
    
    If Left$(strLow, 1) = "#" Then
        send "JOIN " & strLow
        Exit Sub
    End If
    
    'potential exploit, add "gtamp.com/maps/" to any URL and GH will download from it
    
    If blnchkAutoDownload Then
        If InStr(strLow, "gtamp.com/maps/") Then
            If Right$(strLow, 2) = "7z" Or Right$(strLow, 3) = "zip" Or InStr(22, strLow, ".") = 0 Then
                Call CopyURLToFile(txt, GetTmpPath & "gta2map.7z")
                Exit Sub
            End If
        End If
        
        If InStr(strLow, "gtamp.com/mapscript/maplist/download.php?mmp=") Then
            Call CopyURLToFile(txt, GetTmpPath & "gta2map.7z")
            Exit Sub
        End If
        
        If InStr(strLow, "gtamp.com/mapscript/maplist/autodl.php?mmp=") Then
            Call CopyURLToFile(txt, GetTmpPath & "gta2map.7z")
            Exit Sub
        End If
        
        If InStr(strLow, "gtamp.com/maplist/maplump/") Then
            Call CopyURLToFile(txt, GetTmpPath & "gta2map.7z")
            Exit Sub
        End If
        
'        If Left$(strLow, 22) = "http://gtamp.com/maps/" Then
'            If Right$(strLow, 2) = "7z" Or Right$(strLow, 3) = "zip" Or InStr(22, strLow, ".") = 0 Then
'                Call CopyURLToFile(txt, GetTmpPath & "gta2map.7z")
'                Exit Sub
'            End If
'        End If
'
'        If Left$(strLow, 52) = "http://gtamp.com/mapscript/maplist/download.php?mmp=" Then
'            Call CopyURLToFile(txt, GetTmpPath & "gta2map.7z")
'            Exit Sub
'        End If
'
'        If Left$(strLow, 50) = "http://gtamp.com/mapscript/maplist/autodl.php?mmp=" Then
'            Call CopyURLToFile(txt, GetTmpPath & "gta2map.7z")
'            Exit Sub
'        End If
'
'        If Left$(strLow, 33) = "http://gtamp.com/maplist/maplump/" Then
'            Call CopyURLToFile(txt, GetTmpPath & "gta2map.7z")
'            Exit Sub
'        End If
        
        If InStr(strLow, "https://projectcerbera.com/gta/2/") Then
            If Right$(strLow, 3) = "zip" Or Right$(strLow, 2) = "7z" Then
                Call CopyURLToFile(txt, GetTmpPath & "gta2map.7z")
                Exit Sub
            End If
        End If
        
    End If
    
    'if this is a URL then start associated browser and goto web site
    ret = ShellExecute(0&, vbNullString, txt, vbNullString, vbNullString, vbNormalFocus)
End If
    
Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    strErrNum = Err.Number
    displaychat strDestTab, strGHColor, "Error: " & GetTmpPath & "gta2map.7z" & " " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :urlCheck: " & strErrNum
End Sub

Private Sub rtbTopic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo oops
    'Show right-click context menu
    If Button = vbRightButton Then
        m_hwndEdit = rtbTopic(Index).hwnd
        SetEditMenuItemText True
        ' Execution stops here until context menu is hidden.
        ' (TrackPopupMenu() behaves identically!)
        PopupMenu mnuEdit   ' Invokes mnuEdit_Click()
        SetEditMenuItemText True
        Exit Sub
    End If
  
    If Button = vbLeftButton Then 'check if left clicked on a URL
        If rtbTopic(Index).SelLength > 0 Then Exit Sub
        If rtbTopic(Index).SelStart = Len(rtbTopic(Index).Text) Then Exit Sub
        Dim txt As String
        txt = RichWordOver(frmGH.rtbTopic(Index), x, y)
        Call urlCheck(txt)
    End If
   
 Exit Sub

oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Topic URL error: " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :Topic URL error: " & strErrdesc & " Line: " & strErrLine
End Sub

' To change menu command text, this proc must
' be called BEFORE the Edit menu is displayed.
' Will fail miserably if executed when menu IS displayed!!
' Called right before & right after context menu is displayed
' in Form_KeyDown() (Shift+F10) & Rich1_MouseUp()

Private Sub SetEditMenuItemText(bIsContext As Boolean)

  mnuEditUndo.Caption = "&Undo" & IIf(bIsContext, vbNullString, vbTab & "Ctrl+Z")
  mnuEditCut.Caption = "Cu&t" & IIf(bIsContext, vbNullString, vbTab & "Ctrl+X")
  mnuEditCopy.Caption = "&Copy" & IIf(bIsContext, vbNullString, vbTab & "Ctrl+C")
  mnuEditPaste.Caption = "&Paste" & IIf(bIsContext, vbNullString, vbTab & "Ctrl+V")
  mnuEditDelete.Caption = "&Delete" & IIf(bIsContext, vbNullString, vbTab & "Del")
  mnuEditSelAll.Caption = "Select &All" & IIf(bIsContext, vbNullString, vbTab & "Ctrl+A")

End Sub

' >> The Edit menu is already displayed before this code is executed !!! <<

Private Sub mnuEdit_Click()
  EnableEditMenuItems
End Sub

' The Rich ctrl processes edit accelerator keys but does
' not display a corresponding right-click edit context menu (?)
' So NO shortcut keys are defined for the following cmds:
' Undo, Cut, Copy, Paste, Delete & Select All.
' Allows their respective shortcut key strings to be added
' & remove appropriately by mnuEdit_Click() above.
' Their corresponding click events below are called only from
' their respective menu item cmds (not from accelerators).
' The Enabled property is the menu's default property & is either True or False.
' Happens after menu is displayed!

Private Sub EnableEditMenuItems()

Dim cr As CHARRANGE, dwTxtLen As Long

' Fill CR w/ current selection range
SendMessage m_hwndEdit, EM_EXGETSEL, 0, cr

' Disable Paste, Delete & Undo if the read only rtbHistory box has focus

Dim i As Integer
Dim blnHistory As Boolean 'true if you right click chat history

'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
For i = 1 To rtbHistory.UBound
  If m_hwndEdit = rtbHistory(i).hwnd Then
      blnHistory = True
      Exit For
  End If
Next

If blnHistory Then
  mnuEditPaste = False
  mnuEditDelete = False
  mnuEditUndo = False
  mnuEditCut = False
Else
  ' Paste: enable of clipboard text, disable if not
  mnuEditPaste = Len(Clipboard.GetText)
  ' Undo:
  mnuEditUndo = SendMessage(m_hwndEdit, EM_CANUNDO, 0, 0)
  ' Copy, Delete: enable if a selection, disable if no selection
  mnuEditDelete = (cr.cpMin < cr.cpMax)
  mnuEditCut = (cr.cpMin < cr.cpMax)
End If

' Copy: enable if a selection, disable if no selection
mnuEditCopy = (cr.cpMin < cr.cpMax)

' Select All: disable if everything's already selected, enable otherwise.
' The Rich ctrl ALWAYS has a CrLf at the end of it's contents
' which is not seen by WM_GETTEXTLENGTH.
dwTxtLen = SendMessage(m_hwndEdit, WM_GETTEXTLENGTH, 0, 0)
mnuEditSelAll = Not (cr.cpMin = 0 And cr.cpMax = dwTxtLen + 2&)

'Debug.Prin "TxtLen: "; SendMessage(m_hwndEdit, WM_GETTEXTLENGTH, 0, 0)
'Debug.Prin "CR.cpMin: " & CR.cpMin
'Debug.Prin "CR.cpMax: " & CR.cpMax

End Sub

Private Sub mnuEditUndo_Click()   ' Ctrl+Z
  SendMessage m_hwndEdit, EM_UNDO, 0, 0
End Sub

Private Sub mnuEditCut_Click()   ' Ctrl+X
  SendMessage m_hwndEdit, WM_CUT, 0, 0
End Sub

Private Sub mnuEditCopy_Click()   ' Ctrl+C
  SendMessage m_hwndEdit, WM_COPY, 0, 0
End Sub

Private Sub mnuEditPaste_Click()   ' Ctrl+V
  'rtbChatbox(tabIRC.SelectedItem.Index - 1).SelText = rtbChatbox(tabIRC.SelectedItem.Index - 1).SelText & Clipboard.GetText()
  SendMessage m_hwndEdit, WM_PASTE, 0, 0
End Sub

Private Sub mnuEditDelete_Click()   ' Del
  SendMessage m_hwndEdit, WM_CLEAR, 0, 0
End Sub

Private Sub mnuEditSelAll_Click()   ' Ctrl+A
  Dim cr As CHARRANGE
  cr.cpMin = 0: cr.cpMax = -1
  SendMessage m_hwndEdit, EM_EXSETSEL, 0, cr
End Sub

Private Sub mnuHelpCommands_Click()
Call displayCommands
End Sub

Public Sub RemovePlayersGameFromList()
On Error GoTo oops:
    Dim i As Integer
    
    For i = 1 To lvGames(0).ListItems.count
        If strNick = lvGames(0).ListItems.Item(i) Then
            lvGames(0).ListItems.Remove (i)
            If blnConnected = False Then Exit Sub
            send "NOTICE " & strChannel & " :C"
            If strStatusMsg <> "A" And strStatusMsg <> "=HW" Then changeStatus (vbNullString)
            Exit For
        End If
    Next
    
    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error removing your game from list: " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :Error removing game from list: " & strErrdesc & " Line: " & strErrLine
End Sub

Public Sub AdvertiseHostedGame()
Dim strScriptFile As String
Dim strMapFile As String
Dim strStyleFile As String
Dim i As Integer
Dim intDupeNo As Integer 'The list index number of the duplicate game
Dim strPlayReplay As String

'when hosting add your own game to the list
On Error GoTo oops

    If blnConnected = False Or blnLogin = False Then Exit Sub
    
    Dim blnDuplicateFound As Boolean
    
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueKey = "MapDesc"
        .ValueType = REG_SZ
        .Value = strGTA2MapDesc
    End With
    
    'check if the game is already in the list
    For intDupeNo = 1 To lvGames(0).ListItems.count
        If strNick = lvGames(0).ListItems.Item(intDupeNo) Then
            blnDuplicateFound = True
            Exit For
        End If
    Next
    
    If strGTA2MapDesc = vbNullString Then 'if the map desc is empty then no point hosting
        Exit Sub
    End If
    
    'search lvMMPlist for map description to get the corresponding MMP file
    For i = 1 To lvMMPlist.ListItems.count
        If LCase(lvMMPlist.ListItems.Item(i).ListSubItems(4).Text) = LCase(strGTA2MapDesc) Then
            strGTA2MMP = lvMMPlist.ListItems.Item(i).Text
            strMapFile = lvMMPlist.ListItems.Item(i).ListSubItems(1).Text & ".gmp"
            strStyleFile = lvMMPlist.ListItems.Item(i).ListSubItems(2).Text & ".sty"
            strScriptFile = lvMMPlist.ListItems.Item(i).ListSubItems(3).Text & ".scr"
            strGTA2MapDesc = lvMMPlist.ListItems.Item(i).ListSubItems(4).Text
            intPlayerCount = Val(lvMMPlist.ListItems.Item(i).ListSubItems(5).Text)
            If intPlayerCount = 0 Then intPlayerCount = 6
            Exit For
        End If
    Next i
    
    If strGTA2MMP = vbNullString Then
        displaychat strChannel, strGHColor, "Failed to find an MMP file containing " & strGTA2MapDesc
        strGTA2MapDesc = vbNullString
        Exit Sub
    End If
    
    If blnDuplicateFound = True Then 'You already advertised your game
        'Check if you changed any settings from last time
        'I tried doing a With lvGames(0).ListItems.Item(i) here but the if statement didn't like it

        If lvGames(0).ListItems.Item(intDupeNo).ListSubItems.Item(1) = strPasswordProtectGame And _
        lvGames(0).ListItems.Item(intDupeNo).ListSubItems.Item(2) = Right$(strCountries(intCountryIndex), 2) And _
        lvGames(0).ListItems.Item(intDupeNo).SmallIcon = intCountryIndex + 1 And _
        lvGames(0).ListItems.Item(intDupeNo).ListSubItems.Item(3) = strGTA2MapDesc And _
        lvGames(0).ListItems.Item(intDupeNo).ListSubItems.Item(3).ToolTipText = strGTA2MMP And _
        lvGames(0).ListItems.Item(intDupeNo).ListSubItems.Item(4) = TXT_GHVER Then Exit Sub
        'You didn't change anything, abort!

        'You changed a setting on your hosted game...
        With lvGames(0).ListItems.Item(intDupeNo)
            .SmallIcon = intCountryIndex + 1
            .ToolTipText = Left$(strCountries(intCountryIndex), Len(strCountries(intCountryIndex)) - 5)
            With .ListSubItems
                .Item(1) = strPasswordProtectGame 'Yes or No
                .Item(2) = Right$(strCountries(intCountryIndex), 2) 'CC change
                .Item(2).ToolTipText = lvGames(0).ListItems.Item(intDupeNo).ToolTipText
                .Item(3) = strGTA2MapDesc
                .Item(3).ToolTipText = strGTA2MMP
                .Item(4) = TXT_GHVER
                If blnPlayReplay = True Then
                    strPlayReplay = "Play Replay"
                End If
                .Item(4).ToolTipText = strPlayReplay
                '.Item(5) = strGTA2version
            End With
        End With
    Else
        With lvGames(0).ListItems.Add(, , strNick, , intCountryIndex + 1)
            .ToolTipText = Left$(strCountries(intCountryIndex), Len(strCountries(intCountryIndex)) - 5)
            .ListSubItems.Add , , strPasswordProtectGame 'Yes or No
            .ListSubItems.Add , , Right$(strCountries(intCountryIndex), 2), , _
              Left$(strCountries(intCountryIndex), Len(strCountries(intCountryIndex)) - 5)
            .ListSubItems.Add , , strGTA2MapDesc, , strGTA2MMP
            .ListSubItems.Add , , TXT_GHVER
            '.ListSubItems.Add , , strGTA2version
        End With
    End If
    
    If strPreviousMapFile <> strMapFile Then
        strPreviousMapFile = strMapFile
        strMapChecksum = calc_crc32(strGTA2path & "data\" & strMapFile)
    End If
    
    If strPreviousScriptFile <> strScriptFile Then
        strPreviousScriptFile = strScriptFile
        strScriptChecksum = calc_crc32(strGTA2path & "data\" & strScriptFile)
    End If
    
    If blnCalculatedGTA2checksum = False Then
        strExecutableChecksum = calc_crc32(strGTA2path & TXT_GTA2EXE)
        blnCalculatedGTA2checksum = True
    End If
    
    'send game to channel
    Dim strAdvertisement As String
    Dim strGameOptions As String
    Dim strLocked As String
    If strPasswordProtectGame = "Yes" Then strLocked = "P"
    If blnPlayReplay = True Then strPlayReplay = "R"
    
    If strPlayReplay & strLocked <> vbNullString Then
        strGameOptions = "/" & strPlayReplay & strLocked
    End If
                    
    strAdvertisement = "G" & strGTA2MMP & strGameOptions
    send "NOTICE " & strChannel & " :" & strAdvertisement
    
    Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error during advertise hosted game: " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :Error during advertise hosted game: " & strErrdesc & " Line: " & strErrLine
End Sub

Public Sub AddDescriptionAndFileToListView()
    On Error GoTo oops
    Dim strGTA2MapFileName As String
    Dim strMMPfullpath As String
    Dim NoPlayerCountMapsArray As New cSortArray
    Dim i As Long
    Dim j As Long
    Dim blnRefresh As Boolean
    Dim strRemoved(3) As String
    
    'add the filenames to a ListView for searching through later
    frmGH.lvMMPlist.ListItems.Clear
    If frmGH.lvMMPlist.ColumnHeaders.count = 0 Then
        frmGH.lvMMPlist.ColumnHeaders.Add(, , "MMP filename", 3500).Tag = "STRING"
        frmGH.lvMMPlist.ColumnHeaders.Add(, , "GMP", 3500).Tag = "STRING"
        frmGH.lvMMPlist.ColumnHeaders.Add(, , "STY", 3500).Tag = "STRING"
        frmGH.lvMMPlist.ColumnHeaders.Add(, , "SCR", 3500).Tag = "STRING"
        frmGH.lvMMPlist.ColumnHeaders.Add(, , "Description", 3500).Tag = "STRING"
        frmGH.lvMMPlist.ColumnHeaders.Add(, , "PlayerCount", 3500).Tag = "STRING"
    End If
    frmGH.lvMMPlist.View = lvwReport

    'Add the MMP filename and description to frmGH.lvMMPlist
    strGTA2MapFileName = Dir(strGTA2path & "data\*.mmp", vbHidden + vbNormal + vbSystem + vbReadOnly + vbArchive)
    Do While strGTA2MapFileName <> vbNullString
        strMMPfullpath = strGTA2path & "data\" & strGTA2MapFileName
        With frmGH.lvMMPlist.ListItems.Add(, , Left$(strGTA2MapFileName, Len(strGTA2MapFileName) - 4))
            Dim strGMP As String
            Dim strSCR As String
            Dim strSTY As String
            Dim strDescription As String
            Dim strPlayerCount As String
            strGMP = readINI("MapFiles", "GMPFile", strMMPfullpath)
            strSTY = readINI("MapFiles", "STYFile", strMMPfullpath)
            strSCR = readINI("MapFiles", "SCRFile", strMMPfullpath)
            If Len(strGMP) > 4 Then strGMP = Left$(strGMP, Len(strGMP) - 4)
            If Len(strSTY) > 4 Then strSTY = Left$(strSTY, Len(strSTY) - 4)
            If Len(strSCR) > 4 Then strSCR = Left$(strSCR, Len(strSCR) - 4)
            strDescription = readINI("MapFiles", "Description", strMMPfullpath)
            strPlayerCount = readINI("MapFiles", "PlayerCount", strMMPfullpath)
            .ListSubItems.Add , , strGMP
            .ListSubItems.Add , , strSTY
            .ListSubItems.Add , , strSCR
            .ListSubItems.Add , , strDescription
            .ListSubItems.Add , , strPlayerCount
            If strPlayerCount = vbNullString Then NoPlayerCountMapsArray.AddItem (strGMP & strSTY & strSCR)
        End With
        
        strGTA2MapFileName = Dir
    Loop
    
    'Remove from list and erase any MMP files that point to GMP, STY or SCR files that don't exist
'    For i = 1 To lvMMPlist.ListItems.count
'
'        With lvMMPlist.ListItems(i)
'            Erase strRemoved
'
'            If Exists(strGTA2path & "data\" & .ListSubItems(1)) = False Then
'                strRemoved(0) = .ListSubItems(1) & " "
'            End If
'
'            If Exists(strGTA2path & "data\" & .ListSubItems(2)) = False Then
'                strRemoved(1) = .ListSubItems(2) & " "
'            End If
'
'            If Exists(strGTA2path & "data\" & .ListSubItems(3)) = False Then
'                strRemoved(2) = .ListSubItems(3)
'            End If
'
'            strRemoved(3) = strRemoved(0) & strRemoved(1) & strRemoved(2)
'
'            If strRemoved(3) <> vbNullString Then
'                If modFileKill(strGTA2path & "data\" & lvMMPlist.ListItems(i).Text & ".mmp") = True Then
'                    displaychat strChannel, strGHColor, "Removed " & lvMMPlist.ListItems(i).Text & ".mmp because " _
'                    & strRemoved(3) & " doesn't exist"
'                    blnRefresh = True
'                End If
'            End If
'
'            'Rename obsolete MMP files
'            For j = 1 To NoPlayerCountMapsArray.count - 1
'            With lvMMPlist.ListItems(i)
'                If .SubItems(5) <> vbNullString Then
'                    If LCase$(.SubItems(1) & .SubItems(2) & .SubItems(3)) = _
'                    LCase$(NoPlayerCountMapsArray.Item(j)) Then
'                        If modFileKill(strGTA2path & "data\" & .Text & ".mmp") Then
'                            displaychat strChannel, strGHColor, "Obsolete file removed: " & .Text & ".mmp"
'                            blnRefresh = True
'                        End If
'                    End If
'                End If
'            End With
'        Next j
'        End With
'    Next
 
'    If blnRefresh = True Then Call PreHost
    
Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, strTextColor, "error during AddDescriptionAndFileToListView: " & strErrLine & " " & strErrdesc
    send "PRIVMSG " & gta2ghbot & " :AddDescription: " & strErrLine & " " & strErrdesc
End Sub

Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
'Key Code Constants: http://msdn.microsoft.com/en-us/library/aa243025(VS.60).aspx
Dim i As Integer

'Which buttons were pressed?
Select Case True

Case Shift = vbCtrlMask And KeyCode = vbKeyG 'Ctrl+G
    cmdHost_Click
'Close all tabs:
Case Shift = vbCtrlMask And KeyCode = vbKeyF4 'Ctrl+F4
    KeyCode = 0
    
    'Is there only 1 tab?
    If tabIRC.Tabs.count < 2 Then Exit Sub 'don't close it, end here
    
    '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
    For i = 2 To rtbHistory.UBound
        '''BenMillard''' changed from Unload to invisible:
        rtbHistory(i).Visible = False
        rtbChatbox(i).Visible = False
        rtbTopic(i).Visible = False
    Next
    
    'Player lists:
    '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
    If lvPlayers.UBound > 1 Then
        For i = 2 To lvPlayers.UBound
            lvPlayers(i).Visible = False '''BenMillard''' stopped using Unload for lvPlayers here
            send "PART " & lvPlayers(i).Tag
        Next
    End If
    
    'Leave any channels other than #gta2gh:
    For i = 2 To tabIRC.Tabs.count
        If Left$(tabIRC.Tabs(i).Caption, 1) = "#" Then send "PART " & tabIRC.Tabs(i).Caption
    Next
    
    'Remove all tabs, then restore #gta2gh tab:
    tabIRC.Tabs.Clear
    tabIRC.Tabs.Add , , strChannel '#gta2 set in Form_Load
    
    'Properly select first tab:
    Call SelectTab(1)
    
''BenMillard''' seperated these because they didn't run as the same Case anyway

'Close selected tab:
Case Shift = vbCtrlMask And KeyCode = vbKeyW 'Ctrl+W
    KeyCode = 0 'reset
    
    'Is there only 1 tab?
    If tabIRC.Tabs.count < 2 Then Exit Sub 'don't close it, end here
    
    'Cannot close the first tab:
    If tabIRC.SelectedItem.Index = 1 Then Exit Sub 'stop here
    
    'Close selected tab:
    Dim strTab As String
    strTab = tabIRC.SelectedItem.Caption
    
    '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
    For i = 1 To rtbHistory.UBound
        If LCase(rtbHistory(i).Tag) = LCase(strTab) Then
            rtbHistory(i).Visible = False
            rtbChatbox(i).Visible = False
            rtbTopic(i).Visible = False
        End If
    Next
    
    'Is this tab for a channel?
    If Left$(strTab, 1) = "#" Then
        'Leave the channel:
        send "PART " & strTab
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        For i = 2 To lvPlayers.UBound '''wrong to start at 2?
            If LCase(lvPlayers(i).Tag) = LCase(strTab) Then
                lvPlayers(i).Visible = False
            End If
        Next
    End If
    
    'Remove the selected tab, but keep the chat history controls:
    i = tabIRC.SelectedItem.Index
    tabIRC.Tabs.Remove tabIRC.SelectedItem.Index
    
    'Switch to previous tab before removing current tab:
    Call SelectTab(i - 1, False) 'previous, not looping round to the end
End Select

'''BenMillard''' separated these while refactoring the Case lines above.

'Switch to a different tab:
Select Case KeyCode

'Switch to tab 'n' when Alt+'n' or Ctrl+'n' are pushed:
Case vbKey1 To vbKey9
    If (Shift = vbAltMask) Or (Shift = vbCtrlMask) Then
        '''BenMillard''' rewrote this to use ShowTab(i) instead of changing .Selected and expecting tabIRC_Click to do it.
        Call SelectTab(KeyCode - 48) 'pick that tab, clamping to either end instead of looping round
        KeyCode = 0 'reset
    End If
    
'Next/Previous tab:
Case vbKeyTab
    KeyCode = 0
    If Shift = vbCtrlMask + vbShiftMask Then
        'Ctrl+Shift+Tab to move to the tab on the left
        Call switchTabLeft
    End If

    If Shift = vbCtrlMask Then
        'ctrl+tab to move to the tab on the right
        Call switchTabRight
    End If

'Previous tab:
Case vbKeyLeft
    If Shift = vbAltMask Then Call switchTabLeft
    
'Next tab:
Case vbKeyRight
    If Shift = vbAltMask Then Call switchTabRight
End Select


If Shift = 0 Then
    Select Case KeyCode
    
    Case vbKeyF1
        Call displayCommands
    Case vbKeyF2
        Call displayPortHelp
    Case vbKeyF4
        cmdOptions_Click
    Case vbKeyF6
        frmAbout.about
    Case vbKeyF8
        cmdGTA2Manager_Click
    End Select
End If

End Sub

'''BenMillard''' rewrote this to use ShowTab(i) instead of changing .Selected and expecting tabIRC_Click to do it.
Private Sub switchTabLeft()
Call SelectTab(tabIRC.SelectedItem.Index - 1, True) 'previous, looping round to the end
'If tabIRC.SelectedItem.Index = 1 Then
'    tabIRC.Tabs.Item(tabIRC.Tabs.count).Selected = True
'Else
'    tabIRC.Tabs.Item(tabIRC.SelectedItem.Index - 1).Selected = True
'End If
End Sub

'''BenMillard''' rewrote this to use ShowTab(i) instead of changing .Selected and expecting tabIRC_Click to do it.
Private Sub switchTabRight()
Call SelectTab(tabIRC.SelectedItem.Index + 1, True) 'next, looping back to the start
'If tabIRC.SelectedItem.Index = tabIRC.Tabs.count Then
'    tabIRC.Tabs(1).Selected = True
'Else
'    tabIRC.Tabs.Item(tabIRC.SelectedItem.Index + 1).Selected = True
'End If
End Sub

Private Function calc_crc32(ByVal strFile As String) As String
On Error GoTo oops
    Dim cStream As New cBinaryFileStream
    Dim cCRC32 As New cCRC32
    cStream.File = strFile
    lCRC32 = cCRC32.GetFileCrc32(cStream)
    calc_crc32 = Pad_String(Hex(lCRC32), 8, "0")

Exit Function
oops:
    Call ErrorHandler("calc_crc32", Err.Description, Erl)
End Function

Public Function Pad_String(work As String, ReqLength As Long, padChar As String) As String
    On Error GoTo oops:
    Pad_String = String$(ReqLength - Len(work), padChar) & work
    Exit Function
oops:
    Call ErrorHandler("padstring", Err.Description, Erl)
End Function

'Sign In
'''BenMillard''' has combined the buttons into an array.
Public Sub cmdToolbar_Click(Index As Integer)
On Error GoTo oops
Dim i As Integer


'''FOCUS'''
Call giveChatFocus 'return to chat box, then allow button to do its thing


'Detect which button was pressed:
Select Case Index
Case BTN_SIGN_IN
    'displaychat strChannel, strGHColor, "Latest GH version: " & CopyURLToRAM("http://gtamp.com/version.txt")

     'clear games list when signing in
    lvGames(0).ListItems.Clear
    'clear player lists when signing in
    For i = 1 To lvPlayers.UBound
        lvPlayers(i).ListItems.Clear
    Next i
    
'    If blnConnected = True Then
'        send "JOIN " & strChannel & " " & strKey
'
'        If tabIRC.Tabs.Count > 1 Then
'            For i = 1 To tabIRC.Tabs.Count
'                send "JOIN " & tabIRC.Tabs(i).Caption
'            Next i
'        End If
        cmdToolbar(BTN_SIGN_IN).Enabled = False
        mnuFileSignIn.Enabled = False
        cmdToolbar(BTN_SIGN_OUT).Enabled = True
        mnuFileSignOut.Enabled = True
'        Exit Sub
'    End If
    
    With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueKey = "ExternalHostName"
        If strCountryCode = vbNullString Or .Value = vbNullString Then
            blnGotCC = True
            Call getCC
        Else
            blnGotCC = False
        End If
    End With
    
    blnLogin = False
    If strStatusMsg = "=HW" Then strStatusMsg = vbNullString
    strFailedCountryIP = vbNullString
    If strPreferedNick = vbNullString Or Left$(strNick, 3) = "Ped" Or Left$(strNick, 5) = "Guest" Then
        displaychat strDestTab, vbRed, "Set an IRC name and GTA2 name. Press F4 to display options."
        Exit Sub
    End If
    
    If strPassword = vbNullString Then
        displaychat strDestTab, vbRed, "You must set a password. Press F4 to display options."
        Exit Sub
    End If
    
    sockIRC.Close
    strKey = "digdug"
    
    strDestTab = strChannel
    intNickservWaitTime = 0
    blnConnectClick = True
    Call MoveMMPfiles 'moves MMP files from tempMMP to data and then removes tempMMP folder
    
    Call DetectGTA2version
    
    timTimeout.Enabled = True
    If strPort = "" Then strPort = 6667
    displaychat strDestTab, strConnectionColor, "Connecting to: " & strServer(0) & " port " & strPort
    'sockIRC.Connect Left$(strServer(intServerNum), InStr(strServer(intServerNum) & " ", " ") - 1), strPort 'connect to the server
    sockIRC.Connect Left$(strServer(intServerNum), InStr(strServer(0) & " ", " ") - 1), strPort 'connect to the server
Case BTN_CREATE
    Call cmdHost_Click
Case BTN_OPTIONS
    Call cmdOptions_Click
Case BTN_MANAGER
    Call cmdGTA2Manager_Click
Case BTN_SIGN_OUT
    Call cmdDisconnectClick
Case BTN_CANCEL
    Call cmdCancel_Click
End Select

Exit Sub
    
oops:
  strErrLine = Erl
  If Err.Description = "Path not found" Then
      displaychat strDestTab, vbRed, "GTA2 folder was not found!"
      cmdToolbar(BTN_SIGN_IN).Enabled = True
      mnuFileSignIn.Enabled = True
  Else
      displaychat strDestTab, vbRed, "Error connecting: " & Err.Description & " Line: " & strErrLine
      cmdToolbar(BTN_SIGN_IN).Enabled = True
      mnuFileSignIn.Enabled = True
  End If
End Sub

Public Sub cmdOptions_Click()
    If frmOptions.Visible = False Then Call frmOptions.loadSettings
    frmOptions.Show
End Sub

Private Sub cmdGTA2Manager_Click()
    On Error GoTo oops
    Call setGTA2path
    
    If Exists(strGTA2path & "gta2manager.exe") = True Then
        Call FindProcess("gta2manager.exe", True) 'Find and kill process
        displaychat strDestTab, 32896, "Launching GTA2 Manager"
        'Call shellandwait(vbQuote & strGTA2path & "gta2manager.exe" & vbQuote, strGTA2path)
        'lngPID = shellandwait(strGTA2path & "gta2manager.exe", strGTA2path)
        ShellExecute hwnd, "runas", "gta2manager.exe", "", strGTA2path, vbNormalFocus
    Else
        displaychat strDestTab, vbRed, "Can't find " & strGTA2path & "gta2manager.exe"
    End If
Exit Sub

oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, strTextColor, "error launching GTA2 Manager:  " & strErrdesc
End Sub

Private Sub mnuFileCreateGame_Click()
    cmdHost_Click
End Sub

Private Sub mnuFileSignIn_Click()
    Call cmdToolbar_Click(BTN_SIGN_IN)
End Sub

Private Sub mnuFileSignOut_Click()
    cmdDisconnectClick
End Sub

Private Sub launch(strFile As String)
On Error GoTo oops
    If Exists(App.Path & "\" & strFile) = True Then
        Call ShellExecute(Me.hwnd, "Open", App.Path & "\" & strFile, vbNullString, App.Path, vbMaximizedFocus)
    Else
        displaychat strDestTab, strGHColor, App.Path & "\" & strFile & " is missing"
    End If
    
    Exit Sub
oops:
    strErrdesc = Err.Description
    displaychat strDestTab, strTextColor, "Failed to open  " & strFile & " - " & strErrdesc
End Sub
Private Sub mnuToolsOptions_Click()
    cmdOptions_Click
End Sub

Private Sub mnuViewBMTheme_Click()
    intTheme = 1
    mnuViewBMTheme.Checked = True
    mnuViewDarkTheme.Checked = False
    mnuViewLightTheme.Checked = False
    mnuViewClassicTheme.Checked = False
    strTextColor = &H80000008  'Window Text
    strJoinColor = 32768   'green
    strQuitColor = 8388608 'blue
    strForeColor = &H80000008  'Window Text
    strBackColor = &H80000005 'Window Background
    strFontName = "Arial"
    strFontSize = 10
    blnFontBold = False
    blnFontItalic = False
    blnUnderline = True
    lvGames(0).Font = "Microsoft Sans Serif"
    lvGames(0).Font.Size = 8
    lvGames(0).Font.Bold = False
    lvGames(0).Font.Italic = False
    lvGames(0).ForeColor = strForeColor
    lvGames(0).BackColor = strBackColor
    frmGH.BackColor = &H8000000F
    Call applyFormat
End Sub

Private Sub mnuViewDarkTheme_Click()
    intTheme = 2
    mnuViewDarkTheme.Checked = True
    mnuViewBMTheme.Checked = False
    mnuViewClassicTheme.Checked = False
    mnuViewLightTheme.Checked = False
    strTextColor = 12632256 'silver
    strJoinColor = 32768    'green
    strQuitColor = 32768    'green
    strForeColor = strTextColor
    'strBackColor = &H232323 'black
    strBackColor = vbBlack '&H232323 'black
    strFontName = "DejaVu Sans Mono"
    strFontSize = 9
    blnFontBold = False
    blnFontItalic = False
    blnUnderline = False
    lvGames(0).Font = "Microsoft Sans Serif"
    lvGames(0).Font.Size = 8
    lvGames(0).Font.Bold = False
    lvGames(0).Font.Italic = False
    lvGames(0).ForeColor = strForeColor
    lvGames(0).BackColor = strBackColor
    lvGames(0).Appearance = ccFlat
    lvPlayers(1).Appearance = ccFlat
    frmGH.BackColor = &H232323
    Call applyFormat
End Sub

Private Sub mnuViewLightTheme_Click()
    intTheme = 3
    mnuViewLightTheme.Checked = True
    mnuViewDarkTheme.Checked = False
    mnuViewBMTheme.Checked = False
    mnuViewClassicTheme.Checked = False
    strTextColor = &H80000008  'Window Text
    strJoinColor = 32768   'green
    strQuitColor = 8388608 'blue
    strForeColor = &H80000008  'Window Text
    'strBackColor = &H80000005 'Window Background
    strBackColor = 15658734 'mIRC's Placid Hues theme grey
    'strFontName = "FixedSys"
    strFontName = "DejaVu Sans Mono"
    strFontSize = 9
    blnFontBold = False
    blnFontItalic = False
    blnUnderline = False
    lvGames(0).Font = "Microsoft Sans Serif"
    lvGames(0).Font.Size = 8
    lvGames(0).Font.Bold = False
    lvGames(0).Font.Italic = False
    lvGames(0).ForeColor = strForeColor
    lvGames(0).BackColor = strBackColor
    frmGH.BackColor = &H8000000F
    Call applyFormat
   
End Sub

Private Sub mnuViewClassicTheme_Click()
    intTheme = 4
    mnuViewClassicTheme.Checked = True
    mnuViewDarkTheme.Checked = False
    mnuViewLightTheme.Checked = False
    mnuViewBMTheme.Checked = False
    strTextColor = &H80000008  'Window Text
    strJoinColor = 32768   'green
    strQuitColor = 8388608 'blue
    strForeColor = &H80000008  'Window Text
    strBackColor = &H80000005 'Window Background
    strFontName = "Fixedsys"
    strFontSize = 9
    blnFontBold = False
    blnFontItalic = False
    blnUnderline = False
    lvGames(0).Font = "Microsoft Sans Serif"
    lvGames(0).Font.Size = 8
    lvGames(0).Font.Bold = False
    lvGames(0).Font.Italic = False
    lvGames(0).ForeColor = strForeColor
    lvGames(0).BackColor = strBackColor
    frmGH.BackColor = &H8000000F
    Call applyFormat
End Sub
    
Private Sub applyFormat()
On Error GoTo oops

    Dim i As Integer
    Dim lngSelPos As Long
    
    'ListViews match each other:
    '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
    For i = 1 To lvPlayers.UBound
        lvPlayers(i).Font = lvGames(0).Font
        lvPlayers(i).Font.Size = lvGames(0).Font.Size
        lvPlayers(i).Font.Bold = lvGames(0).Font.Bold
        lvPlayers(i).Font.Italic = lvGames(0).Font.Italic
        lvPlayers(i).ForeColor = lvGames(0).ForeColor
        lvPlayers(i).BackColor = lvGames(0).BackColor
    Next
    
    '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
    For i = 1 To rtbChatbox.UBound

        With rtbHistory(i)
           lngSelPos = .SelStart
          .BackColor = strBackColor
          .Font = strFontName
          .Font.Size = strFontSize
          .Font.Bold = blnFontBold
          .Font.Italic = blnFontItalic
          .SelStart = 0
          .SelLength = Len(.Text)
          .SelColor = strForeColor
          .SelBold = blnFontBold
          .SelItalic = blnFontItalic
          .SelFontName = strFontName
          .SelFontSize = strFontSize
          .SelLength = 0
          .SelStart = lngSelPos
        End With
    
        With rtbChatbox(i)
            lngSelPos = .SelStart
            .BackColor = strBackColor
            .Font = strFontName
            .Font.Size = strFontSize
            .Font.Bold = blnFontBold
            .Font.Italic = blnFontItalic
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelColor = strForeColor
            .SelBold = blnFontBold
            .SelItalic = blnFontItalic
            .SelFontName = strFontName
            .SelFontSize = strFontSize
            .SelLength = 0
            .SelStart = lngSelPos
        End With
    
        With rtbTopic(i)
            lngSelPos = .SelStart
            .BackColor = strBackColor
            .Font = strFontName
            .Font.Size = strFontSize
            .Font.Bold = blnFontBold
            .Font.Italic = blnFontItalic
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelColor = strForeColor
            .SelBold = blnFontBold
            .SelItalic = blnFontItalic
            .SelFontName = strFontName
            .SelFontSize = strFontSize
            .SelLength = 0
            .SelStart = lngSelPos
            .TextRTF = .TextRTF
        End With
    
    Next i
    
    Call Form_Resize
    
    '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
    rtbTopic(1).SelRTF = rtbTopic(1).SelRTF
    Exit Sub
oops:
 strErrLine = Erl
 strErrdesc = Err.Description
 displaychat strDestTab, strTextColor, "Error applying format: " & strErrdesc & ", Line: " & strErrLine
 send "PRIVMSG " & gta2ghbot & " :ApplyFormat " & strErrdesc & " " & strErrLine
End Sub

Private Sub mnuToolsGTA2manager_Click()
    cmdGTA2Manager_Click
End Sub

'if the systray icon is clicked then show the form
Private Sub picTray_mouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        blnSystray = False
        WindowState = intPrevWinState
        Show
        If blnchkTray = False Then
            RemoveSystray
        Else
            frmGH.picTray.Picture = Me.Icon
            Call drawTrayIcon
        End If
    End If
End Sub

'remove systray
Private Sub RemoveSystray()
    Shell_NotifyIcon NIM_DELETE, TaskIcon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo oops
    'hide form if close to tray is ticked
    If blnchkCloseTray = True Then
        Cancel = 1
        Hide
        Call Systray
    Else
        cmdExit_Click
    End If

    Exit Sub
oops:
  strErrdesc = Err.Description
  displaychat strDestTab, strTextColor, "Error number: " & Err.Number & " description: " & strErrdesc
  Select Case MsgBox("Are you sure you want to quit?", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)
      Case vbYes
          cmdExit_Click
  End Select
End Sub

Private Sub winRes(lngWidth As Long, lngHeight As Long)
'''Set window to x,y res

WindowState = vbNormal

'Convert from px to twip:
lngWidth = lngWidth * Screen.TwipsPerPixelX
lngHeight = lngHeight * Screen.TwipsPerPixelY

'Apply new size:
Move Left, Top, lngWidth, lngHeight

End Sub


Private Sub loadViewSettings()
On Error GoTo oops
   Dim i As Integer
   With cr
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "SOFTWARE\GTA2 Game Hunter"
        .ValueKey = "chkTimestamp"
        blnchkTime = .Value
        
        frmGH.mnuViewTimestamp.Checked = blnchkTime
        
        .ValueKey = "Gridlines"
        frmGH.mnuViewGridlines.Checked = .Value
        
        frmGH.lvGames(0).GridLines = frmGH.mnuViewGridlines.Checked
        '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
        For i = 1 To frmGH.lvPlayers.UBound
            frmGH.lvPlayers(i).GridLines = frmGH.mnuViewGridlines.Checked
        Next
    End With
    
    Exit Sub
oops:
  strErrdesc = Err.Description
  displaychat strChannel, vbRed, "Error loading theme settings: " & strErrdesc
End Sub

Public Sub changeStatus(ByVal strStatus As String, Optional strName As String)

'Find a player in the player list and set their status
'If the player is you then also change strStatusMsg
 
If strName = vbNullString Then
    strName = strNick
    
    Select Case strStatus
        Case "A"
            strStatusMsg = "A"
            strStatus = "Away"
        Case "2"
            strStatusMsg = "2"
            strStatus = "GTA2"
        Case Else
            strStatusMsg = strStatus
    End Select
End If

strStatus = Replace(strStatus, "=", vbNullString)

Dim i As Integer
Dim itmX As ListItem

'''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
For i = 1 To lvPlayers.UBound
    Set itmX = frmGH.lvPlayers(i).FindItem(strName, lvwText)
    If Not itmX Is Nothing Then
        frmGH.lvPlayers(i).ListItems.Item(itmX.Index).ListSubItems(2) = strStatus
    End If
Next i

ShowListViewColumnHeaderSortIcon frmGH.lvPlayers(1)
Call SortColumn(frmGH.lvPlayers(1), frmGH.lvPlayers(1).SortKey + 1)

End Sub

Private Sub back()
    strStatusMsg = vbNullString
    'send "AWAY"
    send "NOTICE " & strChannel & " :S"
End Sub

Private Sub toggleAwayStatus(Optional ByVal strAwayMessage As String)
    If strStatusMsg = "A" Then
        Call back
    Else
        strStatusMsg = "A"
        strAwayMsg = strAwayMessage
        'send "AWAY :" & strAwayMsg
        send "NOTICE " & strChannel & " SA"
    End If
    
    Call changeStatus(strStatusMsg)
End Sub


'''BenMillard''' has removed the now obsolete hideWindows Sub.

Private Sub saveChannels()
'On Error Resume Next

Dim strChannels As String
Dim i As Integer

For i = 1 To frmGH.tabIRC.Tabs.count
    If Left$(frmGH.tabIRC.Tabs(i).Caption, 1) = "#" Then
        strChannels = strChannels & frmGH.tabIRC.Tabs(i)
    End If
Next

With cr
    .ClassKey = HKEY_CURRENT_USER
    .SectionKey = "SOFTWARE\GTA2 Game Hunter"
    .ValueType = REG_SZ
    .ValueKey = "Channels"
    .Value = strChannels
End With
                
End Sub

Private Sub findReplace()
    Dim strText As String
    'strText = InputBox(vbQuote, vbQuote)
    'rtbHistory(0).Find(strText)
      
End Sub

Private Function SetMouseCursor(CursorType As Long)
  Dim hCursor As Long
  hCursor = LoadCursorLong(0&, CursorType)
  hCursor = SetCursor(hCursor)
End Function

'''FOCUS'''
Public Sub giveChatFocus()
On Error GoTo oops
    Dim i As Integer
    If WindowTitle(GetActiveWindow) <> Me.Caption Then Exit Sub
    
    '''BenMillard''' changed to 1-based control array to match .Tabs 1-based collection:
    For i = 1 To frmGH.rtbChatbox.UBound
        If frmGH.rtbChatbox(i).Visible = True Then
            Call giveFocus(frmGH.rtbChatbox(i)) '''FOCUS
            'frmGH.rtbChatbox(i).SetFocus
            Exit For
        End If
   Next
Exit Sub

oops:
    Call ErrorHandler("giveChatFocus", Err.Description, Erl)
End Sub

Private Sub findGTA2()
If Exists(strGTA2path & TXT_GTA2EXE) = False Then
    displaychat strDestTab, vbRed, "Can't find " & strGTA2path & TXT_GTA2EXE
    Dim strTemp As String
    strTemp = BrowseFile(strGTA2path & TXT_GTA2EXE)
    If Len(strTemp) <= Len(TXT_GTA2EXE) Then
        Exit Sub
    Else
        strGTA2path = Left$(strTemp, Len(strTemp) - Len(TXT_GTA2EXE))
    End If
End If
End Sub

Private Sub txtSlave_Change()
    If Left$(txtSlave.Text, 4) = "Game" Then
        frmGH.Caption = txtSlave.Text
    Else
        displaychat strChannel, strGHColor, txtSlave
    End If
End Sub

