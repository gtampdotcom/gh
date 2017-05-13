VERSION 5.00
Begin VB.Form frmDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Downloading"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2005, 2014 GTAMP.com gtamulti@gmail.com
'License: Do whatever you want with this code. No warranty.

Option Explicit

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SendMessage lngMaster, WM_SETTEXT, 0, ByVal "Game Hunter v" & TXT_GHVER
SendMessage lngMaster, WM_SETTEXT, 0, ByVal "Download cancelled"
End
End Sub
