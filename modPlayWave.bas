Attribute VB_Name = "modPlayWave"
Option Explicit

Public Const SND_ASYNC As Long = &H1
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
