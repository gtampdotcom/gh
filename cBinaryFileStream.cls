VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBinaryFileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sFile As String
Private m_iFile As Integer
Private m_iLen As Long
Private m_iOffset As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Property Get File() As String
   File = m_sFile
End Property
Public Property Let File(ByVal sFile As String)
On Error GoTo oops
   Dispose
   m_sFile = sFile
   Dim lErr As Long
   If (FileExists(m_sFile, lErr)) Then
      m_iFile = FreeFile
      'Open m_sFile For Binary Access Read Lock Write As #m_iFile
      Open m_sFile For Binary Access Read As #m_iFile
      m_iLen = LOF(m_iFile)
   End If
   
Exit Property
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, strGHColor, strErrdesc & " " & sFile
    send "PRIVMSG " & gta2ghbot & " :Let File: " & sFile & " " & strErrdesc & " Line: " & strErrLine
End Property

Private Function FileExists(ByVal sFile As String, ByRef lErr As Long) As Boolean
   
   lErr = 0
   On Error GoTo oops:
   sFile = Replace(sFile, vbQuote, vbNullString)
   Dim sDir As String
   sDir = Dir(sFile, vbHidden + vbNormal + vbSystem + vbReadOnly + vbArchive)
   lErr = Err.Number
   On Error GoTo 0
   
   If (lErr = 0) Then
      If (Len(sDir) > 0) Then
         FileExists = True
      Else
         lErr = 53
      End If
   End If

Exit Function
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error during FileExists: " & strErrdesc & " - Line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :Error during FileExists: " & strErrdesc & " Line: " & strErrLine
End Function

Public Property Get length() As Long
   length = m_iLen
End Property

Public Function Read( _
      Buffer() As Byte, _
      ByVal readSize As Long _
   ) As Long
   
   On Error GoTo oops
   Dim lReadSize As Long
   lReadSize = readSize
   If (m_iOffset + lReadSize >= m_iLen) Then
      readSize = m_iLen - m_iOffset
      If (readSize > 0) Then
         ReDim newBuffer(0 To readSize - 1) As Byte
         Get #m_iFile, , newBuffer
         CopyMemory Buffer(0), newBuffer(0), readSize
      Else
         Dispose
      End If
      m_iOffset = m_iOffset + readSize
   Else
      ' Can read
      Get #m_iFile, , Buffer
      m_iOffset = m_iOffset + readSize
   End If
   Read = readSize
   
   Exit Function
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error during read: " & strErrdesc & " - Line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :Error during read: " & strErrdesc & " Line: " & strErrLine
   
End Function

Public Sub Dispose()
    On Error GoTo oops:
   If (m_iFile) Then
      Close #m_iFile
      m_iFile = 0
   End If
Exit Sub
oops:
    strErrdesc = Err.Description
    strErrLine = Erl
    displaychat strDestTab, vbRed, "Error during dispose " & strErrdesc & " - Line: " & strErrLine
    send "PRIVMSG " & gta2ghbot & " :Error during dispose: " & strErrdesc & " Line: " & strErrLine
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub
