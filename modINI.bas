Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "Kernel32" _
           Alias "GetPrivateProfileStringA" _
                 (ByVal sSectionName As String, _
                  ByVal sKeyName As String, _
                  ByVal sDefault As String, _
                  ByVal sReturnedString As String, _
                  ByVal lSize As Long, _
                  ByVal sFileName As String) As Long

'Private Declare Function GetPrivateProfileInt Lib "kernel32" _
'           Alias "GetPrivateProfileIntA" _
'                 (ByVal sSectionName As String, _
'                  ByVal sKeyName As String, _
'                  ByVal lDefault As Long, _
'                  ByVal sFileName As String) As Long

Public Declare Function WriteINI Lib "Kernel32" _
           Alias "WritePrivateProfileStringA" _
                 (ByVal sSectionName As String, _
                  ByVal sKeyName As String, _
                  ByVal sString As String, _
                  ByVal sFileName As String) As Long

Public Function readINI(ByVal strSection As String, ByVal strItem As String, ByVal strFile As String) As String
Dim lngTemp As Long
readINI = Space(255)
lngTemp = GetPrivateProfileString(strSection, strItem, vbNullString, readINI, 255, strFile)
readINI = Trim(Left(readINI, lngTemp))
End Function
