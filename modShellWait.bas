Attribute VB_Name = "modShellWait"
Option Explicit

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Declare Function CreateProcess Lib "Kernel32" _
   Alias "CreateProcessA" _
   (ByVal lpApplicationName As String, _
   ByVal lpCommandLine As String, _
   lpProcessAttributes As Any, _
   lpThreadAttributes As Any, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   lpEnvironment As Any, _
   ByVal lpCurrentDriectory As String, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long

'Private Declare Function OpenProcess Lib "kernel32.dll" _
'   (ByVal dwAccess As Long, _
'   ByVal fInherit As Integer, _
'   ByVal hObject As Long) As Long

'Private Declare Function TerminateProcess Lib "kernel32" _
'   (ByVal hProcess As Long, _
'   ByVal uExitCode As Long) As Long

Private Declare Function CloseHandle Lib "Kernel32" _
   (ByVal hObject As Long) As Long

Const SYNCHRONIZE = 1048576
Const NORMAL_PRIORITY_CLASS = &H20&

Public Function shellandwait(strFilename As String, strFolder As String) As Long
Dim pInfo As PROCESS_INFORMATION
Dim sInfo As STARTUPINFO
Dim sNull As String
Dim lSuccess As Long
Dim lRetValue As Long

sInfo.cb = Len(sInfo)

lSuccess = CreateProcess(sNull, _
                        strFilename, _
                        ByVal 0&, _
                        ByVal 0&, _
                        1&, _
                        NORMAL_PRIORITY_CLASS, _
                        ByVal 0&, _
                        strFolder, _
                        sInfo, _
                        pInfo)

shellandwait = pInfo.dwProcessId
'lRetValue = TerminateProcess(pInfo.hProcess, 0&)
lRetValue = CloseHandle(pInfo.hThread)
lRetValue = CloseHandle(pInfo.hProcess)

End Function
