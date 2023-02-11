Attribute VB_Name = "modTerminate"
Option Explicit

'Used by getVersion Start
Private Declare Function GetVersionExA Lib "Kernel32" _
   (lpVersionInformation As OSVERSIONINFO) As Integer

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long           '1 = Windows 95.
                                  '2 = Windows NT
   szCSDVersion As String * 128
End Type
'Used by getVersion End

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "Kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Const PROCESS_TERMINATE As Long = &H1

Private Declare Function Process32First Lib "Kernel32" ( _
   ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "Kernel32" ( _
   ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" _
   (ByVal Handle As Long) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" _
  (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, _
      ByVal dwProcId As Long) As Long

Private Declare Function EnumProcesses Lib "psapi.dll" _
   (ByRef lpidProcess As Long, ByVal cb As Long, _
      ByRef cbNeeded As Long) As Long

Private Declare Function GetModuleFileNameExA Lib "psapi.dll" _
   (ByVal hProcess As Long, ByVal hModule As Long, _
      ByVal ModuleName As String, ByVal nSize As Long) As Long

Private Declare Function EnumProcessModules Lib "psapi.dll" _
   (ByVal hProcess As Long, ByRef lphModule As Long, _
      ByVal cb As Long, ByRef cbNeeded As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" ( _
   ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Private Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long           ' This process
   th32DefaultHeapID As Long
   th32ModuleID As Long            ' Associated exe
   cntThreads As Long
   th32ParentProcessID As Long     ' This process's parent process
   pcPriClassBase As Long          ' Base priority of process threads
   dwFlags As Long
   szExeFile As String * 260       ' MAX_PATH
End Type

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
'Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
'Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const TH32CS_SNAPPROCESS = &H2&
Private Const hNull = 0


' Terminate the process.
Private Function terminate(target_process_id As Long, Optional appName As String)
Dim target_process_handle As Long
    ' Open the process.
    target_process_handle = OpenProcess( _
        SYNCHRONIZE Or PROCESS_TERMINATE, _
        ByVal 0&, target_process_id)
    If target_process_handle = 0 Then
        displaychat strDestTab, 32896, "Process not found"
        Exit Function
    End If

    ' Terminate the process.
    If TerminateProcess(target_process_handle, 0&) = 0 Then
        displaychat strDestTab, 32896, "Failed to terminate process " & target_process_id
    Else
        displaychat strDestTab, 32896, "Process " & target_process_id & " " & appName & " was terminated"
    End If

    ' Close the process.
    CloseHandle target_process_handle
End Function

Public Function FindProcess(sAppName As String, Optional blnKill As Boolean, Optional lngID As Long) As Boolean
On Error GoTo oops
sAppName = LCase$(sAppName)
      
Select Case bln98
    Case True 'Windows 95/98

    Dim f As Long, sName As String
    Dim hSnap As Long, proc As PROCESSENTRY32
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If hSnap = hNull Then Exit Function
    proc.dwSize = Len(proc)
    ' Iterate through the processes
    f = Process32First(hSnap, proc)
    Do While f
        sName = StrZToStr(proc.szExeFile)
            
        If lngPID > 0 Then
             If proc.th32ProcessID = lngID Then
                 FindProcess = True
                 Exit Function
             End If
        Else
            If InStr(LCase$(sName), sAppName) Then
                 FindProcess = True
                 If blnKill = True Then
                     Call terminate(proc.th32ProcessID, sName)
                 Else
                     Exit Function
                 End If
            End If
        End If
        f = Process32Next(hSnap, proc)
    Loop

      Case Else 'Windows NT

         Dim cb As Long
         Dim cbNeeded As Long
         Dim NumElements As Long
         Dim ProcessIDs() As Long
         Dim cbNeeded2 As Long
         Dim Modules(1 To 200) As Long
         Dim lRet As Long
         Dim ModuleName As String
         Dim nSize As Long
         Dim hProcess As Long
         Dim i As Long
         Dim j As Integer
         'Get the array containing the process id's for each process object
         cb = 8
         cbNeeded = 96
         Do While cb <= cbNeeded
            cb = cb * 2
            ReDim ProcessIDs(cb / 4) As Long
            lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
         Loop
         NumElements = cbNeeded / 4

         For i = 1 To NumElements
            'Get a handle to the Process
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
               Or PROCESS_VM_READ, 0, ProcessIDs(i))
            'Got a Process handle
            If hProcess <> 0 Then
                If lngID > 0 Then
                    If ProcessIDs(i) = lngID Then
                        FindProcess = True
                        lRet = CloseHandle(hProcess)
                        Exit Function
                    End If
                Else
                    'Get an array of the module handles for the specified
                    'process
                    lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                                 cbNeeded2)
                    'If the Module Array is retrieved, Get the ModuleFileName
                    If lRet <> 0 Then
                        ModuleName = Space(MAX_PATH)
                        nSize = 500
                        lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                                       ModuleName, nSize)
                        ModuleName = Replace(Trim(ModuleName), vbNullChar, vbNullString)
                        
                        j = InStrRev(ModuleName, "\", , vbBinaryCompare)
                        If j > 0 And Len(ModuleName) > j + 4 Then
                            If LCase$(Mid$(ModuleName, j + 1, 666)) = sAppName Or _
                            LCase$(Mid$(ModuleName, j + 1, Len(ModuleName) - (j + 4))) = sAppName Then
                                'Close the handle to the process
                                lRet = CloseHandle(hProcess)
                                
                                FindProcess = True
                                If blnKill = True Then
                                    Call terminate(ProcessIDs(i), ModuleName)
                                Else
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
         'Close the handle to the process
         lRet = CloseHandle(hProcess)
         Next

      End Select
      Exit Function
oops:
    Static blnFindProcessErr As Boolean
    If blnFindProcessErr = False Then
        strErrLine = Erl
        strErrdesc = Err.Description
        send "PRIVMSG " & gta2ghbot & " :Find process error: " & strErrdesc & " " & strErrLine
        strErrdesc = Err.Description
        displaychat strDestTab, strGHColor, "Find process error: " & strErrdesc & " " & strErrLine
    End If
    blnFindProcessErr = True
End Function

'Get Windows version
Public Function getVersion() As Boolean
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
        
    'Computer\HKEY_LOCAL_MACHINE\SYSTEM\Setup\MoSetup\Volatile
    Dim cr As New cRegistry
    
    With cr
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "SYSTEM\Setup\MoSetup\Volatile"
        .ValueKey = "DownlevelProductName"
        strOSV = .Value
    End With
        
    'strOSV = osinfo.dwMajorVersion & "." & osinfo.dwMinorVersion & "." & osinfo.dwBuildNumber
    If osinfo.dwPlatformId = 1 Then
        getVersion = True
    Else
        getVersion = False
    End If
End Function
