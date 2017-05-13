Attribute VB_Name = "modBrowse"
'Option Explicit
'
'Dim cR As New cRegistry
'
'
''BROWSEINFO.ulFlags values - copied from http://vbnet.mvps.org/index.html?code/browse/browseadv.htm
'Private Const BIF_RETURNONLYFSDIRS   As Long = &H1 'only file system directories
''Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2  'no network folders below domain level
''Private Const BIF_STATUSTEXT As Long = &H4         'include status area for callback
''Private Const BIF_RETURNFSANCESTORS As Long = &H8  'only return file system ancestors
'Private Const BIF_EDITBOX As Long = &H10           'add edit box
'Private Const BIF_NEWDIALOGSTYLE As Long = &H40    'use the new dialog layout
''Private Const BIF_UAHINT As Long = &H100
'Private Const BIF_NONEWFOLDERBUTTON As Long = &H200 'hide new folder button
''Private Const BIF_NOTRANSLATETARGETS As Long = &H400 'return lnk file
'Private Const BIF_USENEWUI As Long = BIF_NEWDIALOGSTYLE Or BIF_EDITBOX
''Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000 'only return computers
''Private Const BIF_BROWSEFORPRINTER As Long = &H2000 'only return printers
''Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000 'browse for everything
''Private Const BIF_SHAREABLE As Long = &H8000 'sharable resources, requires BIF_USENEWUI
'
'Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, _
'    ByVal nFolder As Long, ppidl As Long) As Long
'Public Const CSIDL_DRIVES = &H11
''   Program Files folder. A typical path is C:\Program Files.
''   Version 5
'Public Const CSIDL_PROGRAM_FILES = &H2A
'Public Type BrowseInfo
'    hwndOwner As Long
'    pIDLRoot As Long
'    pszDisplayName As String
'    lpszTitle As String
'    ulFlags As Long
'    lpfn As Long
'    lParam As Long
'    iImage As Long
'End Type
'
'Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
'    (lpbi As BrowseInfo) As Long
'
'Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
'
'Declare Function SHGetPathFromIDList Lib "shell32.dll" _
'  Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
'  ByVal pszPath As String) As Long
'
'Public Type SH_ITEMID
'    cb As Long
'    abID As Byte
'End Type
'
'Public Type ITEMIDLIST
'    mkid As SH_ITEMID
'End Type
'
'Private Function Address_Of(ByVal n As Long) As Long
'    Address_Of = n
'End Function
'
'Private Function BrowseCallbackProc(ByVal hwnd As Long, _
'                                    ByVal uMsg As Long, _
'                                    ByVal lParam As Long, _
'                                    ByVal lpData As Long) As Long
'    Const WM_USER = &H400&
'    Const BFFM_INITIALIZED = 1
'    Const BFFM_SETSELECTIONA = (WM_USER + 102)
'
'    Dim default_path() As Byte
'
'    If uMsg = BFFM_INITIALIZED Then
'        default_path = StrConv(strGTA2path, vbFromUnicode)
'        SendMessage hwnd, BFFM_SETSELECTIONA, 1&, ByVal VarPtr(default_path(0))
'    End If
'End Function
'
'Public Function BrowseFolder() As Boolean
'    On Error GoTo oops
'    Dim bi As BrowseInfo  ' structure passed to the function
'    Dim pidl As Long  ' PIDL to the user's selection
'    Dim physpath As String  ' string used to temporarily hold the physical path
'    Dim RetVal As Long  ' return value
'    ' Initialize the structure to be passed to the function.
'    physpath = Space(260)
'
'    RetVal = SHGetPathFromIDList(pidl, physpath) 'do I need this line?
'
'    With bi
'        ' The owner of the dialog box.
'        .hwndOwner = frmGH.hwnd
'        ' Make room in the buffer to get the [virtual] folder's display name.
'        .pszDisplayName = Space(260)
'        .lpszTitle = "Where is your GTA2 folder?"
'        .ulFlags = BIF_USENEWUI + BIF_NONEWFOLDERBUTTON + BIF_RETURNONLYFSDIRS
'        .lpfn = Address_Of(AddressOf BrowseCallbackProc)
'    End With
'
'    ' Open the Browse for Folder dialog box.
'    pidl = SHBrowseForFolder(bi)
'
'    If pidl <> 0 Then
'        ' Remove the empty space from the display name variable.
'        bi.pszDisplayName = Left(bi.pszDisplayName, InStr(bi.pszDisplayName, vbNullChar) - 1)
'        ' If the folder is not a virtual folder, display its physical location.
'        physpath = Space(260)
'        RetVal = SHGetPathFromIDList(pidl, physpath)
'        If RetVal > 0 Then
'            ' Remove the empty space and display the result.
'            physpath = Left(physpath, InStr(physpath, vbNullChar) - 1)
'            If Right(physpath, 1) <> "\" Then physpath = physpath & "\"
'        End If
'        ' Free the pidl returned by the function.
'        CoTaskMemFree pidl
'    End If
'
'    ' Whether successful or not, free the PIDL which was used to
'    ' identify the My Computer virtual folder.
'    CoTaskMemFree bi.pIDLRoot
'
'    If pidl <> 0 Then 'if the OK button is pushed
'        BrowseFolder = True
'        If RetVal <> 0 And Exists(physpath & TXT_GTA2EXE) = True Then
'            strGTA2path = physpath
'
'            With cR
'                .ClassKey = HKEY_CURRENT_USER
'                .SectionKey = "SOFTWARE\GTA2 Game Hunter"
'                .ValueKey = "GTA2Folder"
'                .ValueType = REG_SZ
'                .Value = strGTA2path
'            End With
'
'            modVersion.DetectGTA2version
'        Else
'            Call BrowseFolder
'        End If
'    Else
'        BrowseFolder = False
'    End If
'
'    Set cR = Nothing
'    Exit Function
'
'oops:
'    strErrdesc = Err.Description
'      strErrLine = Erl
'    displaychat strDestTab, vbRed, "Error reading folder: " & strErrdesc & " Line: " & strErrLine
'    Set cR = Nothing
'End Function
