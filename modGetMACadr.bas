Attribute VB_Name = "modGetMACadr"
Option Explicit

' Declarations needed for GetAdaptersInfo & GetIfTable
'Private Const MIB_IF_TYPE_OTHER                   As Long = 1
'Private Const MIB_IF_TYPE_ETHERNET                As Long = 6
'Private Const MIB_IF_TYPE_TOKENRING               As Long = 9
'Private Const MIB_IF_TYPE_FDDI                    As Long = 15
'Private Const MIB_IF_TYPE_PPP                     As Long = 23
'Private Const MIB_IF_TYPE_LOOPBACK                As Long = 24
'Private Const MIB_IF_TYPE_SLIP                    As Long = 28

'Private Const MIB_IF_ADMIN_STATUS_UP              As Long = 1
'Private Const MIB_IF_ADMIN_STATUS_DOWN            As Long = 2
'Private Const MIB_IF_ADMIN_STATUS_TESTING         As Long = 3
'
'Private Const MIB_IF_OPER_STATUS_NON_OPERATIONAL  As Long = 0
'Private Const MIB_IF_OPER_STATUS_UNREACHABLE      As Long = 1
'Private Const MIB_IF_OPER_STATUS_DISCONNECTED     As Long = 2
'Private Const MIB_IF_OPER_STATUS_CONNECTING       As Long = 3
'Private Const MIB_IF_OPER_STATUS_CONNECTED        As Long = 4
'Private Const MIB_IF_OPER_STATUS_OPERATIONAL      As Long = 5

Private Const MAX_ADAPTER_DESCRIPTION_LENGTH      As Long = 128
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH_p    As Long = MAX_ADAPTER_DESCRIPTION_LENGTH + 4
Private Const MAX_ADAPTER_NAME_LENGTH             As Long = 256
Private Const MAX_ADAPTER_NAME_LENGTH_p           As Long = MAX_ADAPTER_NAME_LENGTH + 4
Private Const MAX_ADAPTER_ADDRESS_LENGTH          As Long = 8
'Private Const DEFAULT_MINIMUM_ENTITIES            As Long = 32
'Private Const MAX_HOSTNAME_LEN                    As Long = 128
'Private Const MAX_DOMAIN_NAME_LEN                 As Long = 128
'Private Const MAX_SCOPE_ID_LEN                    As Long = 256

'Private Const MAXLEN_IFDESCR                      As Long = 256
'Private Const MAX_INTERFACE_NAME_LEN              As Long = MAXLEN_IFDESCR * 2
'Private Const MAXLEN_PHYSADDR                     As Long = 8

' Information structure returned by GetIfEntry/GetIfTable
'Private Type MIB_IFROW
'    wszName(0 To MAX_INTERFACE_NAME_LEN - 1) As Byte    ' MSDN Docs say pointer, but it is WCHAR array
'    dwIndex             As Long
'    dwType              As Long
'    dwMtu               As Long
'    dwSpeed             As Long
'    dwPhysAddrLen       As Long
'    bPhysAddr(MAXLEN_PHYSADDR - 1) As Byte
'    dwAdminStatus       As Long
'    dwOperStatus        As Long
'    dwLastChange        As Long
'    dwInOctets          As Long
'    dwInUcastPkts       As Long
'    dwInNUcastPkts      As Long
'    dwInDiscards        As Long
'    dwInErrors          As Long
'    dwInUnknownProtos   As Long
'    dwOutOctets         As Long
'    dwOutUcastPkts      As Long
'    dwOutNUcastPkts     As Long
'    dwOutDiscards       As Long
'    dwOutErrors         As Long
'    dwOutQLen           As Long
'    dwDescrLen          As Long
'    bDescr As String * MAXLEN_IFDESCR
'End Type

Private Type TIME_t
    aTime As Long
End Type

Private Type IP_ADDRESS_STRING
    IPadrString     As String * 16
End Type

Private Type IP_ADDR_STRING
    AdrNext         As Long
    IpAddress       As IP_ADDRESS_STRING
    IpMask          As IP_ADDRESS_STRING
    NTEcontext      As Long
End Type

' Information structure returned by GetIfEntry/GetIfTable
Private Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName         As String * MAX_ADAPTER_NAME_LENGTH_p
    Description         As String * MAX_ADAPTER_DESCRIPTION_LENGTH_p
    MACadrLength        As Long
    MACaddress(0 To MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    AdapterIndex        As Long
    AdapterType         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    DhcpEnabled         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    CurrentIpAddress    As Long
    IpAddressList       As IP_ADDR_STRING
    GatewayList         As IP_ADDR_STRING
    DhcpServer          As IP_ADDR_STRING
    HaveWins            As Long             ' MSDN Docs say "Bool", but is 4 bytes
    PrimaryWinsServer   As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained       As TIME_t
    LeaseExpires        As TIME_t
End Type

     
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, ByRef source As Any, ByVal numbytes As Long)

Public Declare Function GetAdaptersInfo Lib "IPHLPAPI.dll" (ByRef pAdapterInfo As Any, ByRef pOutBufLen As Long) As Long
Public Declare Function GetNumberOfInterfaces Lib "IPHLPAPI.dll" (ByRef pdwNumIf As Long) As Long
Public Declare Function GetIfEntry Lib "IPHLPAPI.dll" (ByRef pIfRow As Any) As Long
'Private Declare Function GetIfTable Lib "iphlpapi.dll" _
'        (ByRef pIfTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long


'-----------------------------------------------------------------------------------
' Get the system's MAC address(es) via GetAdaptersInfo API function (IPHLPAPI.DLL)
'
' Note: GetAdaptersInfo returns information about physical adapters
'-----------------------------------------------------------------------------------
Public Function GetMACs_AdaptInfo() As String
On Error GoTo oops
    Dim AdapInfo As IP_ADAPTER_INFO, bufLen As Long, sts As Long
    Dim retStr As String, numStructs%, i%, IPinfoBuf() As Byte, srcPtr As Long
    
    ' Get size of buffer to allocate
    sts = GetAdaptersInfo(AdapInfo, bufLen)
    If (bufLen = 0) Then Exit Function
    numStructs = bufLen / Len(AdapInfo)
    retStr = numStructs & " Adapter(s):" & vbCrLf
    
    ' reserve byte buffer & get it filled with adapter information
    ' !!! Don't Redim AdapInfo array of IP_ADAPTER_INFO,
    ' !!! because VB doesn't allocate it contiguous (padding/alignment)
    ReDim IPinfoBuf(0 To bufLen - 1) As Byte
    sts = GetAdaptersInfo(IPinfoBuf(0), bufLen)
    If (sts <> 0) Then Exit Function
    
    ' Copy IP_ADAPTER_INFO slices into UDT structure
    srcPtr = VarPtr(IPinfoBuf(0))
    For i = 0 To numStructs - 1
        If (srcPtr = 0) Then Exit For
'        CopyMemory AdapInfo, srcPtr, Len(AdapInfo)
        CopyMemory AdapInfo, ByVal srcPtr, Len(AdapInfo)
        
        ' Extract Ethernet MAC address
        With AdapInfo
            retStr = retStr & vbCrLf & "[" & i & "] " & sz2string(.Description) _
                    & vbCrLf & vbTab & MAC2String(.MACaddress) & vbCrLf
            
            strMacAddresses = strMacAddresses & MAC2String(.MACaddress) & " "
            
'            Peebug.Print .AdapterIndex
'            Peebug.Print .AdapterName
'            Peebug.Print .AdapterType
'            Peebug.Print .ComboIndex
'            Peebug.Print .CurrentIpAddress
'            Peebug.Print .Description
'            Peebug.Print .DhcpEnabled
'            Peebug.Print .DhcpServer.AdrNext
'            Peebug.Print .DhcpServer.IpAddress.IPadrString
'            Peebug.Print .DhcpServer.IpMask.IPadrString
'            Peebug.Print .DhcpServer.NTEcontext
'            Peebug.Print .GatewayList.AdrNext
'            Peebug.Print .GatewayList.IpAddress.IPadrString
'            Peebug.Print .GatewayList.IpMask.IPadrString
'            Peebug.Print .GatewayList.NTEcontext
'            Peebug.Print .HaveWins
'            Peebug.Print .IpAddressList.AdrNext
'            Peebug.Print .IpAddressList.IpAddress.IPadrString
'            Peebug.Print .IpAddressList.IpMask.IPadrString
'            Peebug.Print .IpAddressList.NTEcontext
'            Peebug.Print .LeaseExpires.aTime
'            Peebug.Print .LeaseObtained.aTime
'            Peebug.Print .MACaddress
'            Peebug.Print .MACadrLength
'            Peebug.Print .Next
'            Peebug.Print .PrimaryWinsServer.AdrNext
'            Peebug.Print .PrimaryWinsServer.IpAddress.IPadrString
'            Peebug.Print .PrimaryWinsServer.IpMask.IPadrString
'            Peebug.Print .PrimaryWinsServer.NTEcontext
'            Peebug.Print .SecondaryWinsServer.AdrNext
'            Peebug.Print .SecondaryWinsServer.IpAddress.IPadrString
'            Peebug.Print .SecondaryWinsServer.IpMask.IPadrString
'            Peebug.Print .SecondaryWinsServer.NTEcontext
'            Peebug.Print "--------------------------------"
'
            
            Dim strLocalIP As String
            Dim strGateway As String
            strLocalIP = Trim(Replace(.IpAddressList.IpAddress.IPadrString, vbNullChar, " "))
            strGateway = Trim(Replace(.GatewayList.IpAddress.IPadrString, vbNullChar, " "))
            
            If InStr(.Description, "Hamachi") = False And InStr(.Description, "VPN") = False And _
            strLocalIP <> "0.0.0.0" And strLocalIP <> vbNullString And _
            strGateway <> "0.0.0.0" And strGateway <> vbNullString Then
                Dim strTemp As String
                strTemp = MAC2String(.MACaddress)
                If strTemp <> "0053450000" Then
                    
                    strMacAddress = strTemp
                    If Left$(strGateway, 7) = "192.168" Or Left$(strGateway, 3) = "10." Then
                        displaychat strChannel, strGHColor, "GTA2 requires that all players forward some ports from their router to their computer."
                        'displaychat strChannel, strGHColor, "If the join button is disabled or if no one can join your game then you need to forward ports."
                        displaychat strChannel, strGHColor, "Click this link to access your router settings: http://" & strGateway
                        displaychat strChannel, strGHColor, "Default router passwords can be found here: http://www.routerpasswords.com"
                        displaychat strChannel, strGHColor, "Forward 47624 TCP and 2300-2400 TCP/UDP to " & strLocalIP & "."
                        displaychat strChannel, strGHColor, "Your LAN MAC address is " & strMacAddress & ". You should set the LAN IP reservation in your router to always give " & strLocalIP & " to " & strMacAddress & "."
                    End If
                End If
            End If

'            If InStr(AdapInfo.Description, "Hamachi") = True Then
'                frmCreateGame.txtHamachiNetwork = .IpAddressList.IpAddress.IPadrString
'                strHamachiIP = frmCreateGame.txtHamachiNetwork
'                'displaychat strDestTab, strConnectionColor, "Hamachi IP: " & strHamachiIP
'            End If
        End With
        srcPtr = AdapInfo.Next
    Next i
    
    ' Return list of MAC address(es)
    GetMACs_AdaptInfo = retStr
    Exit Function
oops:
MsgBox "GetMACs_AdaptInfo: " & Err.Description
End Function


''-----------------------------------------------------------------------------------
'' Get the system's MAC address(es) via GetIfTable API function (IPHLPAPI.DLL)
''
'' Note: GetIfTable returns information also about the virtual loopback adapter
''-----------------------------------------------------------------------------------
'Public Function GetMACs_IfTable() As String
'
'    Dim NumAdapts As Long, nRowSize As Long, i%, retStr As String
'    Dim IfInfo As MIB_IFROW, IPinfoBuf() As Byte, bufLen As Long, sts As Long
'
'
'    ' Get # of interfaces defined (sometimes 1 more than GetIfTable)
'    sts = GetNumberOfInterfaces(NumAdapts)
'
'    ' Get size of buffer to allocate
'    sts = GetIfTable(ByVal 0&, bufLen, 1)
'    If (bufLen = 0) Then Exit Function
'
'    ' reserve byte buffer & get it filled with adapter information
'    ReDim IPinfoBuf(0 To bufLen - 1) As Byte
'    sts = GetIfTable(IPinfoBuf(0), bufLen, 1)
'    If (sts <> 0) Then Exit Function
'
'    NumAdapts = IPinfoBuf(0)
'    nRowSize = Len(IfInfo)
'    retStr = NumAdapts & " Interface(s):" & vbCrLf
'
'    For i = 1 To NumAdapts
'        ' copy one IfRow chunk of byte data into an MIB_IFROW structure
'        Call CopyMemory(IfInfo, IPinfoBuf(4 + (i - 1) * nRowSize), nRowSize)
'
'        ' Take adapter address if correct type
'        With IfInfo
'            retStr = retStr & vbCrLf & "[" & i & "] " & Left$(.bDescr, .dwDescrLen - 1) & vbCrLf
'            If (.dwType = MIB_IF_TYPE_ETHERNET) Then
'                retStr = retStr & vbTab & MAC2String(.bPhysAddr) & vbCrLf
'            End If
'        End With
'    Next i
'
'    GetMACs_IfTable = retStr
'
'End Function


' Convert a byte array containing a MAC address to a hex string
Private Function MAC2String(AdrArray() As Byte) As String
    Dim aStr As String, hexStr As String, i%
    
    For i = 0 To 5
        If (i > UBound(AdrArray)) Then
            hexStr = "00"
        Else
            hexStr = Hex$(AdrArray(i))
        End If
        
        If (Len(hexStr) < 2) Then hexStr = "0" & hexStr
        aStr = aStr & hexStr
        'If (i < 5) Then aStr = aStr & "-"
    Next i
    
    MAC2String = aStr
    
End Function


' Convert a zero-terminated fixed string to a dynamic VB string
Private Function sz2string(ByVal szStr As String) As String
    sz2string = Left$(szStr, InStr(1, szStr, vbNullChar) - 1)
End Function
