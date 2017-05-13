Attribute VB_Name = "modSystemInfo"
'http://satepadangajo.blogspot.com/2010/10/get-ip-address.html
'16:15 10/11/2010 change return type of GetIpAddr() to Collection and LocalAdapters() to ADAPTER_INFO array
'11:08 09/10/2010 get LocalIP from network adapters or IpAddrTable
'
'Copyright © 2010 RENO

Option Compare Text
Option Explicit
#Const DEBUG_ = True

'util function
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, source As Any, ByVal length As Long)

'-------------------------------------------
'Struct,const required for GetIpAddrTable()
'-------------------------------------------
Private Type MIB_IPADDRROW
    dwAddr(3)   As Byte
    dwIndex     As Long
    dwMask      As Long
    dwBCastAddr As Long
    dwReasmSize As Long
    unused1     As Integer
    wType       As Integer
End Type
Private Declare Function GetIpAddrTable Lib "IPHLPAPI.dll" (pIpAddrTable As Any, pdwSize As Long, ByVal bOrderSort As Long) As Long
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122
Private Const NO_ERROR                  As Long = 0

'-------------------------------------------------
'Struct,const required for GetAdaptersAddresses()
'-------------------------------------------------
Public Enum NET_IF_CONNECTION_TYPE
    NET_IF_CONNECTION_DEDICATED = 1
    NET_IF_CONNECTION_PASSIVE
    NET_IF_CONNECTION_DEMAND
    NET_IF_CONNECTION_MAXIMUM
End Enum

Private Type SOCKET_ADDRESS '8byte
    lpSockaddr      As Long         'pointer to sockaddr struct
    iSockaddrLength As Long
End Type
Private Type sockaddr_in            'ipv4 sockaddr
    sin_family      As Integer
    sin_port        As Integer
    sin_addr(3)     As Byte
    sin_zero(7)     As Byte
End Type

Private Type guid   '16bytes
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(7)        As Byte
End Type
Private Type IP_ADAPTER_ADDRESSES   '376byte, current 374byte, still missing 2byte?
    length                  As Long
    IfIndex                 As Long
    pNext                   As Long     'pointer to next IP_ADAPTER_ADDRESSES
    pAdapterName            As Long     'pointer to adapter name
    pFirstUnicastAddress    As Long     'pointer to first IP_ADAPTER_UNICAST_ADDRESS
    pFirstAnycastAddress    As Long
    pFirstMulticastAddress  As Long
    pFirstDnsServerAddress  As Long
    pDnsSuffix              As Long
    pDescription            As Long
    pFriendlyName           As Long
    PhysicalAddress(7)      As Byte
    PhysicalAddressLength   As Long
    Flags                   As Long
    Mtu                     As Long
    IfType                  As Long
    OperStatus              As Byte
    Ipv6IfIndex             As Long
    ZoneIndices(15)         As Long
    pFirstPrefix            As Long     'pointer to IP_ADAPTER_PREFIX
    TransmitLinkSpeed       As Currency 'bits/s multiply by 10k, currency store data with 4dp
    ReceiveLinkSpeed        As Currency
    pFirstWinsServerAddress As Long
    pFirstGatewayAddress    As Long
    Ipv4Metric              As Long
    Ipv6Metric              As Long
    Luid                    As Currency
    Dhcpv4Server            As SOCKET_ADDRESS   'struct SOCKET_ADDRESS
    CompartmentId           As Long             'NET_IF_COMPARTMENT_ID
    NetworkGuid             As guid
    ConnectionType          As NET_IF_CONNECTION_TYPE   'LONG
    TunnelType              As Long
    Dhcpv6Server            As SOCKET_ADDRESS   'struct SOCKET_ADDRESS
    Dhcpv6ClientDuid(129)   As Byte             'MAX_DHCPV6_DUID_LENGTH=130
    Dhcpv6ClientDuidLength  As Long
    Dhcpv6Iaid              As Long
    pFirstDnsSuffix         As Long             'pointer to IP_ADAPTER_DNS_SUFFIX
End Type

Private Type IP_ADAPTER_UNICAST_ADDRESS '48bytes
    length              As Long
    Flags               As Long
    pNext               As Long         'pointer to next IP_ADAPTER_UNICAST_ADDRESS
    Address             As SOCKET_ADDRESS
    PrefixOrigin        As Long
    SuffixOrigin        As Long
    DadState            As Long
    ValidLifetime       As Long
    PreferredLifetime   As Long
    LeaseLifetime       As Long
    OnLinkPrefixLength  As Byte         'UINT8 The length, in bits, of the prefix or network part of the IP address.
End Type
Private Declare Function GetAdaptersAddresses Lib "IPHLPAPI.dll" _
    (ByVal Family As Long, ByVal Flags As Long, ByVal Reserved As Long, AdapterAddresses As Any, SizePointer As Long) As Long
Private Const IfOperStatusUp            As Byte = 1
Private Const IF_TYPE_SOFTWARE_LOOPBACK As Long = 24
Private Const AF_INET                   As Long = 2     'ipv4
Private Const AF_INET6                  As Long = 23    'ipv6
Private Const AF_UNSPEC                 As Long = 0     'both ipv4 and ipv6
Private Const GAA_FLAG_INCLUDE_PREFIX   As Long = &H10
Private Const ERROR_BUFFER_OVERFLOW     As Long = 111

'custom datatype for return, modify as needed
Public Type ADAPTER_INFO
    FriendlyName        As String
    Address             As String
    MAC                 As String
    ConnectionType      As NET_IF_CONNECTION_TYPE
    IfType              As Long
End Type

Public Function LocalIP(Optional UseGetAdaptersAddressesAPI As Boolean = True) As String
    If UseGetAdaptersAddressesAPI Then
        Dim ai() As ADAPTER_INFO
        ai = LocalAdapters()
        If (Not ai) <> -1& Then
            Dim i: For i = 0 To UBound(ai)
                If ai(i).IfType <> IF_TYPE_SOFTWARE_LOOPBACK Then
                    LocalIP = LocalIP & ai(i).FriendlyName & " = " & ai(i).Address
                    If ai(i).MAC <> "" Then LocalIP = LocalIP & " (MAC:" & ai(i).MAC & ")"
                    LocalIP = LocalIP & vbCrLf
                End If
            Next
        End If
    Else 'use GetIpAddr()
        Dim ip: For Each ip In GetIpAddr()
            If InStr(1, ip, "127.0.0.1") = 0 Then LocalIP = LocalIP & ip & vbCrLf
        Next
    End If
End Function

Private Sub test()
    'Debug.Print "SystemInfo " & vbCrLf & String(15, "-")
    'Debug.Print SystemInfo()
End Sub

'Public Function SystemInfo() As String
''   retrieve system information from registry
'    SystemInfo = GetRegistryValue(HKEY_LOCAL_MACHINE, "HARDWARE\Description\System\Bios", "SystemManufacturer") & " " & _
'                 GetRegistryValue(HKEY_LOCAL_MACHINE, "HARDWARE\Description\System\Bios", "SystemProductName") & vbCrLf & _
'                 GetRegistryValue(HKEY_LOCAL_MACHINE, "HARDWARE\Description\System\CentralProcessor\0", "ProcessorNameString")
'End Function

Public Function LocalAdapters(Optional Family As Long = AF_INET, Optional StatusUpOnly As Boolean = True) As ADAPTER_INFO()
'   http://msdn.microsoft.com/en-us/library/aa365915
'   Minimum supported client: Windows XP (from MSDN)
'   Get local network adapters information via GetAdaptersAddresses() API
On Error GoTo ErrHandler
    Dim b()         As Byte: ReDim b(0)
    Dim n           As Long
    Dim iaa         As IP_ADAPTER_ADDRESSES
    Dim iaua        As IP_ADAPTER_UNICAST_ADDRESS
    Dim sa          As sockaddr_in
    Dim ai()        As ADAPTER_INFO
    
    'get buffer size and redim the buffer needed
    If GetAdaptersAddresses(Family, GAA_FLAG_INCLUDE_PREFIX, 0&, b(0), n) = ERROR_BUFFER_OVERFLOW Then ReDim b(n - 1)
    'now get the data
    If GetAdaptersAddresses(Family, GAA_FLAG_INCLUDE_PREFIX, 0&, b(0), n) = NO_ERROR Then
        'read the buffer into IP_ADAPTER_ADDRESSES struct
        RtlMoveMemory iaa, b(0), LenB(iaa)

        'enumerate all the adapters
        Do While True
            #If DEBUG_ Then
            Debug.Print "Length of IP_ADAPTER_ADDRESS struct : " & iaa.length & " Current Struct Len=" & LenB(iaa)
            Debug.Print "IfIndex (IPv4 interface) : " & iaa.IfIndex
            Debug.Print "AdapterName : " & ReadPointerAsString(iaa.pAdapterName, False)
            Debug.Print "DNS suffix : " & ReadPointerAsString(iaa.pDnsSuffix)
            Debug.Print "Description : " & ReadPointerAsString(iaa.pDescription)
            Debug.Print "Friendly Name : " & ReadPointerAsString(iaa.pFriendlyName)
            Debug.Print "Flags : " & iaa.Flags
            Debug.Print "Mtu : " & iaa.Mtu
            Debug.Print "IfType : " & iaa.IfType
            Debug.Print "OperStatus : " & iaa.OperStatus
            Debug.Print "TransmitLinkSpeed : " & iaa.TransmitLinkSpeed * 10 & "kbps"
            Debug.Print "ReceiveLinkSpeed : " & iaa.ReceiveLinkSpeed * 10 & "kbps"
            Debug.Print "Ipv6IfIndex (IPv6 interface): " & iaa.Ipv6IfIndex
            Debug.Print "Ipv4Metric : " & iaa.Ipv4Metric
            Debug.Print "Ipv6Metric : " & iaa.Ipv6Metric
            Debug.Print "ConnectionType : " & iaa.ConnectionType
            Debug.Print "TunnelType : " & iaa.TunnelType
            Debug.Print "PhysicalAddressLength : " & iaa.PhysicalAddressLength
            Debug.Print "NetworkGuid Data1 : " & iaa.NetworkGuid.Data1
            Debug.Print "Dhcpv4Server length : " & iaa.Dhcpv4Server.iSockaddrLength
            Debug.Print "Dhcpv6ClientDuidLength : " & iaa.Dhcpv6ClientDuidLength
            #End If
            
            If StatusUpOnly And iaa.OperStatus = IfOperStatusUp Then
                'read mac address
                Dim MAC As String: MAC = ""
                Dim i: For i = 0 To iaa.PhysicalAddressLength - 1
                    MAC = MAC & Right("0" & Hex$(iaa.PhysicalAddress(i)), 2) & "-"
                Next
                If Len(MAC) > 0 Then MAC = Left(MAC, Len(MAC) - 1)
                'Debug.Print "PhysicalAddress : " & mac
                
                'read each IP Address of adapter
                RtlMoveMemory iaua, ByVal iaa.pFirstUnicastAddress, LenB(iaua)
                Do While True
                    RtlMoveMemory sa, ByVal iaua.Address.lpSockaddr, LenB(sa)

                    If (Not ai) = -1& Then ReDim ai(0) Else ReDim Preserve ai(UBound(ai) + 1)
                    ai(UBound(ai)).FriendlyName = ReadPointerAsString(iaa.pFriendlyName)
                    ai(UBound(ai)).Address = sa.sin_addr(0) & "." & sa.sin_addr(1) & "." & sa.sin_addr(2) & "." & sa.sin_addr(3)
                    If MAC <> "" Then ai(UBound(ai)).MAC = MAC
                    ai(UBound(ai)).ConnectionType = iaa.ConnectionType
                    ai(UBound(ai)).IfType = iaa.IfType
                    
                    'move to next unicast address
                    If iaua.pNext = 0 Then Exit Do
                    RtlMoveMemory iaua, ByVal iaua.pNext, LenB(iaua)
                Loop
            End If
            
            'move to next adapter
            If iaa.pNext = 0 Then Exit Do
            RtlMoveMemory iaa, ByVal iaa.pNext, LenB(iaa)
        Loop
    End If
ExitHere:
    LocalAdapters = ai
    Exit Function
ErrHandler:
    handleerror "LocalAdapters()"
    Resume ExitHere
End Function

Public Function ReadPointerAsString(ptr As Long, Optional Unicode = True) As String
'   helper function to read PCHAR, PWCHAR and return the result as vb string
On Error GoTo ErrHandler
    If ptr <> 0 Then
        Dim s As String * 512
        RtlMoveMemory ByVal s, ByVal ptr, Len(s)
        If Unicode Then s = StrConv(s, vbFromUnicode)
        ReadPointerAsString = TrimNull(s)
    End If
ExitHere:
    Exit Function
ErrHandler:
    handleerror "ReadPointerAsString()"
    Resume ExitHere
End Function

Public Function GetIpAddr() As Collection
'   http://msdn.microsoft.com/en-us/library/Aa365949
'   Minimum supported client: Windows 2000 Professional (from MSDN)
'   Get local ip information via GetIpAddrTable() API
On Error GoTo ErrHandler
    Dim b() As Byte: ReDim b(0)
    Dim n As Long
    
    'get buffer size and redim the buffer needed
    If GetIpAddrTable(b(0), n, 0) = ERROR_INSUFFICIENT_BUFFER Then ReDim b(n)
    'now get the table
    If GetIpAddrTable(b(0), n, 0) = NO_ERROR Then
        Set GetIpAddr = New Collection
        'retrieve header for number of entries, 4 byte
        RtlMoveMemory n, b(0), LenB(n)
        'read each row of the ip addresses
        Dim ip() As MIB_IPADDRROW: ReDim ip(n)
        Dim i: For i = 0 To n - 1
            RtlMoveMemory ip(i), b(4 + i * LenB(ip(0))), LenB(ip(0))
            GetIpAddr.Add ip(i).dwAddr(0) & "." & ip(i).dwAddr(1) & "." & ip(i).dwAddr(2) & "." & ip(i).dwAddr(3)
        Next
    End If
ExitHere:
    Exit Function
ErrHandler:
    handleerror "GetIpAddr()"
    Resume ExitHere
End Function

Private Function TrimNull(s As String) As String
    Dim i As Long
    i = InStr(s, vbNullChar)
    If i = 0 Then
        TrimNull = s
    Else
        TrimNull = Left$(s, i - 1)
    End If
End Function

Private Function handleerror(strErr As String)
Debug.Print "Error: " & strErr
End Function
