Attribute VB_Name = "modMac"
'http://www.experts-exchange.com/Programming/Languages/Q_26236625.html

Option Explicit

Private Const HEAP_ZERO_MEMORY = &H8&
Private Const ERROR_BUFFER_OVERFLOW = 111
Private Const GAA_FLAG_INCLUDE_PREFIX = &H10&
Private Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Private Const MAX_ADAPTER_NAME_LENGTH = 256
Private Const AF_UNSPEC = 0
Private Const NO_ERROR = 0

Private Declare Function GetAdaptersAddresses Lib "Iphlpapi" ( _
  ByVal Family As Long, _
  ByVal Flags As Long, _
  ByVal Reserved As Long, _
  ByVal AdapterAddresses As Long, _
  ByRef SizePointer As Long) As Long

Private Declare Function GetProcessHeap Lib "Kernel32" ( _
    ) As Long
    
Private Declare Function HeapAlloc Lib "Kernel32" ( _
    ByVal hHeap As Long, _
    ByVal dwFlags As Long, _
    ByVal dwBytes As Long) As Long

Private Declare Function HeapReAlloc Lib "Kernel32" ( _
    ByVal hHeap As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMem As Long, _
    ByVal dwBytes As Long) As Long

Private Declare Function HeapFree Lib "Kernel32" ( _
    ByVal hHeap As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMem As Long) As Long
    
Private Declare Sub RtlMoveMemory Lib "Kernel32" ( _
    Destination As Any, _
    source As Any, _
    ByVal length As Long)
    
Private Declare Function lstrlenW Lib "Kernel32" ( _
  ByVal ptr As Long) As Long

Private Enum IF_TYPE
IF_TYPE_OTHER = 1                     'Some other type of network interface.
IF_TYPE_ETHERNET_CSMACD = 6           'An Ethernet network interface.
IF_TYPE_ISO88025_TOKENRING = 9        'A token ring network interface.
IF_TYPE_PPP = 23                      'A PPP network interface.
IF_TYPE_SOFTWARE_LOOPBACK = 24        'A software loopback network interface.
IF_TYPE_ATM = 37                      'An ATM network interface.
IF_TYPE_IEEE80211 = 71                'An IEEE 802.11 wireless network interface. On Windows Vista and later, wireless network cards are reported as IF_TYPE_IEEE80211. On earlier versions of Windows, wireless network cards are reported as IF_TYPE_ETHERNET_CSMACD.
IF_TYPE_TUNNEL = 131                  'A tunnel type encapsulation network interface.
IF_TYPE_IEEE1394 = 144                'An IEEE 1394 (Firewire) high performance serial bus network interface.
End Enum

Private Enum IF_OPER_STATUS
IfOperStatusUp = 1                    'The interface is up and able to pass packets.
IfOperStatusDown = 2                  'The interface is down and not in a condition to pass packets. The IfOperStatusDown state has two meanings, depending on the value of AdminStatus member. If AdminStatus is not set to NET_IF_ADMIN_STATUS_DOWN and ifOperStatus is set to IfOperStatusDown then a fault condition is presumed to exist on the interface. If AdminStatus is set to IfOperStatusDown, then ifOperStatus will normally also be set to IfOperStatusDown or IfOperStatusNotPresent and there is not necessarily a fault condition on the interface.
IfOperStatusTesting = 3               'The interface is in testing mode.
IfOperStatusUnknown = 4               'The operational status of the interface is unknown.
IfOperStatusDormant = 5               'The interface is not actually in a condition to pass packets (it is not up), but is in a pending state, waiting for some external event. For on-demand interfaces, this new state identifies the situation where the interface is waiting for events to place it in the IfOperStatusUp state.
IfOperStatusNotPresent = 6            'A refinement on the IfOperStatusDown state which indicates that the relevant interface is down specifically because some component (typically, a hardware component) is not present in the managed system.
IfOperStatusLowerLayerDown = 7        'A refinement on the IfOperStatusDown state. This new state indicates that this interface runs on top of one or more other interfaces and that this interface is down specifically because one or more of these lower-layer interfaces are down.
End Enum

Private Type LARGE_INTEGER
LowPart As Long
HighPart As Long
End Type

Private Type IP_ADAPTER_ADDRESSES
Alignment As LARGE_INTEGER
Next As Long
AdapterName As Long                   'PCHAR
FirstUnicastAddress As Long           'IP_ADAPTER_UNICAST_ADDRESS
FirstAnycastAddress As Long           'IP_ADAPTER_ANYCAST_ADDRESS
FirstMulticastAddress As Long         'IP_ADAPTER_MULTICAST_ADDRESS
FirstDnsServerAddress As Long         'IP_ADAPTER_DNS_SERVER_ADDRESS
DnsSuffix As Long                     'PWCHAR
Description As Long                   'PWCHAR
FriendlyName As Long                  'PWCHAR
PhysicalAddress(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
PhysicalAddressLength As Long
Flags As Long
Mtu As Long
IfType As IF_TYPE
OperStatus As IF_OPER_STATUS
End Type

Private Function PointerStringW(ByVal ptr As Long) As String
  Dim Buffer()               As Byte
  Dim lpSize                 As Long
  lpSize = lstrlenW(ptr) * 2
  If lpSize <> 0 Then
    ReDim Buffer(lpSize) As Byte
    RtlMoveMemory Buffer(0), ByVal ptr, lpSize
    PointerStringW = Buffer
  End If
  Erase Buffer
End Function

Public Sub QueryAdaptersAddresses()

  Dim ipaa                  As IP_ADAPTER_ADDRESSES
  Dim pAdapterAddresses     As Long
  Dim outBufLen             As Long
  Dim Flags                 As Long
  Dim Family                As Long
  Dim dwRetVal              As Long
  Dim dwIndex               As Long
  
  '   Initialize flags
  Flags = GAA_FLAG_INCLUDE_PREFIX
  Family = AF_UNSPEC
  outBufLen = 0
  
  '   Allocate a small buffer (32 bytes)
  pAdapterAddresses = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, 32)
  
  '   Pass a small buffer size as indicated in the SizePointer parameter in the first call
  '   to the GetAdaptersAddresses function, so the function will fail with ERROR_BUFFER_OVERFLOW.
  dwRetVal = GetAdaptersAddresses(Family, Flags, 0, pAdapterAddresses, outBufLen)
  
  '   ReAllocate the buffer with the size needed.
  If dwRetVal = ERROR_BUFFER_OVERFLOW Then
    pAdapterAddresses = HeapReAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, pAdapterAddresses, outBufLen)
  End If
  
  '   Make the second call passing in the correct buffer size.
  dwRetVal = GetAdaptersAddresses(Family, Flags, 0, pAdapterAddresses, outBufLen)
  
  If (dwRetVal = NO_ERROR) Then
  
  '   The first IP_ADAPTER_ADDRESSES.
    RtlMoveMemory ipaa, ByVal pAdapterAddresses, Len(ipaa)
    Debug.Print PointerStringW(ipaa.FriendlyName); ipaa.OperStatus; ipaa.IfType; ipaa.FirstDnsServerAddress
    
  '   Walk the buffer for additional IP_ADAPTER_ADDRESSES.
    While ipaa.Next <> 0
      RtlMoveMemory ipaa, ByVal ipaa.Next, Len(ipaa)
      Debug.Print PointerStringW(ipaa.FriendlyName); ipaa.OperStatus; ipaa.IfType; ipaa.FirstAnycastAddress
    Wend
    
  End If
  '   Don't attempt to access the structures pointers
  '   after the memory has been released.
  '
  '   Free the memory
  HeapFree GetProcessHeap, 0, pAdapterAddresses


End Sub

