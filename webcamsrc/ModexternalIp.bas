Attribute VB_Name = "ModexternalIp"
Option Explicit

Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function


Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

'---

Public Function GetInternetIP(pboolReturnExternalIP As Boolean) As String
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ipaddress() As Byte
    Dim iCounter As Integer
    Dim strIPaddress As String
    Dim strCheckIP As String
    Dim strInternIP As String
    Dim strExternIP As String
    
    GetInternetIP = ""

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        MsgBox "Socket-Error with Winsock.dll."
    Else
        hostname = Trim$(hostname)
    
        hostent_addr = gethostbyname(hostname)
    
        If hostent_addr = 0 Then
            MsgBox "There's an error with Winsock.dll."
            
        Else
    
            RtlMoveMemory host, hostent_addr, LenB(host)
            RtlMoveMemory hostip_addr, host.hAddrList, 4
           
            Do
                ReDim temp_ipaddress(1 To host.hLength)
                RtlMoveMemory temp_ipaddress(1), hostip_addr, host.hLength
        
        
                For iCounter = 1 To host.hLength
                    strIPaddress = strIPaddress & temp_ipaddress(iCounter) & "."
                Next
                strIPaddress = Mid$(strIPaddress, 1, Len(strIPaddress) - 1)
            
            
                strInternIP = strCheckIP
                strExternIP = strIPaddress
                strCheckIP = strIPaddress
        
                
                
                host.hAddrList = host.hAddrList + LenB(host.hAddrList)
                RtlMoveMemory hostip_addr, host.hAddrList, 4
                
                strIPaddress = ""
            Loop While (hostip_addr <> 0)
            
            If Trim(strInternIP) = "" Then ' same as External
                strInternIP = strExternIP
            End If
            
            If Trim(strExternIP) = "" Then 'just for sure
                strExternIP = strInternIP  ' no one knows, what
            End If                         ' micrososft does next :>
            
            GetInternetIP = strInternIP
            If pboolReturnExternalIP = True Then
                GetInternetIP = strExternIP
            End If
         
         End If
    End If

End Function


