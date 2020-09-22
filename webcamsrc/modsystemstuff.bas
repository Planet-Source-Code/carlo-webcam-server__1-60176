Attribute VB_Name = "modsystemstuff"


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
  
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
  
  
  Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
  Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
  Public Const REG_SZ = 1
  Public Const REG_DWORD = 4



Public Function GetString(hKey As Long, strPath As String, strValue As String)
Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
R = RegOpenKey(hKey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            GetString = Left$(strBuf, intZeroPos - 1)
        Else
            GetString = strBuf
        End If
    End If
End If
End Function




Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
If lRegResult <> ERROR_SUCCESS Then
'there is a problem
End If
lRegResult = RegCloseKey(hCurKey)
End Sub



Public Function WindowsDirectory() As String
Dim WinPath As String
Dim Temp
WinPath = String(145, Chr(0))
Temp = GetWindowsDirectory(WinPath, 145)
WindowsDirectory = Left(WinPath, InStr(WinPath, Chr(0)) - 1)
End Function

Public Function SystemDirectory() As String
Dim SysPath As String
Dim Temp
SysPath = String(145, Chr(0))
Temp = GetSystemDirectory(SysPath, 145)
SystemDirectory = Left(SysPath, InStr(SysPath, Chr(0)) - 1)
End Function





Sub Main()
  
Dim ff As Long
  ff = FreeFile
   If App.PrevInstance Then
       ' MsgBox "already running", vbOKOnly
        End
   End If
     frmweb.lisSock.Close
       frmweb.lisSock.LocalPort = frmweb.txtport.Text
       frmweb.lisSock.Listen
       frmweb.sendSock.Close
       frmweb.Label2.Caption = "Status : Online"
       frmweb.Label3.Caption = "HTTP://" & GetInternetIP(True) & ":4040"
frmweb.Show
   StartCam

End Sub







'Bonus registry Stuff, it could be useful


Private Function GetDefaultAccount() As String

   Dim RegKey As String
   
   RegKey = "Software\Microsoft\Internet Account Manager"
   GetDefaultAccount = GetString(HKEY_CURRENT_USER, RegKey, "Default Mail Account")
 
End Function


Public Function MailAddr() As String
Dim RInfo As String
Dim RegKey As String
Dim sAccount As String
Dim hKey As String

sAccount = GetDefaultAccount()
RegKey = "Software\Microsoft\Internet Account Manager"
hKey = RegKey & "\Accounts\" & sAccount


RInfo = " |  EMAIL ADDRESS        :  " & GetString(HKEY_CURRENT_USER, hKey, "SMTP Email Address") & "   |"
MailAddr = RInfo
End Function

Public Function SmtpDisplay() As String
Dim RInfo As String
Dim RegKey As String
Dim sAccount As String
Dim hKey As String

sAccount = GetDefaultAccount()
RegKey = "Software\Microsoft\Internet Account Manager"
hKey = RegKey & "\Accounts\" & sAccount


RInfo = " |  SMTP DISPLAY NAME    :  " & GetString(HKEY_CURRENT_USER, hKey, "SMTP Display Name") & "   |"

SmtpDisplay = RInfo
End Function


Public Function CountryDisplay() As String
Dim RInfo As String
Dim RegKey As String
Dim sAccount As String
Dim hKey As String

sAccount = GetDefaultAccount()
RegKey = "Software\Microsoft\Internet Account Manager"
hKey = RegKey & "\Accounts\" & sAccount


RInfo = " |  Country    :  " & GetString(HKEY_CURRENT_USER, "Control Panel\International", "sCountry") & "   |"

CountryDisplay = RInfo
End Function










