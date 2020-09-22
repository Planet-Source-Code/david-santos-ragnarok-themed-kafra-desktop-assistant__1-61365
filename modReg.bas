Attribute VB_Name = "modReg"
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long

Public Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'REGISTRY_KEY
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

Public Const KEY_ALL_ACCESS = &HF003F
Public Const KEY_CREATE_LINK = &H20&
Public Const KEY_CREATE_SUB_KEY = &H4&
Public Const KEY_ENUMERATE_SUB_KEYS = &H8&
Public Const KEY_EXECUTE = &H20019
Public Const KEY_NOTIFY = &H10&
Public Const KEY_QUERY_VALUE = &H1&
Public Const KEY_READ = &H20019
Public Const KEY_SET_VALUE = &H2&
Public Const KEY_WRITE = &H20006

Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_DWORD_BIG_ENDIAN As Long = 5
Public Const REG_DWORD_LITTLE_ENDIAN As Long = 4
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_NONE As Long = 0
Public Const REG_RESOURCE_LIST As Long = 8
Public Const REG_SZ As Long = 1
'Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
'Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Public Function SetStringKey(startkey As Long, subkey As String, key As String, Value As String)
Dim hkey As Long            ' receives handle to the registry key
Dim secattr As SECURITY_ATTRIBUTES  ' security settings for the key
Dim neworused As Long       ' receives flag for if the key was created or opened
Dim stringbuffer As String  ' the string to put into the registry
Dim retval As Long          ' return value
    
    Value = Value & vbNullChar
    
    retval = RegCreateKeyEx(startkey, subkey, 0, "", 0, KEY_WRITE, secattr, hkey, neworused)
    retval = RegSetValueEx(hkey, key, 0, REG_SZ, ByVal Value, Len(Value))
    RegCloseKey hkey

End Function

Public Function GetStringKey(startkey As Long, subkey As String, key As String) As String
On Error GoTo errhandler
Dim hkey As Long
Dim secattr As SECURITY_ATTRIBUTES
Dim retval As Long
Dim stringbuffer As String
Dim slength As Long
Dim ret As Long
Dim dtype As Long
Dim neworused As Long

    stringbuffer = Space(255)
    slength = 255
    
    retval = RegCreateKeyEx(startkey, subkey, 0, "", 0, KEY_QUERY_VALUE, secattr, hkey, neworused)
    retval = RegQueryValueEx(hkey, key, 0, dtype, ByVal stringbuffer, slength)

    If retval <> 0 Then Err.Raise vbObjectError + retval
    
    GetStringKey = Left(stringbuffer, slength - 1)
    RegCloseKey hkey
    
    Exit Function

errhandler:
    GetStringKey = ""
End Function

Public Function SetHexKey(startkey As Long, subkey As String, key As String, lBuffer As Long)
Dim hkey As Long
Dim secattr As SECURITY_ATTRIBUTES  ' security settings for the key
Dim retval As Long
Dim neworused As Long       ' receives flag for if the key was created or opened
    
    ret = RegCreateKeyEx(startkey, subkey, 0, "", 0, KEY_WRITE, secattr, hkey, neworused)
    retval = RegSetValueEx(hkey, key, 0, REG_DWORD, lBuffer, Len(lBuffer))
End Function

Public Function GetHexKey(startkey As Long, subkey As String, key As String) As Long
On Error GoTo errhandler
Dim hkey As Long
Dim secattr As SECURITY_ATTRIBUTES  ' security settings for the key
Dim lBuffer As Long
Dim ret As Long
Dim neworused As Long       ' receives flag for if the key was created or opened
Dim dtype As Long

    ret = RegCreateKeyEx(startkey, subkey, 0, "", 0, KEY_READ, secattr, hkey, neworused)
    ret = RegQueryValueEx(hkey, key, 0, dtype, lBuffer, Len(lBuffer))
        
    If ret <> 0 Then Err.Raise vbObjectError + ret
    GetHexKey = lBuffer
    Exit Function

errhandler:
    GetHexKey = 0&
End Function

Public Function RemoveKey(startkey As Long, subkey As String, key As String)
Dim hkey As Long            ' receives handle to the registry key
Dim retval As Long          ' return value
    
    retval = RegOpenKeyEx(startkey, subkey, 0, KEY_ALL_ACCESS, hkey)
    If retval = 0 Then
      retval = RegDeleteValue(hkey, key)
      retval = RegCloseKey(hkey)
    End If
End Function

Public Sub SetAutoRun(mode As Boolean)
    If mode = True Then
        SetStringKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Desktop Kafra", App.path & "\" & App.EXEName & ".exe"
    Else
        RemoveKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", "Desktop Kafra"
    End If
End Sub

