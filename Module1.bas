Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function RegOpenKey Lib _
"advapi32" Alias "RegOpenKeyA" (ByVal hKey _
As Long, ByVal lpSubKey As String, _
phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
Lib "advapi32" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As _
String, lpReserved As Long, lptype As _
Long, lpData As Any, lpcbData As Long) _
As Long

Private Declare Function RegCloseKey& Lib _
"advapi32" (ByVal hKey&)

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const ERROR_SUCCESS = 0
Public Const HKEY_CLASSES_ROOT = &H80000000

Public Function GetRegString(hKey As Long, _
strSubKey As String, strValueName As _
String) As String
Dim strSetting As String
Dim lngDataLen As Long
Dim lngRes As Long
If RegOpenKey(hKey, strSubKey, _
lngRes) = ERROR_SUCCESS Then
   strSetting = Space(255)
   lngDataLen = Len(strSetting)
   If RegQueryValueEx(lngRes, _
   strValueName, ByVal 0, _
   REG_EXPAND_SZ, ByVal strSetting, _
   lngDataLen) = ERROR_SUCCESS Then
      If lngDataLen > 1 Then
      GetRegString = Left(strSetting, lngDataLen - 1)
   End If
End If

If RegCloseKey(lngRes) <> ERROR_SUCCESS Then
   MsgBox "RegCloseKey Failed: " & _
   strSubKey, vbCritical
End If
End If
End Function


