Attribute VB_Name = "modUserValidate"
Option Explicit

Private Declare Function NetUserChangePassword Lib "Netapi32.dll" (ByVal sDomain As String, ByVal sUserName As String, ByVal sOldPassword As String, ByVal sNewPassword As String) As Long
Private Declare Function NetUserGetLocalGroups _
  Lib "Netapi32.dll" (lpServer As Any, UserName As Byte, _
   ByVal Level As Long, ByVal Flags As Long, lpBuffer As Long, _
   ByVal MaxLen As Long, lpEntriesRead As Long, _
   lpTotalEntries As Long) As Long
  
  
Private Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" (Destination As Any, Source As Any, _
    ByVal Length As Long)

Private Declare Function lstrlenW Lib "kernel32" _
   (ByVal lpString As Long) As Long

Private Declare Function NetApiBufferFree Lib "netapi32" _
  (ByVal pBuffer As Long) As Long



'Checks if username with password exists
Function UserValidate(sUserName As String, sPassword As String, Optional sDomain As String) As Boolean
    Dim lReturn As Long
    Const NERR_BASE = 2100
    Const NERR_PasswordCantChange = NERR_BASE + 143
    Const NERR_PasswordHistConflict = NERR_BASE + 144
    Const NERR_PasswordTooShort = NERR_BASE + 145
    Const NERR_PasswordTooRecent = NERR_BASE + 146
    
    If Len(sDomain) = 0 Then
        sDomain = Environ$("USERDOMAIN")
    End If
    
    'Call API to check password.
    lReturn = NetUserChangePassword(StrConv(sDomain, vbUnicode), StrConv(sUserName, vbUnicode), StrConv(sPassword, vbUnicode), StrConv(sPassword, vbUnicode))
    
    'Test return value.
    Select Case lReturn
    Case 0, NERR_PasswordCantChange, NERR_PasswordHistConflict, NERR_PasswordTooShort, NERR_PasswordTooRecent
        UserValidate = True
    Case Else
        UserValidate = False
    End Select
End Function

'Checks if username provided is a local administrator
Public Function IsUserLocalAdmin(ByVal UserName As String) As Boolean
Dim bytUser() As Byte
Dim bytServer() As Byte

Dim lBuffer As Long
Dim lMaxLen As Long
Dim lTotalEntries As Long
Dim lRet As Long
Dim lGroups() As Long
Dim sGroups() As String
Dim bytBuffer() As Byte
Dim iCtr As Integer
Dim lLen As Long
Dim ServerName As String
Dim blnRet As Boolean
bytServer = vbNullChar
bytUser = UserName & vbNullChar

 lRet = NetUserGetLocalGroups(bytServer(0), bytUser(0), 0, 0, _
   lBuffer, 1024, lMaxLen, lTotalEntries)

    
If lRet = 0 And lMaxLen > 0 Then
      ReDim lGroups(lMaxLen - 1) As Long
      ReDim sGroups(lMaxLen - 1) As String
      CopyMemory lGroups(0), ByVal lBuffer, lMaxLen * 4
      For iCtr = 0 To lMaxLen - 1
          lLen = lstrlenW(lGroups(iCtr)) * 2
           If lLen > 0 Then
                ReDim bytBuffer(lLen - 1) As Byte
                CopyMemory bytBuffer(0), ByVal lGroups(iCtr), lLen
                sGroups(iCtr) = bytBuffer
                If sGroups(iCtr) = "Administrators" Then
                    blnRet = True
                    Exit For
                End If
                
            End If
      Next
End If
 
If lBuffer > 0 Then NetApiBufferFree (lBuffer)

IsUserLocalAdmin = blnRet
End Function



