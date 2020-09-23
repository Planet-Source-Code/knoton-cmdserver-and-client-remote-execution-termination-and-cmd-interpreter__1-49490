Attribute VB_Name = "modRexec"
Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As String, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Const SMTO_BLOCK = &H1
Public Const SMTO_ABORTIFHUNG = &H2
Public Const WM_CLOSE = &H10
Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Public RunningApps As String

'Get Full Paths shortpath name
Private Function GetShortPath(strFileName As String) As String
Dim Ret As Long
Dim tmp As String
Dim ShortFileName As String * 260
Ret = GetShortPathName(strFileName, ShortFileName, Len(ShortFileName))
GetShortPath = Left$(ShortFileName, Ret)
End Function

'Open/Execute a file
Public Function RExec(ByVal strPath As String) As Long
Dim blnError As Boolean
On Error GoTo errhandler
Dim Hinst As Long
Hinst = Shell(GetShortPath(strPath), 1)
RExec = Hinst
errhandler:
If Not blnError Then
    strPath = CurDir & strPath
    blnError = True
    Resume
End If
End Function

'Terminate a running process
Public Sub RTerminate(ByVal RHwnd As Long)
Dim hwnd As Long
Dim Ret As Long
Dim lngResult As Long
Dim lngProcessID As Long
Dim lngProcess As Long

hwnd = RHwnd
'Try a clean shutdown
Ret = SendMessageTimeout(hwnd, WM_CLOSE, 0&, 0&, SMTO_BLOCK, 5000, lngResult)

'Incase the clean shutdown didnt go well do for safe sake a kill
Ret = GetWindowThreadProcessId(hwnd, lngProcessID)
lngProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, lngProcessID)
Ret = TerminateProcess(lngProcess, 0&)

End Sub

'Returns a string with all running apps eg topmost visible windows
Public Function GetRunningApps() As String
Dim Ret As Long
RunningApps = ""
Ret = EnumWindows(AddressOf GetRunningWindows, 0)
GetRunningApps = RunningApps

End Function

'Get running apps/ topmost visible windows
Private Function GetRunningWindows(ByVal hwnd As Long, ByVal lParam As Long) As Long

Dim ForeGroundWindow As Long
Dim TextLen As Long
Dim WindowText As String
Dim Ret As Long
Static LastWindowText As String

ForeGroundWindow = hwnd
TextLen = GetWindowTextLength(ForeGroundWindow) + 1

WindowText = Space(TextLen)
Ret = GetWindowText(ForeGroundWindow, WindowText, TextLen)
WindowText = Left(WindowText, Len(WindowText) - 1)

If WindowText = "" Then GoTo ExitSub

If IsWindowVisible(ForeGroundWindow) > 0 Then
    If WindowText = Form1.Caption Then GoTo ExitSub
    RunningApps = RunningApps & "HWND = " & ForeGroundWindow & vbTab & WindowText & vbCrLf
    LastWindowText = WindowText
End If

ExitSub:
GetRunningWindows = 1
End Function

