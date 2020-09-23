Attribute VB_Name = "modSystray"
Option Explicit

'This module declares the necessary WinAPIs and Constants necessary to create, modify
'and remove system tray icons.

'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize           As Long
    hwnd             As Long
    uId              As Long
    uFlags           As Long
    uCallBackMessage As Long
    hIcon            As Long
    szTip            As String * 64
End Type

'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0              'Used to add an icon to the systray
Public Const NIM_MODIFY = &H1           'Used to make any changes to an existing systray icon
Public Const NIM_DELETE = &H2           'used to remvoe an icon from the systray
Public Const NIF_MESSAGE = &H1          'The flag that we are specifying which event triggers systray activity
Public Const NIF_ICON = &H2             'The flag that we are including the icon image
Public Const NIF_TIP = &H4              'the flag that we are including a tooltip text over the systray icon
Public Const WM_MOUSEMOVE = &H200       'the event that we choose to process systray events.
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public nID As NOTIFYICONDATA            'The systray icon data

'WinAPI functions for creating and using a systray item....
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'loads the systray item, makes it easier to call from VB,
'see form_load of the frmMain Form...
Public Sub AddSystray(myForm As Form, myTip As String)
    With nID
        .cbSize = Len(nID)
        .hwnd = myForm.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = myForm.Icon
        .szTip = myTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nID
End Sub

'use this to modify the systray icon whether it be the
'icon or tool tip text
Public Sub ModifySystray(myForm As Form, myTip As String)
    With nID
        .cbSize = Len(nID)
        .hwnd = myForm.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = myForm.Icon
        .szTip = myTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_MODIFY, nID
End Sub

'use this to remove the systray icon
Public Sub RemoveSystray()
    Shell_NotifyIcon NIM_DELETE, nID
End Sub


