VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CMD Server"
   ClientHeight    =   420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2145
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   2145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox txtPort 
      Height          =   225
      Left            =   480
      TabIndex        =   1
      Text            =   "512"
      Top             =   60
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   660
      Top             =   660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   512
   End
   Begin VB.Label Label1 
      Caption         =   "Port:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetVersion Lib "kernel32" () As Long
Private blnOnline As Boolean

'Listen/close the server for connections
Private Sub cmdListen_Click()
If cmdListen.Caption = "Listen" Then
    cmdListen.Caption = "Close"
    Server.LocalPort = txtPort.Text
    StartServer
Else
    cmdListen.Caption = "Listen"
    ServerClose
End If
End Sub

Private Sub Form_Load()
Dim Port As Integer, prevHwnd As Long
If IsWin9x Then
    MsgBox "CMDServer can´t run on Win9X Systems !", vbInformation
    End
End If
AddSystray Me, "CMDServer"
End Sub

'Start listening, change current directory to C:\
Private Sub StartServer()
Server.Listen
CD "c:\"
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveSystray
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim rtn As Long
If Me.ScaleMode = vbPixels Then
    rtn = x
  Else
    rtn = x / Screen.TwipsPerPixelX
End If

Select Case rtn
  Case WM_LBUTTONUP
    SetForegroundWindow Me.hwnd
    Me.WindowState = 0
    Me.Show
End Select
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Me.Hide
End Sub

Private Sub Server_Close()
ServerClose
End Sub

'Close the server
Private Sub ServerClose()
blnOnline = False
blnPass = False
User = ""
Pass = ""
Server.Close
CD "C:\"
ModifySystray Me, "CMDServer"

End Sub

'Triggers when a client attempts to connect to the server
Private Sub Server_ConnectionRequest(ByVal requestID As Long)
If Not blnOnline Then
    blnOnline = True
    blnPass = False
    If Server.State <> sckClosed Then Server.Close
    Server.Accept requestID
    ModifySystray Me, Server.RemoteHostIP & " Is online"
    Server.SendData "Send UserName"
End If
End Sub

'Triggers when a client sends data to the server
Private Sub Server_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Server.GetData strData

If User = "" Then 'First data sent to the server must be Username
    User = Trim(strData)
    Server.SendData "Send Password"
ElseIf Pass = "" Then 'Second data sent to the server must be Pass
    Pass = Trim(strData)
    If UserValidate(User, Pass) Then 'Validate the user and pass with the local Security
        Server.SendData "Welcome, User and Pass is approved"
        DoEvents
        Server.SendData "CURDIRC:\>" 'tell the client current directory is C:\
        DoEvents
        blnPass = True
    Else 'If the user validation failed kick the remote client out
        Server.SendData "Access denied, User and Pass is not approved"
        DoEvents
        Call ServerClose
    End If
ElseIf blnPass Then 'if the user is validated and accepted check the data for commands
    Server.SendData GetCMD(strData)
End If
End Sub

'Check if the server is NT/Win2K/XP
Private Function IsWin9x() As Boolean
  IsWin9x = CBool(GetVersion() And &H80000000)
End Function

Private Sub txtPort_Change()
If Not IsNumeric(txtPort.Text) Then txtPort.Text = "512"
End Sub
