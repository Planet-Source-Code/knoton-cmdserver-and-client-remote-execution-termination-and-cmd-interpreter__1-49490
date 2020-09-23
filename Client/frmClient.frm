VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H80000008&
   Caption         =   "Remote CMD     (type cmdhelp and press Return to get help)"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10635
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Client 
      Left            =   4920
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   512
   End
   Begin VB.TextBox txtCMD 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   5700
      Width           =   10215
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      ForeColor       =   &H00FFFF00&
      Height          =   5655
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   10575
   End
   Begin VB.Label lblCurDir 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "C:\>"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   5700
      Width           =   315
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnConnected As Boolean

'Triggers when the client close the connection or is being closed
Private Sub Client_Close()
blnConnected = False
frmClient.Caption = "Connection is closed."
txtCMD.Text = ""
txtCMD.SetFocus
End Sub

'Triggers when the client attempts a connection
Private Sub Client_Connect()
blnConnected = True
frmClient.Caption = "You are connected to " & Client.RemoteHostIP & ":" & Client.RemotePort
txtCMD.Text = ""
txtCMD.SetFocus
End Sub

'Triggers when data arrives
Private Sub Client_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Client.GetData strData
If Mid(strData, 1, 6) = "CURDIR" Then
    lblCurDir.Caption = Replace(strData, "CURDIR", "")
    txtCMD.Left = lblCurDir.Width
    txtCMD.Width = frmClient.Width - lblCurDir.Width
    strData = lblCurDir.Caption
End If
txtOutput.Text = strData
txtOutput.SelStart = Len(txtOutput.Text)
txtCMD.Text = ""
txtCMD.SetFocus
End Sub

'Triggers when error occurs on the connection
Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtOutput.Text = Description
txtOutput.SelStart = Len(txtOutput.Text)
txtCMD.Text = ""
txtCMD.SetFocus
blnConnected = False
End Sub


Private Sub Form_Resize()
If Me.WindowState <> 1 Then
    txtOutput.Width = frmClient.Width - 170
    txtOutput.Height = frmClient.ScaleHeight - txtCMD.Height
    txtCMD.Top = txtOutput.Height
    lblCurDir.Top = txtCMD.Top
    txtCMD.Left = lblCurDir.Width
    txtCMD.Width = frmClient.Width - lblCurDir.Width
End If
End Sub

Private Sub txtCMD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' Enter was pressed
    If LCase(Mid(txtCMD.Text, 1, 7)) = "connect" Then
        Connect txtCMD.Text
    ElseIf LCase(Trim(txtCMD.Text)) = "exit" Then
        Client.Close
        blnConnected = False
        frmClient.Caption = "Connection is closed."
        txtCMD.Text = ""
        txtCMD.SetFocus
    ElseIf LCase(Trim(txtCMD.Text)) = "cls" Then
        txtOutput.Text = ""
        txtCMD.Text = ""
        txtCMD.SetFocus
    ElseIf LCase(Trim(txtCMD.Text)) = "cmdhelp" Then
        txtOutput.Text = "CONNECT IP/Host[:Port] (Connect to CMDServer on Port, Default = 512)" & vbCrLf & _
                         "EXIT (Close connection)" & vbCrLf & _
                         "REXIT (Close down CMDServer )" & vbCrLf & _
                         "CMDHELP (Show this help)" & vbCrLf & _
                         "GETRUNAPP (Show running apps)" & vbCrLf & _
                         "REXEC [FullPath/]Filename.ext [parameters] (Execute remote app)" & vbCrLf & _
                         "RTERM HWND (Close running app)" & vbCrLf & _
                         "CLS (Clear Screen)" & vbCrLf & vbCrLf & _
                         "OBSERVE that if you remotly start an application via the command interpreter the server will not return until the application is terminated." & vbCrLf & _
                         "Use REXEC instead" & vbCrLf & vbCrLf & _
                         "When you connect to the server you will be prompt for username and password" & vbCrLf & _
                         "note that the user most have administrator rights on the remote server."
        
        txtCMD.Text = ""
        txtCMD.SetFocus
    Else
        If Client.State = 7 Then
            Client.SendData txtCMD.Text
            txtCMD.Text = "Processing command..."
        Else
            Client.Close
            blnConnected = False
            frmClient.Caption = "Connection is closed."
            txtCMD.Text = ""
            txtCMD.SetFocus
        End If
    End If
End If
End Sub

'Initate the connection
Private Sub Connect(strConnect As String)
Dim varSplit As Variant, IP As String, Port As String
On Error GoTo Errhandler
If Not blnConnected Then
    varSplit = Split(Trim(strConnect))
    If UBound(varSplit) > 0 Then
        If InStr(1, varSplit(1), ":") Then
            IP = Trim(Mid(varSplit(1), 1, InStr(1, varSplit(1), ":") - 1))
            Port = Trim(Mid(varSplit(1), InStr(1, varSplit(1), ":") + 1))
        Else
            IP = Trim(varSplit(1))
            Port = Trim("512")
        End If
    End If
    
    Client.RemoteHost = IP
    Client.RemotePort = Port
    
    txtOutput.Text = "Connecting to " & IP & ":" & Port
    txtOutput.SelStart = Len(txtOutput.Text)
    txtCMD.Text = ""
    txtCMD.SetFocus

    Client.Connect
End If
Exit Sub
Errhandler:
txtOutput.Text = Err.Source & vbCrLf & Err.Description
txtOutput.SelStart = Len(txtOutput.Text)
txtCMD.Text = ""
txtCMD.SetFocus

End Sub
