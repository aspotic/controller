VERSION 5.00
Begin VB.Form frmPopup 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4080
      Top             =   1680
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Control Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblHappening 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrClose_Timer()
    Unload Me
End Sub
