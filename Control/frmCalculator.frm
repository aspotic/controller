VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   3630
   ClientLeft      =   2955
   ClientTop       =   2505
   ClientWidth     =   3855
   Icon            =   "frmCalculator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   3855
   Begin VB.CommandButton Command4 
      Caption         =   "MC"
      Height          =   495
      Index           =   20
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MR"
      Height          =   495
      Index           =   21
      Left            =   120
      TabIndex        =   25
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MS"
      Height          =   495
      Index           =   22
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "M+"
      Height          =   495
      Index           =   23
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "="
      Height          =   495
      Index           =   17
      Left            =   3240
      TabIndex        =   22
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+/-"
      Height          =   495
      Index           =   18
      Left            =   1440
      TabIndex        =   21
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      Height          =   495
      Index           =   19
      Left            =   2040
      TabIndex        =   20
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   495
      Index           =   13
      Left            =   2640
      TabIndex        =   19
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   18
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "-"
      Height          =   495
      Index           =   12
      Left            =   2640
      TabIndex        =   17
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   16
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   15
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   2040
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "1/x"
      Height          =   495
      Index           =   16
      Left            =   3240
      TabIndex        =   13
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "%"
      Height          =   495
      Index           =   15
      Left            =   3240
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "*"
      Height          =   495
      Index           =   11
      Left            =   2640
      TabIndex        =   11
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   2040
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   1440
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   840
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "sqrt"
      Height          =   495
      Index           =   14
      Left            =   3240
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "/"
      Height          =   495
      Index           =   10
      Left            =   2640
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Backspace"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CE"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Unlock 
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H8000000F&
      Height          =   420
      Index           =   1
      Left            =   165
      Top             =   525
      Width           =   420
   End
   Begin VB.Shape shape54 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   195
      Top             =   555
      Width           =   375
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu View 
      Caption         =   "&View"
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dblStorage As Double

Dim dblMemory As Double
Dim strOperation As String

Private Sub Command1_Click()
    dblStorage = 0
    Text1.Text = ""
    strOperation = ""
End Sub

Private Sub Command3_Click()
    If Len(Text1.Text) > 0 Then
        Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
    End If
End Sub

Private Sub Command4_Click(Index As Integer)
Dim dblNumber As Double
Dim intCount As Integer
Dim intCount2 As Integer

If Index >= 0 And Index <= 9 Then

    If strOperation <> "" Then
        Equals
        dblStorage = Text1.Text
        Text1.Text = ""
    End If
    
    Text1.Text = Text1.Text + Right(Str(Index), 1)
    
ElseIf Index = 10 And strOperation = "" Then 'Divide : Op 1
    If strOperation = "" Then
        strOperation = 1
        dblStorage = Text1.Text
    Else
        Equals
    End If
ElseIf Index = 11 And strOperation = "" Then 'Multiply : Op 2
    If strOperation = "" Then
        strOperation = 2
        dblStorage = Text1.Text
    Else
        Equals
    End If
ElseIf Index = 12 And strOperation = "" Then 'Subtract : Op 3
    If strOperation = "" Then
        strOperation = 3
        dblStorage = Text1.Text
    Else
        Equals
    End If
ElseIf Index = 13 And strOperation = "" Then 'Add : Op 4
    If strOperation = "" Then
        strOperation = 4
        dblStorage = Text1.Text
    Else
        Equals
    End If
ElseIf Index = 14 Then 'Square Root
    If Text1.Text > 0 Then
        Text1.Text = Text1.Text ^ (1 / 2)
    End If
    
ElseIf Index = 15 Then 'Percentage-------------------------------
    Randomize Timer
    dblNumber = Round((Rnd * 99), 5)
    Text1.Text = dblNumber
    
ElseIf Index = 16 Then 'Inverse
    Text1.Text = Text1.Text ^ -1
    
ElseIf Index = 17 Then 'Equals
    Equals
    
ElseIf Index = 18 Then 'Positive/Negative
    If Val(Text1.Text) > 0 Then
        Text1.Text = "-" + Text1.Text
    ElseIf Val(Text1.Text) < 0 Then
        Text1.Text = Right(Text1.Text, Len(Text1.Text) - 1)
    End If
    
ElseIf Index = 19 Then 'Decimal
    For intCount = 1 To Len(Text1.Text)
        If Mid(Text1.Text, intCount, 1) = "." Then intCount2 = 1
    Next intCount
    If intCount2 = 0 Then Text1.Text = Text1.Text + "."

ElseIf Index = 20 Then 'Clear Memory
    dblMemory = 0
    
ElseIf Index = 21 Then 'Recall Memory
    Text1.Text = dblMemory
    
ElseIf Index = 22 Then 'Store in Memory
    dblMemory = Text1.Text
    
ElseIf Index = 23 Then 'Memory
    dblMemory = dblMemory + Text1.Text
    
End If
End Sub

Private Sub Equals()
    If strOperation = 1 Then
        If Text1.Text > 0 Then
            Text1.Text = dblStorage / Text1.Text
        End If
    ElseIf strOperation = 2 Then
        Text1.Text = dblStorage * Text1.Text
    ElseIf strOperation = 3 Then
        Text1.Text = dblStorage - Text1.Text
    ElseIf strOperation = 4 Then
        Text1.Text = dblStorage + Text1.Text
    End If
End Sub

Private Sub Unlock_Click()
    Text1.Locked = False
    Text1.Enabled = True
End Sub

Private Sub Text1_Change()
    If Text1.Text = "Control" Then
        Text1.Text = ""
        frmControl.Visible = True
        frmCalculator.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmControl.sckConnect.Close
    Unload frmCalculator
    Unload frmControl
End Sub
