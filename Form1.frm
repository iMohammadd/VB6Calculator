VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command17 
      Caption         =   "Reset"
      Height          =   615
      Left            =   1680
      TabIndex        =   18
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2295
      Begin VB.CommandButton Command11 
         Caption         =   "."
         Height          =   615
         Left            =   1560
         TabIndex        =   17
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "0"
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Caption         =   "9"
         Height          =   615
         Left            =   1560
         TabIndex        =   15
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "8"
         Height          =   615
         Left            =   840
         TabIndex        =   14
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "7"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "6"
         Height          =   615
         Left            =   1560
         TabIndex        =   12
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "5"
         Height          =   615
         Left            =   840
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "3"
         Height          =   615
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2"
         Height          =   615
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton Command16 
      Caption         =   "="
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      Caption         =   "/"
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "*"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "-"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "+"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Dim b As Double
Dim c As Double
Dim action As String

Private Sub Command1_Click()
txtInput.Text = txtInput.Text + "1"
End Sub

Private Sub Command10_Click()
txtInput.Text = txtInput.Text + "0"
End Sub

Private Sub Command11_Click()
txtInput.Text = txtInput.Text + "."
End Sub

Private Sub Command12_Click()
a = Val(txtInput.Text)
action = "+"
txtInput.Text = ""
End Sub

Private Sub Command13_Click()
a = Val(txtInput.Text)
action = "-"
txtInput.Text = ""
End Sub

Private Sub Command14_Click()
a = Val(txtInput.Text)
action = "*"
txtInput.Text = ""
End Sub

Private Sub Command15_Click()
a = Val(txtInput.Text)
action = "/"
txtInput.Text = ""
End Sub

Private Sub Command16_Click()
b = Val(txtInput.Text)
Select Case (action)
Case "+"
c = a + b
Case "-"
c = a - b
Case "*"
c = a * b
Case "/"
c = a / b
End Select
txtInput.Text = c
action = ""
End Sub

Private Sub Command17_Click()
action = ""
txtInput = ""
End Sub

Private Sub Command2_Click()
txtInput.Text = txtInput.Text + "2"
End Sub

Private Sub Command3_Click()
txtInput.Text = txtInput.Text + "3"
End Sub

Private Sub Command4_Click()
txtInput.Text = txtInput.Text + "4"
End Sub

Private Sub Command5_Click()
txtInput.Text = txtInput.Text + "5"
End Sub

Private Sub Command6_Click()
txtInput.Text = txtInput.Text + "6"
End Sub

Private Sub Command7_Click()
txtInput.Text = txtInput.Text + "7"
End Sub

Private Sub Command8_Click()
txtInput.Text = txtInput.Text + "8"
End Sub

Private Sub Command9_Click()
txtInput.Text = txtInput.Text + "9"
End Sub
