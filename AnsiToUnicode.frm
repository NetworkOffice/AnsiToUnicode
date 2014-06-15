VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "AnsiToUnicode"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unicode >>> ANSI"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ANSI >>> Unicode"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Unicode:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ANSI:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2.Text = StrConv(Text1.Text, vbFromUnicode)
End Sub

Private Sub Command2_Click()
Text1.Text = StrConv(Text2.Text, vbUnicode)
End Sub
