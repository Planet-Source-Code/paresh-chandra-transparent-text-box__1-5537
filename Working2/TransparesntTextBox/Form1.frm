VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin Project1.TransTextBox TransTextBox3 
      Height          =   4215
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7435
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":0000
   End
   Begin Project1.TransTextBox TransTextBox1 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":E2F2
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change"
      Height          =   5175
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton Option3 
         Caption         =   "Center"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   4800
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Right"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   4560
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Left"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   4320
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "assign"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   3135
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Text            =   "Form1.frx":1C5E4
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "User name"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
TransTextBox3.Text = Text1
End Sub

Private Sub Form_DblClick()
MsgBox txt1.Text
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Label1.Caption = Label1.Caption & Chr$(KeyAscii)
End Sub


Private Sub Option1_Click()
TransTextBox3.Alignment = Left_Justify
End Sub


Private Sub Option2_Click()
TransTextBox3.Alignment = Right_Justify
End Sub


Private Sub Option3_Click()
TransTextBox3.Alignment = Center
End Sub


