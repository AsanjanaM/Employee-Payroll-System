VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "welcome form"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15555
   LinkTopic       =   "Form2"
   ScaleHeight     =   8265
   ScaleWidth      =   15555
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "welcome"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   7560
      TabIndex        =   0
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   10950
      Left            =   0
      Picture         =   "welform.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20100
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show
Me.Hide

End Sub

