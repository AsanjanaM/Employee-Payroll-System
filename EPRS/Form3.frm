VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "main form"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16785
   LinkTopic       =   "Form3"
   ScaleHeight     =   9630
   ScaleWidth      =   16785
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command13 
      Caption         =   "INCOME TAX REPORT"
      Height          =   375
      Left            =   8400
      TabIndex        =   22
      Top             =   9240
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      Caption         =   "PAY SLIP REPORT"
      Height          =   375
      Left            =   4800
      TabIndex        =   21
      Top             =   9240
      Width           =   2535
   End
   Begin VB.CommandButton Command11 
      Caption         =   "EMPLOYEE REPORT"
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   9240
      Width           =   2655
   End
   Begin VB.CommandButton Command10 
      Caption         =   "monthly attendence"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   19
      Top             =   7800
      Width           =   2535
   End
   Begin VB.PictureBox Picture10 
      Height          =   2535
      Left            =   240
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   3555
      TabIndex        =   18
      Top             =   4920
      Width           =   3615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "personal details"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   7920
      Width           =   1935
   End
   Begin VB.PictureBox Picture9 
      Height          =   2895
      Left            =   5280
      Picture         =   "Form3.frx":198B
      ScaleHeight     =   2835
      ScaleWidth      =   2835
      TabIndex        =   16
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton Command8 
      Caption         =   "employee view"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16560
      TabIndex        =   15
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "department view"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   14
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "personal view"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   13
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "salary calculation"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16440
      TabIndex        =   12
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   " income tax calculation"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      TabIndex        =   11
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "training"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "department details"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   2595
      Left            =   360
      Picture         =   "Form3.frx":907D
      ScaleHeight     =   2535
      ScaleWidth      =   2595
      TabIndex        =   8
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "employee details"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.PictureBox Picture8 
      Height          =   2655
      Left            =   16080
      Picture         =   "Form3.frx":10851
      ScaleHeight     =   2595
      ScaleWidth      =   2595
      TabIndex        =   6
      Top             =   4920
      Width           =   2655
   End
   Begin VB.PictureBox Picture7 
      Height          =   2655
      Left            =   12600
      Picture         =   "Form3.frx":1818B
      ScaleHeight     =   2595
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   4920
      Width           =   2775
   End
   Begin VB.PictureBox Picture6 
      Height          =   2655
      Left            =   8880
      Picture         =   "Form3.frx":1FC4E
      ScaleHeight     =   2595
      ScaleWidth      =   2955
      TabIndex        =   4
      Top             =   4920
      Width           =   3015
   End
   Begin VB.PictureBox Picture5 
      Height          =   2535
      Left            =   15960
      Picture         =   "Form3.frx":284A5
      ScaleHeight     =   2475
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox Picture4 
      Height          =   2535
      Left            =   12120
      Picture         =   "Form3.frx":29AAC
      ScaleHeight     =   2475
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
   Begin VB.PictureBox Picture3 
      Height          =   2535
      Left            =   8160
      Picture         =   "Form3.frx":2B04F
      ScaleHeight     =   2475
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   3960
      Picture         =   "Form3.frx":3162F
      ScaleHeight     =   2475
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   27000
      Left            =   -1680
      Picture         =   "Form3.frx":330A9
      Top             =   -120
      Width           =   43200
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Show
End Sub

Private Sub Command10_Click()
Form8.Show
End Sub

Private Sub Command11_Click()
DataReport3.Show
End Sub

Private Sub Command12_Click()
DataReport2.Show
End Sub

Private Sub Command13_Click()
DataReport4.Show
End Sub

Private Sub Command2_Click()
Form6.Show
End Sub

Private Sub Command3_Click()
Form7.Show
End Sub

Private Sub Command5_Click()
Form9.Show
End Sub

Private Sub Command6_Click()
Form11.Show
End Sub

Private Sub Command8_Click()
Form10.Show
End Sub

Private Sub Command9_Click()
Form5.Show
End Sub

