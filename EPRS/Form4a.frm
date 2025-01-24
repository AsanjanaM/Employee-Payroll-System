VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   Caption         =   "employee details"
   ClientHeight    =   10110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19050
   LinkTopic       =   "Form4"
   ScaleHeight     =   10110
   ScaleWidth      =   19050
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text12 
      DataField       =   "dob"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      TabIndex        =   37
      Text            =   "Text12"
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      DataField       =   "empid"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   3480
      TabIndex        =   35
      Text            =   "Text13"
      Top             =   360
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7680
      Top             =   7800
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\EPRS\EMPRS.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\EPRS\EMPRS.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "empdet"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FF80FF&
      DataField       =   "gender"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      ItemData        =   "Form4a.frx":0000
      Left            =   3360
      List            =   "Form4a.frx":000A
      TabIndex        =   33
      Text            =   "Select"
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "next"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   17040
      TabIndex        =   32
      Top             =   8880
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FFC0FF&
      DataField       =   "doj"
      DataSource      =   "Adodc1"
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
      Left            =   3240
      TabIndex        =   30
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "home"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   14400
      TabIndex        =   29
      Top             =   9720
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "previous"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   28
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "move last"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16920
      TabIndex        =   27
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "move first"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   26
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "delete"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16800
      TabIndex        =   25
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "clear"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   24
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "update"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16680
      TabIndex        =   23
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "add new"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   22
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFC0FF&
      DataField       =   "basicpay"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   21
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFC0FF&
      DataField       =   "pfacc"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   20
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFC0FF&
      DataField       =   "panno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   19
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFC0FF&
      DataField       =   "addhrno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   18
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFC0FF&
      DataField       =   "email"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   17
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFC0FF&
      DataField       =   "empid"
      DataSource      =   "Adodc1"
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
      Left            =   3240
      TabIndex        =   16
      Top             =   9720
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFC0FF&
      DataField       =   "mobno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   8520
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFC0FF&
      DataField       =   "address"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   3240
      TabIndex        =   14
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0FF&
      DataField       =   "lastn"
      DataSource      =   "Adodc1"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0FF&
      DataField       =   "firstn"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "Gender"
      Height          =   375
      Left            =   720
      TabIndex        =   36
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "empid"
      Height          =   615
      Left            =   240
      TabIndex        =   34
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF00FF&
      Caption         =   "employee details"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   6120
      TabIndex        =   31
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FF80FF&
      Caption         =   "basic pay"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FF80FF&
      Caption         =   "pf account"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FF80FF&
      Caption         =   "addhar no"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FF80FF&
      Caption         =   "pan no"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF80FF&
      Caption         =   "emp id"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF80FF&
      Caption         =   "email"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF80FF&
      Caption         =   "mobile no"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF80FF&
      Caption         =   "address"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF80FF&
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF80FF&
      Caption         =   "date of joining"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF80FF&
      Caption         =   "last name"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF80FF&
      Caption         =   "first name"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   11760
      Left            =   0
      Picture         =   "Form4a.frx":001C
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   20850
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Update
MsgBox "RECORD SAVED"
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
MsgBox "RECORD DELETE"
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command8_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command9_Click()
Form3.Show
End Sub

