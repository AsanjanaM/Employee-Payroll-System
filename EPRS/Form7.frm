VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   Caption         =   "training details"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11955
   LinkTopic       =   "Form7"
   ScaleHeight     =   6120
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3120
      TabIndex        =   28
      Text            =   "Combo2"
      Top             =   1200
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   13080
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      RecordSource    =   "tradet"
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
   Begin VB.TextBox Text9 
      BackColor       =   &H0080FFFF&
      DataField       =   "date"
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
      Left            =   3120
      TabIndex        =   27
      Top             =   7560
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0080FFFF&
      DataField       =   "mode"
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
      Height          =   420
      ItemData        =   "Form7.frx":0000
      Left            =   840
      List            =   "Form7.frx":000A
      TabIndex        =   25
      Text            =   "mode of training"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      Caption         =   "home"
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
      Left            =   7680
      TabIndex        =   24
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "next"
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
      Left            =   9600
      TabIndex        =   23
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "pervious"
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
      Left            =   7680
      TabIndex        =   22
      Top             =   5880
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
      Left            =   9600
      TabIndex        =   21
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "move first"
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
      Left            =   7680
      TabIndex        =   20
      Top             =   4680
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
      Left            =   9600
      TabIndex        =   19
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "clear"
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
      Left            =   7680
      TabIndex        =   18
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "update"
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
      Left            =   9600
      TabIndex        =   17
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
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
      Height          =   510
      Left            =   7680
      TabIndex        =   16
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H0080FFFF&
      DataField       =   "traname"
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
      Left            =   3120
      TabIndex        =   15
      Top             =   9480
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H0080FFFF&
      DataField       =   "location"
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
      Left            =   3120
      TabIndex        =   14
      Top             =   8520
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0080FFFF&
      DataField       =   "time"
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
      Left            =   3120
      TabIndex        =   11
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0080FFFF&
      DataField       =   "status"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   5880
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0080FFFF&
      DataField       =   "comdate"
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
      Left            =   3120
      TabIndex        =   7
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080FFFF&
      DataField       =   "trahours"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080FFFF&
      DataField       =   "empname"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FFFF&
      Caption         =   "date"
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
      Left            =   360
      TabIndex        =   26
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FFFF&
      Caption         =   "trainer name"
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
      Left            =   240
      TabIndex        =   13
      Top             =   9360
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FFFF&
      Caption         =   "location"
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
      Left            =   360
      TabIndex        =   12
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      Caption         =   "time"
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
      Left            =   360
      TabIndex        =   10
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FFFF&
      Caption         =   "status"
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
      Left            =   360
      TabIndex        =   8
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "completion date"
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
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Caption         =   "training hours"
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
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "employee name"
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
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "employee id"
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
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "training details"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   11175
      Left            =   0
      Picture         =   "Form7.frx":001F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New adodb.Recordset
Dim rs1 As New adodb.Recordset



 
Private Sub Combo2_Click()
  rs.Open "select * from empdet where empid =" & Val(Combo2.Text) & " ", con, adOpenStatic, adLockReadOnly
        If rs.RecordCount > 0 Then
           Text2.Text = rs("firstn")
            

             End If
            If rs.State = adStateOpen Then rs.Close
  End Sub
Private Sub Command1_Click()

Adodc1.Recordset.AddNew
Set rs = con.Execute("select empid from empdet")
While (Not rs.EOF)
    Combo2.AddItem rs(0)
     rs.MoveNext
     
Wend
rs.Close


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
MsgBox "RECORD DELETED"
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

Private Sub Form_Load()
connectdb
End Sub

