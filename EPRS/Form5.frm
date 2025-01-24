VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   Caption         =   "department details"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   LinkTopic       =   "Form5"
   ScaleHeight     =   8310
   ScaleWidth      =   14355
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3360
      TabIndex        =   33
      Text            =   "Text10"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Single"
      Height          =   375
      Left            =   4080
      TabIndex        =   32
      Top             =   4200
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Married"
      Height          =   375
      Left            =   2280
      TabIndex        =   31
      Top             =   4200
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "empid"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   3600
      TabIndex        =   29
      Text            =   "Combo2"
      Top             =   360
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   7560
      Top             =   5640
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
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
      RecordSource    =   "perdet"
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
      Left            =   14040
      TabIndex        =   27
      Top             =   9600
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
      Left            =   16080
      TabIndex        =   26
      Top             =   8760
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
      Left            =   14040
      TabIndex        =   25
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
      Left            =   16200
      TabIndex        =   24
      Top             =   7440
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
      Left            =   14160
      TabIndex        =   23
      Top             =   6720
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
      Left            =   16080
      TabIndex        =   22
      Top             =   5880
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
      Left            =   14040
      TabIndex        =   21
      Top             =   5160
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
      Left            =   16080
      TabIndex        =   20
      Top             =   4440
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
      Height          =   495
      Left            =   14040
      TabIndex        =   19
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00008000&
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
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   3720
      TabIndex        =   15
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00008000&
      DataField       =   "email"
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
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00008000&
      DataField       =   "dob"
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
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00008000&
      DataField       =   "fathern"
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
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00008000&
      DataField       =   "name"
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
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Frame frame2 
      BackColor       =   &H00008000&
      Caption         =   "bank details"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   3135
      Left            =   6480
      TabIndex        =   7
      Top             =   1800
      Width           =   5535
      Begin VB.TextBox Text9 
         BackColor       =   &H0080FF80&
         DataField       =   "ifsc"
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
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   3000
         TabIndex        =   18
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H0080FF80&
         DataField       =   "accountno"
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
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   3000
         TabIndex        =   17
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H0080FF80&
         DataField       =   "nameonacc"
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
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   3000
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080FF80&
         Caption         =   "isfc code"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080FF80&
         Caption         =   "account no"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "name on the account"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Label Label12 
      Caption         =   "gender"
      Height          =   375
      Left            =   600
      TabIndex        =   30
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Marital Status"
      Height          =   615
      Left            =   360
      TabIndex        =   28
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00008000&
      Caption         =   "previous company worked"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00008000&
      Caption         =   "Educational Qualtification"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      Caption         =   "Dob"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      Caption         =   "emp id"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      Caption         =   "father name"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Caption         =   "name"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "personal details"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   6240
      TabIndex        =   0
      Top             =   360
      Width           =   4485
   End
   Begin VB.Image Image1 
      Height          =   11655
      Left            =   0
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   -720
      Width           =   20730
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New adodb.Recordset
Dim rs1 As New adodb.Recordset



 
Private Sub Combo2_Click()
  rs.Open "select * from empdet where empid =" & Val(Combo2.Text) & " ", con, adOpenStatic, adLockReadOnly
        If rs.RecordCount > 0 Then
           Text1.Text = rs("firstn")
            Text4.Text = rs("dob")
            Text10.Text = rs("gender")

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
adoc1.Recordset.MoveNext
End Sub

Private Sub Command9_Click()
Form3.Show
End Sub

Private Sub Form_Load()
connectdb

End Sub

