VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   Caption         =   "salary calculation"
   ClientHeight    =   9165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18555
   LinkTopic       =   "Form9"
   ScaleHeight     =   9165
   ScaleWidth      =   18555
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      Caption         =   "CALCULATE"
      Height          =   495
      Left            =   7320
      TabIndex        =   35
      Top             =   3360
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   8280
      Top             =   4440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
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
      RecordSource    =   "salcal"
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
      BackColor       =   &H0080FF80&
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
      Height          =   420
      ItemData        =   "Form9.frx":0000
      Left            =   3240
      List            =   "Form9.frx":0010
      TabIndex        =   34
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00C0FFC0&
      DataField       =   "netsal"
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
      TabIndex        =   25
      Top             =   9360
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
      Left            =   12840
      TabIndex        =   22
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "clear"
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
      Left            =   12720
      TabIndex        =   21
      Top             =   3720
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
      Height          =   510
      Left            =   14640
      TabIndex        =   20
      Top             =   3000
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
      Left            =   12720
      TabIndex        =   19
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Caption         =   "deductions"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   840
      TabIndex        =   12
      Top             =   6000
      Width           =   5655
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0FFC0&
         DataField       =   "netded"
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
         Left            =   2520
         TabIndex        =   24
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00C0FFC0&
         DataField       =   "lic"
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
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   18
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFC0&
         DataField       =   "egis"
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
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   4320
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0FFC0&
         DataField       =   "pt"
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
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFC0&
         Caption         =   "net deductions"
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
         TabIndex        =   32
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "LIC"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "EGIS"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ta"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "allowances"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   960
      TabIndex        =   3
      Top             =   2880
      Width           =   5655
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFC0&
         DataField       =   "netall"
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
         Left            =   2520
         TabIndex        =   23
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FFFFFF&
         DataField       =   "others"
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
         ForeColor       =   &H00FF00FF&
         Height          =   420
         Left            =   4440
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFC0&
         DataField       =   "hra"
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
         Left            =   4440
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFC0&
         DataField       =   "ta"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         DataField       =   "da"
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
         ForeColor       =   &H00FF00FF&
         Height          =   420
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "net allowance"
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
         TabIndex        =   31
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "OTHERS"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HRA"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "TA"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "DA"
         BeginProperty Font 
            Name            =   "Segoe Script"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0080FF80&
      DataField       =   "basic"
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
      Left            =   8880
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080FF80&
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
      Height          =   400
      Left            =   8880
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080FF80&
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
      Height          =   420
      Left            =   3240
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackColor       =   &H0080FF80&
      Caption         =   "net salary"
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
      Left            =   960
      TabIndex        =   33
      Top             =   9360
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "basic pay"
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
      Left            =   6480
      TabIndex        =   30
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "PF account no"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6480
      TabIndex        =   29
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "emp name"
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
      Left            =   1080
      TabIndex        =   28
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
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
      Left            =   1080
      TabIndex        =   27
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "salary calculation"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   26
      Top             =   120
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   24000
      Left            =   -5880
      Picture         =   "Form9.frx":0020
      Stretch         =   -1  'True
      Top             =   -2640
      Width           =   27000
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New adodb.Recordset
Dim rs1 As New adodb.Recordset



 
Private Sub Combo1_Click()
  rs.Open "select * from empdet where empid =" & Val(Combo1.Text) & " ", con, adOpenStatic, adLockReadOnly
        If rs.RecordCount > 0 Then
           Text2.Text = rs("firstn")
            Text3.Text = rs("pfacc")
            Text4.Text = rs("basicpay")

             End If
            If rs.State = adStateOpen Then rs.Close
  End Sub




Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command10_Click()
Text1.Text = Val(Text4) * 45 / 100
Text6.Text = Val(Text4) * 10 / 100
Text2.Text = Val(Text4) * 2 / 100
Text8.Text = Val(Text1) + Val(Text6) + Val(Text2) + Val(Text7)
Text9.Text = Val(Text4) * 10 / 100
Text11.Text = 200
Text12.Text = 480
Text13.Text = Val(Text9) + Val(Text11) + Val(Text12)
Text14.Text = Val(Text1) + Val(Text8) - Val(Text14)
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Update
MsgBox "RECORD SAVED"
End Sub

Private Sub Command3_Click()
End
End Sub


Private Sub Command9_Click()
Form3.Show
End Sub

Private Sub Form_Load()
connectdb
End Sub




