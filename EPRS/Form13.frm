VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form13 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form13"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form13"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Home"
      Height          =   735
      Left            =   12480
      TabIndex        =   68
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   855
      Left            =   10800
      TabIndex        =   67
      Top             =   7800
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   11160
      Top             =   9000
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "income"
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
   Begin VB.CommandButton Command3 
      Caption         =   "Calculate Net Taxable Income"
      Height          =   855
      Left            =   11400
      TabIndex        =   66
      Top             =   6720
      Width           =   2175
   End
   Begin VB.TextBox Text30 
      BackColor       =   &H80000003&
      DataField       =   "ded"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      TabIndex        =   65
      Text            =   "0"
      Top             =   9240
      Width           =   1215
   End
   Begin VB.TextBox Text29 
      BackColor       =   &H80000003&
      DataField       =   "ntt"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   14040
      TabIndex        =   64
      Text            =   "0"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text28 
      BackColor       =   &H80000003&
      DataField       =   "ec"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   14040
      TabIndex        =   62
      Text            =   "0"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text27 
      BackColor       =   &H80000002&
      DataField       =   "ttp"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   14040
      TabIndex        =   60
      Text            =   "0"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   14040
      TabIndex        =   58
      Text            =   "0"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   14040
      TabIndex        =   57
      Text            =   "0"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   12240
      TabIndex        =   56
      Text            =   "0"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   12240
      TabIndex        =   55
      Text            =   "0"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calculate Total Taxable Income"
      Height          =   735
      Left            =   4440
      TabIndex        =   49
      Top             =   10320
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H80000003&
      DataField       =   "ttp"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   48
      Text            =   "0"
      Top             =   10200
      Width           =   1815
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   4800
      TabIndex        =   46
      Text            =   "0"
      Top             =   9240
      Width           =   1575
   End
   Begin VB.TextBox Text16 
      DataField       =   "crf"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6720
      TabIndex        =   43
      Text            =   "0"
      Top             =   9720
      Width           =   1455
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   4920
      TabIndex        =   42
      Text            =   "0"
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   4920
      TabIndex        =   41
      Text            =   "0"
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   4920
      TabIndex        =   40
      Text            =   "0"
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   4920
      TabIndex        =   39
      Text            =   "0"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H80000003&
      DataField       =   "gtt"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   32
      Text            =   "0"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      DataField       =   "lhp"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6840
      TabIndex        =   30
      Text            =   "0"
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H80000003&
      DataField       =   "nsi"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      Text            =   "0"
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      DataField       =   "ld"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6840
      TabIndex        =   25
      Text            =   "0"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000003&
      DataField       =   "hra"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Text            =   "0"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4920
      TabIndex        =   21
      Text            =   "0"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Text            =   "0"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Text            =   "0"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "ts"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Text            =   "0"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New"
      Height          =   735
      Left            =   9000
      TabIndex        =   9
      Top             =   7920
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "year"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Form13.frx":0000
      Left            =   2520
      List            =   "Form13.frx":0010
      TabIndex        =   8
      Text            =   "Combo2"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "pan"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   13560
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "ename"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   10200
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "empid"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Form13.frx":0040
      Left            =   6840
      List            =   "Form13.frx":0042
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label28 
      Caption         =   "12. Net Taxable Income"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   63
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label27 
      Caption         =   "11. Education Cess (3%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   61
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label26 
      Caption         =   "Total Tax Payable"
      Height          =   615
      Left            =   9480
      TabIndex        =   59
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label23 
      Caption         =   "Rs 2,50,000 to Rs 5,00,000"
      Height          =   615
      Left            =   9480
      TabIndex        =   54
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label22 
      Caption         =   "Up to Rs 2,50,000"
      Height          =   495
      Left            =   9480
      TabIndex        =   53
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label21 
      Caption         =   "10. Tax on the above"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   52
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Line Line5 
      X1              =   13920
      X2              =   13920
      Y1              =   1560
      Y2              =   6360
   End
   Begin VB.Line Line4 
      X1              =   12120
      X2              =   12120
      Y1              =   1560
      Y2              =   6360
   End
   Begin VB.Label Label7 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   14280
      TabIndex        =   51
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   12240
      TabIndex        =   50
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Line Line3 
      X1              =   8760
      X2              =   8760
      Y1              =   1560
      Y2              =   9720
   End
   Begin VB.Label Label20 
      Caption         =   "9. Taxable Total Income (6-7-8)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   47
      Top             =   10200
      Width           =   3495
   End
   Begin VB.Label Label13 
      Caption         =   "8. U/S 80 G Donation to CRF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   45
      Top             =   9720
      Width           =   3375
   End
   Begin VB.Label Label19 
      Caption         =   "Particular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   44
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label18 
      Caption         =   "Total Deduction"
      Height          =   375
      Left            =   840
      TabIndex        =   38
      Top             =   9240
      Width           =   3015
   End
   Begin VB.Label Label17 
      Caption         =   "LIC"
      Height          =   375
      Left            =   840
      TabIndex        =   37
      Top             =   8880
      Width           =   2535
   End
   Begin VB.Label Label16 
      Caption         =   "NPS"
      Height          =   495
      Left            =   840
      TabIndex        =   36
      Top             =   8520
      Width           =   3015
   End
   Begin VB.Label Label15 
      Caption         =   "PPF/ULIP"
      Height          =   495
      Left            =   840
      TabIndex        =   35
      Top             =   8160
      Width           =   3255
   End
   Begin VB.Label Label14 
      Caption         =   "Housing Loan Principal"
      Height          =   375
      Left            =   840
      TabIndex        =   34
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label13 
      Caption         =   "7. Deduction under 80 c"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   33
      Top             =   7440
      Width           =   3375
   End
   Begin VB.Label Label12 
      Caption         =   "6. Gross Total Income"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   31
      Top             =   6840
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "Interest Paid on housing loan"
      Height          =   615
      Index           =   7
      Left            =   1080
      TabIndex        =   29
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label9 
      Caption         =   "5. Less: Loss from House Property"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   360
      TabIndex        =   28
      Top             =   6000
      Width           =   4095
   End
   Begin VB.Label Label11 
      Caption         =   "4. Net Salary Income(1-2-3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   26
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "a Professional Tax"
      Height          =   615
      Index           =   5
      Left            =   960
      TabIndex        =   24
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label9 
      Caption         =   "3. Less Deduction u/s 16(i) A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   240
      TabIndex        =   23
      Top             =   4560
      Width           =   4335
   End
   Begin VB.Label Label9 
      Caption         =   "c. 40% of Salary"
      Height          =   615
      Index           =   3
      Left            =   960
      TabIndex        =   18
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label10 
      Caption         =   "b. House Rent Paid in excess of 10% salary"
      Height          =   615
      Left            =   960
      TabIndex        =   17
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label Label9 
      Caption         =   "a. Actual Salary Received"
      Height          =   615
      Index           =   2
      Left            =   1080
      TabIndex        =   16
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label9 
      Caption         =   "2.  HRA Exemption u/s 10(13A)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   6480
      Y1              =   1320
      Y2              =   10080
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   4680
      Y1              =   1320
      Y2              =   10080
   End
   Begin VB.Label Label9 
      Caption         =   "1. Total Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label8 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Particular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "PAN"
      Height          =   735
      Left            =   12120
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Employee Name"
      Height          =   495
      Left            =   8760
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Employee ID "
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Assessment Year"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Income Tax Calculation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form13"
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
            Text3.Text = rs("panno")

             End If
            If rs.State = adStateOpen Then rs.Close
  End Sub
Private Sub Command1_Click()

Adodc1.Recordset.AddNew
Text1.Text = 0
Text6.Text = 0
Text8.Text = 0
Text9.Text = 0
Text10.Text = 0
Text11.Text = 0

Set rs = con.Execute("select empid from empdet")
While (Not rs.EOF)
    Combo1.AddItem rs(0)
     rs.MoveNext
     
Wend
rs.Close


End Sub

Private Sub Command2_Click()

Text6.Text = Val(Text5) + Val(Text7) + Val(Text4)
Text9.Text = Val(Text1) - Val(Text6) - Val(Text8)
If Val(Text10) > 200000 Then
Text10.Text = 200000
End If
Text11.Text = Val(Text9) - Val(Text10)


Text17.Text = Val(Text12) + Val(Text13) + Val(Text14) + Val(Text15)
If (Text17.Text) <= 150000 Then
Text30.Text = Val(Text17)
Else
Text30.Text = 150000
End If

Text18.Text = Val(Text11) - Val(Text30) - Val(Text16)
End Sub

Private Sub Command3_Click()
Text19.Text = Val(Text18) - 250000

If Val(Text19) > 0 Then
Text22.Text = Val(Text19)

ElseIf Val(Text19) <= 0 Then
Text22.Text = 0
End If
'Text20.Text = Val(Text22) - 500000
If Val(Text22) <= 0 Then
Text23.Text = 0
Text20.Text = 0

ElseIf Val(Text22) > 0 And Val(Text22) <= 500000 Then
'Text20 = Val(Text22) - 500000
Text23.Text = Val(Text22) * 5 / 100
Text20 = 0
ElseIf Val(Text22) > 500000 Then
Text20 = Val(Text22) - 500000
Text23 = 500000 * 10 / 100
End If
Text27.Text = Val(Text23.Text) + Val((Text20) * (10 / 100))
Text28.Text = Val(Text27.Text) * 4 / 100
Text29.Text = Val(Text27) + Val(Text28)
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Update
MsgBox "Employee Income Tax Statement Calculated Successfully"
End Sub

Private Sub Command5_Click()
Form3.Show
End Sub

Private Sub Form_Load()
connectdb
End Sub

Private Sub Text20_Click()
Text20.Text = Val(Text22) - 500000
End Sub
