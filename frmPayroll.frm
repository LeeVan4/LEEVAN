VERSION 5.00
Begin VB.Form txtmonthlysalary 
   Caption         =   "Form Payroll"
   ClientHeight    =   9510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15870
   LinkTopic       =   "Form1"
   Picture         =   "frmPayroll.frx":0000
   ScaleHeight     =   9510
   ScaleWidth      =   15870
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnetincome 
      Height          =   615
      Left            =   11040
      TabIndex        =   53
      Top             =   7320
      Width           =   3375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Net Income"
      Height          =   615
      Left            =   9720
      TabIndex        =   52
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Compute Deduction"
      Height          =   615
      Left            =   9720
      TabIndex        =   51
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txttotdeduction 
      Height          =   615
      Left            =   11040
      TabIndex        =   50
      Top             =   6480
      Width           =   3375
   End
   Begin VB.TextBox txtpag 
      Height          =   495
      Left            =   11040
      TabIndex        =   48
      Top             =   5880
      Width           =   3375
   End
   Begin VB.TextBox txtphil 
      Height          =   495
      Left            =   11040
      TabIndex        =   47
      Top             =   5280
      Width           =   3375
   End
   Begin VB.TextBox txttax 
      Height          =   495
      Left            =   11040
      TabIndex        =   45
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox txtsss1 
      Height          =   495
      Left            =   11040
      TabIndex        =   43
      Top             =   4080
      Width           =   3375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Compute Gross"
      Height          =   615
      Left            =   960
      TabIndex        =   40
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox txtGrossPay 
      Height          =   615
      Left            =   2760
      TabIndex        =   39
      Top             =   6840
      Width           =   3375
   End
   Begin VB.TextBox txtmeal 
      Height          =   615
      Left            =   2760
      TabIndex        =   37
      Top             =   6120
      Width           =   3375
   End
   Begin VB.TextBox txtperhour 
      Height          =   615
      Left            =   2760
      TabIndex        =   35
      Top             =   5400
      Width           =   3375
   End
   Begin VB.TextBox txtperday 
      Height          =   615
      Left            =   2760
      TabIndex        =   33
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox txt15th 
      Height          =   615
      Left            =   2760
      TabIndex        =   31
      Top             =   3960
      Width           =   3375
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   855
      Left            =   13800
      Picture         =   "frmPayroll.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   855
      Left            =   10320
      Picture         =   "frmPayroll.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find"
      Height          =   975
      Left            =   5400
      Picture         =   "frmPayroll.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   975
      Left            =   4080
      Picture         =   "frmPayroll.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   975
      Left            =   2760
      Picture         =   "frmPayroll.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      Height          =   975
      Left            =   1440
      Picture         =   "frmPayroll.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   10320
      Width           =   1095
   End
   Begin VB.TextBox txtpagibig 
      Height          =   495
      Left            =   12240
      TabIndex        =   21
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtphilhealth 
      Height          =   495
      Left            =   12240
      TabIndex        =   19
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txttin 
      Height          =   495
      Left            =   12240
      TabIndex        =   16
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtSSS 
      Height          =   495
      Left            =   12240
      TabIndex        =   15
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtdatehired 
      Height          =   495
      Left            =   12240
      TabIndex        =   12
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtdateto 
      Height          =   495
      Left            =   6600
      TabIndex        =   10
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtdatefrom 
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtsalary 
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox txtempID 
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txttranno 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label23 
      Caption         =   "Pag-ibig"
      Height          =   375
      Left            =   9840
      TabIndex        =   49
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "Philhealth"
      Height          =   255
      Left            =   9840
      TabIndex        =   46
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "TIN"
      Height          =   255
      Left            =   9840
      TabIndex        =   44
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   "SSS"
      Height          =   255
      Left            =   9720
      TabIndex        =   42
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "List of Deductions"
      Height          =   375
      Left            =   9240
      TabIndex        =   41
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label17 
      Caption         =   "Gross Pay"
      Height          =   375
      Left            =   840
      TabIndex        =   38
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Meal/ Travel Allowance"
      Height          =   375
      Left            =   840
      TabIndex        =   36
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label15 
      Caption         =   "Rate per hour"
      Height          =   375
      Left            =   840
      TabIndex        =   34
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "Rate per day"
      Height          =   375
      Left            =   840
      TabIndex        =   32
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Rate per 15th day of the month"
      Height          =   375
      Left            =   840
      TabIndex        =   30
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Breakdown of Wages"
      Height          =   375
      Left            =   600
      TabIndex        =   29
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Pag-ibig#"
      Height          =   495
      Left            =   10560
      TabIndex        =   20
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "PHILHEALTH#"
      Height          =   495
      Left            =   10560
      TabIndex        =   18
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "TIN#"
      Height          =   495
      Left            =   10560
      TabIndex        =   17
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "SSS#"
      Height          =   495
      Left            =   10560
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Date Hired:"
      Height          =   495
      Left            =   10560
      TabIndex        =   13
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Date Covered To:"
      Height          =   495
      Left            =   4920
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Date Covered From:"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Monthly Salary"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Complete Name"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Employee ID"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Transaction No."
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "txtmonthlysalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
On Error Resume Next

txttranno.SelStart = 0
txttranno.SelLength = Len(txttranno.Text)
txttranno.SetFocus

txttranno.Text = ""
txtempID.Text = ""
txtname.Text = ""
txtmonthlysalary.Text = ""
txtdatefrom.Value = ""
txtdateto.Value = ""
txtdatehired.Value = ""
txtSSS.Text = ""
txttin.Text = ""
txtphilhealth.Text = ""
txtpagibig.Text = ""
txt15th.Text = ""
txtperday.Text = ""
txtperhour.Text = ""
txtmeal.Text = ""
txtGrossPay.Text = ""
txtsss1.Text = ""
txttax.Text = ""
txtphil.Text = ""
txtpag.Text = ""
txttotdeduction.Text = ""
txtnetincome.Text = ""

End Sub

Private Sub cmddeduct_Click()
Dim xsss As Single
Dim xtax As Single
Dim xpag As Single
Dim xphil As Single
Dim xtOTD As Double
xsss = 300
txtsss1.Text = xsss

xtax = 500
txttax.Text = xtax

xpag = 100
txtpag.Text = xpag

xphil = 100
txtphil.Text = xphil

xtOTD = xsss + xtax + xphil + xpag

txttotdeduction.Text = xtOTD


End Sub

Private Sub cmdGross_Click()
Dim xrate15 As Double
Dim xSalary As Double
Dim xrateperday As Double
Dim xrateperhour As Double
Dim xmeal As Double
Dim xGross As Double



xSalary = txtmonthlysalary.Text
xrate15 = xSalary / 2
txt15th.Text = xrate15

xrateperday = txtmonthlysalary.Text / 26
txtperday.Text = xrateperday

xrateperhour = txtperday.Text / 8
txtperhour.Text = xrateperhour

xmeal = 500

txtmeal.Text = xmeal

xGross = xmeal + xrate15

txtGrossPay.Text = xGross

End Sub

Private Sub cmdnet_Click()
Dim xNet As Double
Dim xG As Double
Dim xD As Double


xG = txtGrossPay.Text
xD = txttotdeduction.Text

xNet = xG - xD
txtnetincome.Text = xNet

End Sub

Private Sub cmddelete_Click()
conPayroll.Execute "Delete * from employee where employeeid='" & Trim(txtid.Text) & "'"
MsgBox "Record has been deleted.."
End Sub

Private Sub cmdclose_Click()
Unload Me

End Sub

Private Sub cmdFind_Click()
    Dim searchID As String
    searchID = InputBox("Enter Employee ID to find:", "Find Employee")
    
    ' Check if search ID is provided
    If searchID <> "" Then
        ' Open the recordset to find the employee
        openrstEmployee "SELECT * FROM EMPLOYEE WHERE EMPLOYEEID='" & Trim(searchID) & "'"
        
        ' Check if the employee is found
        If Not rstEmployee.EOF Then
            ' Populate the form fields with employee information
            With rstEmployee
                txtempID.Text = .Fields("employeeid").Value
                txtname.Text = .Fields("empname").Value
                txtmonthlysalary.Text = .Fields("salary").Value
                txtdatehired.Value = .Fields("datehired").Value
            End With
        Else
            ' Display a message if employee is not found
            MsgBox "Employee not found.", vbExclamation
        End If
    End If
End Sub

Private Sub cmdsave_Click()
OPENRSTPAYROLL "SELECT * FROM payroll where tranno='" & Trim(txttranno.Text) & "'"
If Not rstPayroll.EOF Then
'if not found
    With rstPayroll
        .Edit
            .Fields("tranno").Value = Trim(txttranno.Text)
            .Fields("employeeid").Value = Trim(txtempID.Text)
            .Fields("datefrom").Value = Trim(txtdatefrom.Value)
            .Fields("dateto").Value = Trim(txtdateto.Value)
            .Fields("rate15").Value = Trim(txt15th.Text)
            .Fields("rateperday").Value = Trim(txtperday.Text)
            .Fields("rateperhour").Value = Trim(txtperhour.Text)
            .Fields("meal").Value = Trim(txtmeal.Text)
            .Fields("grosspay").Value = Trim(txtGrossPay.Text)
            .Fields("datehired").Value = Trim(txtdatehired.Value)
            .Fields("sssno").Value = Trim(txtSSS.Text)
            .Fields("tinno").Value = Trim(txttin.Text)
            .Fields("philhealthno").Value = Trim(txtphilhealth.Text)
            .Fields("pagibigno").Value = Trim(txtpagibig.Text)
            .Fields("sss").Value = Trim(txtsss1.Text)
            .Fields("tax").Value = Trim(txttax.Text)
            .Fields("pagibig").Value = Trim(txtpag.Text)
            .Fields("philhealth").Value = Trim(txtphil.Text)
            .Fields("totaldeduction").Value = Trim(txttotdeduction.Text)
            .Fields("netincome").Value = Trim(txtnetincome.Text)
            
        .Update
        
    End With
Else
    'not found
        With rstPayroll
            .AddNew
                .Fields("tranno").Value = Trim(txttranno.Text)
                .Fields("employeeid").Value = Trim(txtempID.Text)
                .Fields("datefrom").Value = Trim(txtdatefrom.Value)
                .Fields("dateto").Value = Trim(txtdateto.Value)
                .Fields("rate15").Value = Trim(txt15th.Text)
                .Fields("rateperday").Value = Trim(txtperday.Text)
                .Fields("rateperhour").Value = Trim(txtperhour.Text)
                .Fields("meal").Value = Trim(txtmeal.Text)
                .Fields("grosspay").Value = Trim(txtGrossPay.Text)
                .Fields("datehired").Value = Trim(txtdatehired.Value)
                .Fields("sssno").Value = Trim(txtSSS.Text)
                .Fields("tinno").Value = Trim(txttin.Text)
                .Fields("philhealthno").Value = Trim(txtphilhealth.Text)
                .Fields("pagibigno").Value = Trim(txtpagibig.Text)
                .Fields("sss").Value = Trim(txtsss1.Text)
                .Fields("tax").Value = Trim(txttax.Text)
                .Fields("pagibig").Value = Trim(txtpagibig.Text)
                .Fields("philhealth").Value = Trim(txtphil.Text)
                .Fields("totaldeduction").Value = Trim(txttotdeduction.Text)
                .Fields("netincome").Value = Trim(txtnetincome.Text)
            .Update
            
        End With
End If

    
End Sub

Private Sub Form_Load()
openWORKSPACEODBC
openconPayroll
txttranno.Text = GenerateTransactionNumber()
End Sub

Private Function GenerateTransactionNumber() As String
    ' Your code to generate a transaction number
    GenerateTransactionNumber = "TXN" & Format(Now(), "YYYYMMDDHHMMSS")
End Function
