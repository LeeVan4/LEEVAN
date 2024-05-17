VERSION 5.00
Begin VB.Form frmemployee 
   Caption         =   "Employee"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17490
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   17490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find"
      DisabledPicture =   "frmemployee.frx":0000
      DownPicture     =   "frmemployee.frx":0442
      DragIcon        =   "frmemployee.frx":0884
      Height          =   855
      Left            =   14280
      Picture         =   "frmemployee.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtaddress 
      Height          =   735
      Left            =   960
      TabIndex        =   14
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   855
      Left            =   14280
      Picture         =   "frmemployee.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      DisabledPicture =   "frmemployee.frx":154A
      DownPicture     =   "frmemployee.frx":198C
      Height          =   855
      Left            =   14280
      Picture         =   "frmemployee.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      DisabledPicture =   "frmemployee.frx":2210
      DownPicture     =   "frmemployee.frx":2652
      Height          =   855
      Left            =   14280
      Picture         =   "frmemployee.frx":2A94
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      DisabledPicture =   "frmemployee.frx":2ED6
      DownPicture     =   "frmemployee.frx":3318
      Height          =   855
      Left            =   14280
      Picture         =   "frmemployee.frx":375A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txthire 
      Height          =   735
      Left            =   4680
      TabIndex        =   8
      Top             =   3600
      Width           =   3015
   End
   Begin VB.TextBox txtsalary 
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtposition 
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtid 
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   15
      Top             =   3120
      Width           =   1245
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Date Hired:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   9
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Salary:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   7
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Position:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Complete Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   3
      Top             =   1920
      Width           =   1515
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
txtid.Text = ""
txtname.Text = ""
txtaddress.Text = ""
txtposition.Text = ""
txtsalary.Text = ""
txthire.Text = ""
txtid.SetFocus

End Sub

Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub cmddelete_Click()
conPayroll.Execute "Delete * from employee where employeeid='" & Trim(txtid.Text) & "'"
MsgBox "Record has been deleted.."
End Sub

Private Sub cmdFind_Click()
    Dim searchEmployeeID As String
    Dim foundEmployeeID As String
    
    searchEmployeeID = InputBox("Enter the Employee ID to find:", "Find Employee")
    
    openrstEmployee "SELECT * FROM employee WHERE employeeid='" & Trim(searchEmployeeID) & "'"
    
    If Not rstEmployee.EOF Then
        With rstEmployee
            txtid.Text = .Fields("employeeid").Value
            txtname.Text = .Fields("employeename").Value
            txtaddress.Text = .Fields("address").Value
            txtposition.Text = .Fields("position").Value
            txtsalary.Text = .Fields("salary").Value
            txthire.Text = .Fields("datehired").Value
        End With
        MsgBox "Employee found."
    Else
        MsgBox "Employee not found."
    End If

    rstEmployee.Close
End Sub


Private Sub cmdsave_Click()
openrstEmployee "Select * from employee where employeeid='" & Trim(txtid.Text) & "'"
If Not rstEmployee.EOF Then
    'not found
    With rstEmployee
        .Edit
            .Fields("employeeid").Value = txtid.Text
            .Fields("employeename").Value = txtname.Text
            .Fields("address").Value = txtaddress.Text
            .Fields("position").Value = txtposition.Text
            .Fields("salary").Value = txtsalary.Text
            .Fields("datehired").Value = txthire.Text
            
        .Update
        
        
    End With
Else
    'found
        With rstEmployee
        .AddNew
             .Fields("employeeid").Value = txtid.Text
            .Fields("employeename").Value = txtname.Text
            .Fields("address").Value = txtaddress.Text
            .Fields("position").Value = txtposition.Text
            .Fields("salary").Value = txtsalary.Text
            .Fields("datehired").Value = txthire.Text
        .Update
        
        End With
End If

    
End Sub

Private Sub Form_Load()
openWORKSPACEODBC
openconPayroll

End Sub

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtposition.SetFocus
End If

End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    openrstEmployee "Select * from employee where employeeid ='" & Trim(txtid.Text) & "'"
     If Not rstEmployee.EOF Then
        With rstEmployee
            txtid.Text = .Fields("employeeid").Value
            txtname.Text = .Fields("employeename").Value
            txtaddress.Text = .Fields("address").Value
            txtposition.Text = .Fields("position").Value
            txtsalary.Text = .Fields("salary").Value
            txthire.Text = .Fields("datehired").Value
        End With
    End If
    
End If

End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtaddress.SetFocus
End If

End Sub

Private Sub txtposition_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsalary.SetFocus
End If

End Sub

Private Sub txtsalary_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdatehired.SetFocus
End If
End Sub
