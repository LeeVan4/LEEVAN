VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mainmenu 
   BackColor       =   &H8000000C&
   Caption         =   "Payroll System"
   ClientHeight    =   6510
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15735
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   1588
      ButtonWidth     =   1561
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Employee"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Payroll"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":10DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":21B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuemployee 
         Caption         =   "Employee"
      End
      Begin VB.Menu mnuseparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "Transaction"
      Begin VB.Menu mnupayroll 
         Caption         =   "Payroll"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "Reports"
      Begin VB.Menu mnusummary 
         Caption         =   "Summary"
      End
   End
End
Attribute VB_Name = "mainmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public empSalary As Double
Private Sub mnuemployee_Click()
frmEmployee.Show
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnupayroll_Click()
frmPayroll.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case "1"
frmEmployee.Show vbModel
Case "2"
frmPayroll.Show vbModel
End Select
End Sub
