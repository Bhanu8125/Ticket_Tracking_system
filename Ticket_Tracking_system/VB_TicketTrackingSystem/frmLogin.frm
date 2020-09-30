VERSION 5.00
Object = "{15C9A77A-F2C0-4CE1-B79D-1C8E3C2197EE}#2.0#0"; "Password_OCX.ocx"
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin PasswordOCX.PasswordControl PasswordControl1 
      Height          =   375
      Left            =   2220
      TabIndex        =   8
      Top             =   1380
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2220
      TabIndex        =   7
      Top             =   2400
      Width           =   915
   End
   Begin VB.CommandButton cmdSumbit 
      Caption         =   "Sumbit"
      Height          =   495
      Left            =   660
      TabIndex        =   6
      Top             =   2460
      Width           =   915
   End
   Begin VB.ComboBox cmdDepartment 
      Height          =   315
      Left            =   2220
      TabIndex        =   5
      Top             =   1920
      Width           =   2235
   End
   Begin VB.TextBox txtEmployeeId 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblDepartment 
      AutoSize        =   -1  'True
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1380
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblUserId 
      AutoSize        =   -1  'True
      Caption         =   "Employee Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   960
      Width           =   1320
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Depts As New Collection
Private Sub Text1_Change()

End Sub
Private Sub txtDepartment_Change()

End Sub
Private Sub cmdSumbit_Click()
'    Dim IsChecked As Boolean
'    IsChecked = PasswordControl.
'    If IsChecked Then
'        MsgBox PasswordControl
End Sub

Private Sub Form_Load()
    Set Depts = Loginrepository.GetDept
    FillDeptCombo
End Sub
Private Sub FillDeptCombo()
    For Index = 1 To Depts.Count
         cmdDepartment.AddItem Depts(Index)
    Next
End Sub

