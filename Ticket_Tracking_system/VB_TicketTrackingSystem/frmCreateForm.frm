VERSION 5.00
Begin VB.Form frmCreateTicket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Ticket"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   660
      TabIndex        =   8
      Top             =   2220
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Height          =   375
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1620
      Width           =   2355
   End
   Begin VB.ComboBox cmbSeverity 
      Height          =   315
      ItemData        =   "frmCreateForm.frx":0000
      Left            =   2040
      List            =   "frmCreateForm.frx":000D
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1140
      Width           =   2475
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   660
      Width           =   2415
   End
   Begin VB.ComboBox cmbEmployees 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   180
      Width           =   2475
   End
   Begin VB.Label lblTicketDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Description"
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
      Left            =   60
      TabIndex        =   6
      Top             =   1740
      Width           =   1905
   End
   Begin VB.Label lblSeverity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Severity"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1260
      Width           =   870
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   510
   End
   Begin VB.Label lblEmployeeName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
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
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1740
   End
End
Attribute VB_Name = "frmCreateTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Emplist As New Collection
Private Sub Text1_Change()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSubmit_Click()
    On Error GoTo errHand
    If cmbEmployees.Text <> Empty And txtDate.Text <> Empty And cmbSeverity.Text <> Empty And txtDescription.Text <> Empty Then
        If (Format(txtDate.Text, "DD-MM-YYYY HH:MM") <= Format(Now, "DD-MM-YYYY HHMM")) Then
            Dim check, splitt, checktime As Variant
            check = Split(txtDate.Text, "-")
            splitt = Split(check(2), " ")
            checktime = Split(splitt(1), ":")
            If check(0) < 1 And check(0) > 31 Or check(1) < 1 Or check(1) > 12 Then
                MsgBox " Invalid Date"
            End If
'            MsgBox check(0)
'            MsgBox check(1)
'            MsgBox check(2)
'            MsgBox checktime(0)
'            MsgBox checktime(1)
        
        Else
             MsgBox "Invalid Date", vbOKOnly, "Create Ticket"
        End If
        Dim IsInserted As Boolean
        IsInserted = CreateTicketRepository.insert(CStr(cmbEmployees.Text), CStr(txtDate.Text), CStr(cmbSeverity.Text), CStr(txtDescription.Text))
        If IsInserted Then
            MsgBox "Ticket Created Successfully", vbOKOnly, "Create Ticket"
        End If
    End If
    Exit Sub
errHand:
    
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    Set Emplist = CreateTicketRepository.GetEmployees
    FillEmployees
    Exit Sub
errHand:
    If Err.Number = 1001 Then
        MsgBox "Error While getting Data From Database", vbOKOnly, "Create Ticket"
     End If
End Sub
Private Sub FillEmployees()
        For Index = 1 To Emplist.Count
         cmbEmployees.AddItem Emplist(Index)
    Next
End Sub
