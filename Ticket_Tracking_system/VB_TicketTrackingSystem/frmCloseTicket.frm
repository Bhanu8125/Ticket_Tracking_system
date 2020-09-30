VERSION 5.00
Begin VB.Form frmCloseTicket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Close Ticket"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   2340
      TabIndex        =   7
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
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
      Left            =   780
      TabIndex        =   6
      Top             =   1980
      Width           =   1215
   End
   Begin VB.TextBox txtResolution 
      Height          =   315
      Left            =   1740
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1260
      Width           =   2715
   End
   Begin VB.ComboBox cmbEmployees 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.ComboBox cmbTicketId 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblResolution 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resolution"
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
      Width           =   1125
   End
   Begin VB.Label lblResolvedBY 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resolved By"
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
      Width           =   1350
   End
   Begin VB.Label lblTicketId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TicketId"
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
      Top             =   300
      Width           =   855
   End
End
Attribute VB_Name = "frmCloseTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Emplist As New Collection
Dim Tickets As New Collection
Private Sub cmdSubmit_Click()
    On Error GoTo errHand
    If cmbEmployees.Text <> Empty And cmbTicketId.Text <> Empty And txtResolution.Text <> Empty Then
        Dim IsClosed As Boolean
        IsClosed = CloseTicketRepository.UpdateTicket(CInt(cmbTicketId.Text), CStr(cmbEmployees.Text), CStr(txtResolution.Text))
        If IsClosed = True Then MsgBox CloseTicketRepository.Message, vbOKOnly, "Close Ticket"
        
    End If
    Exit Sub
errHand:
    MsgBox Err.Description
'  If Err.Number = 1001 Then
'    Err.Raise 1001, , "Error in getting Data , check logo file foe more details"
'    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
      Set Emplist = CloseTicketRepository.GetEmployees
      FillEmpCombo
      Set Tickets = CloseTicketRepository.GetTicketId
      FillTicketsCombo
      Exit Sub
errHand:
    If Err.Number = 1001 Then
        MsgBox "Error While Getting Data", vbOKOnly, "Close Report"
     End If
End Sub
Private Sub FillEmpCombo()
     For Index = 1 To Emplist.Count
         cmbEmployees.AddItem Emplist(Index)
    Next
End Sub
Private Sub FillTicketsCombo()
     For Index = 1 To Emplist.Count
         cmbTicketId.AddItem Tickets(Index)
    Next
End Sub

