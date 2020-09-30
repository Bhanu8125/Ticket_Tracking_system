VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   6975
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuTicket 
      Caption         =   "Ticket"
      Begin VB.Menu mnuViewTicket 
         Caption         =   "View Ticket"
      End
      Begin VB.Menu mnuCreateTicket 
         Caption         =   "Create Ticket"
      End
   End
   Begin VB.Menu mnuTicketInformation 
      Caption         =   "Ticket Information"
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCreateTicket_Click()
    frmCreateTicket.Show
End Sub

Private Sub mnuTicketInformation_Click()
        Dim crApp As New CRAXDDRT.Application
        Dim crRpt As New CRAXDDRT.Report
        Dim Path As String
        Path = App.Path & "\Report.rpt"
        Set crRpt = crApp.OpenReport(Path)
        frmReport.CRViewer.ReportSource = crRpt
        frmReport.CRViewer.ViewReport
        frmReport.Show
End Sub

Private Sub mnuViewTicket_Click()
    frmCloseTicket.Show
End Sub
