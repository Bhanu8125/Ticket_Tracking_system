VERSION 5.00
Begin VB.UserControl PasswordControl 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2280
   ScaleHeight     =   390
   ScaleWidth      =   2280
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   60
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "PasswordControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Capital As Integer
Public Small As Integer
Public Special As String
Public Function Check() As Boolean
    Capital = 0
    Small = 0
    Special = 0
   If txtPassword.Text <> Empty Then
        If Len(txtPassword.Text) < 8 Then
            MsgBox "Password Should be minimum 8 characters", vbOKOnly, "Password"
        End If
        For Index = 1 To Len(txtPassword.Text)
            If Asc(Mid(txtPaswword.Text, Index, 1)) >= 65 And Asc(Mid(txtPaswword.Text, Index, 1)) <= 90 Then
                Capital = Capital + 1
            ElseIf Asc(Mid(txtPaswword.Text, Index, 1)) >= 97 And Asc(Mid(txtPaswword.Text, Index, 1)) <= 122 Then
                Small = Small + 1
            Else
                Special = Special + 1
            End If
        Next
        If Capital = 0 Then
            MsgBox "One UpperCase Character is Mandatory", vbOKOnly, "Password"
        ElseIf Small = 0 Then
            MsgBox "One LowerCase Character is Mandatory", vbOKOnly, "Password"
        ElseIf Special = 0 Then
            MsgBox "One Special Character is Mandatory", vbOKOnly, "Password"
        End If
   End If
End Function

