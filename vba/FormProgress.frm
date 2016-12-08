VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProgress 
   Caption         =   "Laddning pågår"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FormProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Progress As Integer

Public Property Get Progress() As Integer
    Progress = m_Progress '
End Property

Public Property Let Progress(Value As Integer)
    m_Progress = Value
    
    If (m_Progress = -1) Then
        ProgressBar1.Value = 100
        Me.Hide
        Unload Me
    Else
        ProgressBar1.Value = m_Progress Mod 100
        Me.Repaint
    End If
End Property

Public Property Let Title(Value As String)
    Me.lbltitle = Value
    Me.Repaint
End Property


Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
        m_Progress = 0
        
        Call FormHelper.SetFormDefaultColors(Me)
        
    Exit Sub
ErrorHandler:
    UI.ShowError ("FormProgress.UserForm_Initialize")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler
    Cancel = True
    
    Exit Sub
ErrorHandler:
    UI.ShowError ("FormProgress.UserForm_QueryClose")
End Sub
