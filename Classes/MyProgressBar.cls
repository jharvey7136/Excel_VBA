VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------ PROGRESS BAR -------------------------------------
Sub UpdateProgressBar(n As Long, m As Long, Optional DisplayText As String, Optional calcLabel As String)

    '// Omit DisplayText to display progress percentage
    On Error GoTo ERR_HANDLE

    If n >= m Then
        ProgressBar.Hide
    Else
        If ProgressBar.Visible = False Then ProgressBar.Show
        ProgressBar![Label1].Caption = calcLabel
        ProgressBar![BoxProgress].Caption = IIf(DisplayText = "", Round(((n / m) * 10000) / 100) & "%", DisplayText)
        ProgressBar![BoxProgress].Width = (n / m) * 204
        DoEvents
    End If
    Exit Sub

ERR_HANDLE:
    Err.Clear
    ProgressBar.Hide
End Sub














