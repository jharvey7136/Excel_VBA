VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'' private variables
'Private sPercentage As Single
'Private sLabel As String
'
'' properties
'Property Let sPercent(sPer As Single)
'    sPercentage = sPer
'End Property
'
'Property Get sPercent() As Single
'    sPercent = sPercentage
'End Property
'
'Property Let progLabel(sLab As String)
'    sLabel = sLab
'End Property
'
'Property Get progLabel() As String
'    progLabel = sLabel
'End Property
'
'Sub ShowBar()
'    ProgressBar.Show
'End Sub
'
''Sub SetLabel(sLabel As String)
''    ProgressBar.Caption = sLabel
''End Sub
'
'Sub UpdateStatus()
'    ProgressBar.Caption = sLabel
'    ProgressBar.Text.Caption = sPercentage * 100 & "% Complete"
'    ProgressBar.Bar.Width = sPercentage * (ProgressBar.FrameProgress.Width)
'    DoEvents
'End Sub
'
''Sub GetPercentDone(pctDone As Single)
''    sPercentage = pctDone
''    'UpdateStatus
''End Sub
'
''Sub FractionComplete(pctDone As Single)
''    ProgressBar.Text.Caption = pctDone * 100 & "% Complete"
''    ProgressBar.Bar.Width = pctDone * 2
''    DoEvents
''End Sub
'
'Sub RemoveBar()
'    Unload ProgressBar
'End Sub
'
'
'
'
'Sub UpdateProgressBar(n As Long, m As Long, Optional DisplayText As String)
''// DarkSprout April08
''// Omit DisplayText to display progress percentage
'On Error GoTo ERR_HANDLE
'
'If n >= m Then
'    ProgressBar.Hide
'Else
'    If ProgressBar.Visible = False Then ProgressBar.Show
'    ProgressBar![BoxProgress].Caption = IIf(DisplayText = "", Round(((n / m) * 10000) / 100) & "%", DisplayText)
'    ProgressBar![BoxProgress].Width = (n / m) * 468
'    DoEvents
'End If
'Exit Sub
'
'ERR_HANDLE:
'    Err.Clear
'    ProgressBar.Hide
'End Sub
'
'
'
'
'
'
'
