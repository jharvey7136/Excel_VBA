VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyChkBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'class to attach click listeners to all the rate reset op selection check boxes

Option Explicit

Private WithEvents chkBox As MSForms.CheckBox
Attribute chkBox.VB_VarHelpID = -1

Public Sub AssignClicks(ctrl As Control)
    Set chkBox = ctrl
End Sub

Private Sub chkBox_Click()
    RateResetForm.CheckRunSelected  'enable/disable run selection button depending on if any boxes are checked
End Sub

