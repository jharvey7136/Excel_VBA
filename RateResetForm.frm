VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RateResetForm 
   Caption         =   "Rate Reset Debugger"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8175
   OleObjectBlob   =   "RateResetForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RateResetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declarations
Dim oRateReset As New MyRateReset   'direct rate reset object
Dim oRRForm As New MyRRForm         'helper class for this form
Private sFileName As String         'filename given for the log.txt

'properties
Property Let sLogFileName(sName As String)  'assign string file name for logging
    sFileName = sName
End Property

Property Get sLogFileName() As String   'fetch string file name for logging
    sLogFileName = sFileName
End Property

'----------------------------------- BUTTONS -----------------------------------
Private Sub bRun_Click()    'Run button
    On Error GoTo eh

    Worksheets("Data").Activate
    
    If oRateReset.LogBool = True And sLogFileName = "" Then
        MsgBox "No log file set. Disable logging or select 'Browse' and choose a .txt file"
        bBrowse.SetFocus
        Exit Sub
    End If
    
    If oRateReset.LogBool = True And sLogFileName <> "" Then
        oRateReset.LogFilePath = sFileName
        oRateReset.SetLoggerPath
    End If
        
    oRRForm.ScanCheckboxs oRateReset 'call MyRRForm class to scan the checkboxs
    
Done:
    'Set oRateReset = Nothing                ' terminate rate reset object
    Exit Sub
eh:
    MsgBox Err.Source & ": bRun_Click() error occured " & Err.Description
    Exit Sub
End Sub

Private Sub bRunAll_Click()
    On Error GoTo eh
    Worksheets("Data").Activate
    
    If oRateReset.LogBool = True And sLogFileName = "" Then
        MsgBox "No log file set. Disable logging or select 'Browse' and choose a .txt file"
        bBrowse.SetFocus
        Exit Sub
    End If
    
    If oRateReset.LogBool = True And sLogFileName <> "" Then
        oRateReset.LogFilePath = sFileName
    End If
        
    oRRForm.RunAll oRateReset 'call MyRRForm class to scan the checkboxs
    
Done:
    'Set oRateReset = Nothing                ' terminate rate reset object
    Exit Sub
eh:
    MsgBox Err.Source & ": bRun_Click() error occured " & Err.Description
    Exit Sub
End Sub

Private Sub ToggleButton1_Click()   'toggle view ops on/off
    oRRForm.ToggleButton
End Sub

Private Sub bSelectAll_Click()  'select all checkbox ops
    oRRForm.CheckAll
End Sub

Private Sub bDeSelectAll_Click()    'deselect all checkbox ops
    oRRForm.UncheckAll
End Sub

Private Sub bReset_Click()  'reset data sheet with original
    oRRForm.ResetDataSheet
End Sub

Private Sub bCloseForm_Click()  'gracefully close form
    oRRForm.CloseRRForm
End Sub

'----------------------------------- END BUTTONS -----------------------------------

'------------------------- LOGGING -------------------------
Private Sub obNo_Click()        'logging disabled
    oRRForm.LogDisable oRateReset
End Sub

Private Sub obYes_Click()       'logging enabled
    oRRForm.LogEnable oRateReset
End Sub

Private Sub bViewLog_Click()    'view log button
    oRRForm.OpenLogFile sFileName
End Sub

Private Sub bBrowse_Click()     'select txt file for logging
    'oRRForm.ChooseLogFile sFileName
    Dim SelectedItem As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Select a .txt file to log to"
        .ButtonName = "Select"
        
        If .Show = -1 Then
            'if user clicks OK
            SelectedItem = .SelectedItems(1)
            MsgBox "You have selected: " & SelectedItem
            Let sLogFileName = SelectedItem
            
        Else
        'if user clicks Cancel
        End If
    End With
End Sub
'------------------------- END LOGGING -------------------------

'------------------------- USERFORM FUNCTIONS -------------------------
Private Sub UserForm_Initialize()
    oRRForm.UserFormInit
End Sub

Private Sub UserForm_Terminate()
    Set oRateReset = Nothing                ' terminate rate reset object
    Set oRRForm = Nothing
End Sub
'------------------------- END USERFORM FUNCTIONS -------------------------
