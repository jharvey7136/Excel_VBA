VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' private variables
Private sFileName As String

' properties
Property Let logFileName(sName As String)
    sFileName = sName
End Property

Property Get logFileName() As String
    logFileName = sFileName
End Property

' initializer
Private Sub Class_Initialize()
    'Logger "MyLogger", "Class_Initialize()", "Class Module MyLogger has been initialized"
End Sub

' public facing sub to do the logging - writes to a txt file
Public Sub Logger(sType As String, sSource As String, sDetails As String)
    On Error GoTo openFileError

    'sFilename = "C:\temp\VB_Logger\MyLogger.txt"
    
    ' open file to write
    Dim txtFile As String
    
    Dim filenumber As Variant
    filenumber = FreeFile
    
    Open sFileName For Append As #filenumber
    
    If sType = "TITLE" Then
        Print #filenumber, sSource & sDetails
        Close #filenumber
        Exit Sub
    End If
    
Done:
    On Error GoTo fileWriteError
    Print #filenumber, CStr(Now) & ": " & padType(sType) & padSource(sSource) & sDetails
    Close #filenumber
    Exit Sub
    
fileWriteError:
    MsgBox Err.Source & ": Logger error writing to file: " & Err.Description
    Exit Sub
    
openFileError:
    MsgBox Err.Source & ": Logger error opening file: " & Err.Description
    Exit Sub
End Sub

' formats the type string with padding
Private Function padType(pType As String) As String
    Dim typeLength As String
    typeLength = "1234567890123"    ' 13 characters long
    LSet typeLength = pType
    padType = typeLength
End Function

' formats the source string with padding
Private Function padSource(pSource As String) As String
    Dim sourceLength As String
    sourceLength = "123456789012345678901234567890" ' 30 characters long
    LSet sourceLength = pSource
    padSource = sourceLength
End Function

' terminates the logging class
Private Sub Class_Terminate()
    Dim myPath As String
    Dim folderPath As String

    folderPath = Application.ActiveWorkbook.Path
    myPath = Application.ActiveWorkbook.FullName
    
    Logger "FolderPath", "ActiveWorkbook", folderPath
    Logger "FilePath", "ActiveWorkbookName", myPath
    Logger "MyLogger", "Class_Terminate()", "Class Module MyLogger has been terminated." & vbNewLine & vbNewLine
End Sub




