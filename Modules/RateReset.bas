Attribute VB_Name = "RateReset"
'
'                       RATE RESET
' -Public facing module to run the rate reset subroutines
' -MyRateReset Class used to modularize subs/functions for easy debugging
' -Log file located at: "C:\temp\VB_Logger\RateResetLog.txt"
'

Sub RunRateReset()
    On Error GoTo eh                        ' error handling
    Application.Calculation = xlAutomatic   ' turn on auto calculations if previously turned off
    
    Dim oRateReset As New MyRateReset       ' initialize new rate reset object
    
    oRateReset.RemainingLoanTerm            ' calc remaining loan term in months
    oRateReset.Auth                         ' build auth codes
    oRateReset.FormatRR                     ' apply correct formatting
    oRateReset.JoinData
    oRateReset.SetMaxExtWithTransScores     ' calc new max extension with trans scores
    oRateReset.AdjustMaxExt                 ' adjust max ext if necessary
    oRateReset.DeDuplicate                  ' remove duplicate loans
    oRateReset.RemoveMinMax12               ' remove loans where minExt = maxExt = 12 months
    oRateReset.FilterRemainingTerm          ' show loans with >13 months remaining
    oRateReset.CleanUp                      ' move final data to new sheet
    
    
    
Done:
    Set oRateReset = Nothing                ' terminate rate reset object
    Exit Sub
eh:
    MsgBox Err.Source & ": RunRateReset() error occured " & Err.Description
    Exit Sub
End Sub
