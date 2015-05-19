Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Time
' Level:        Framework module
' Version:      1.00
' Description:  File and directory related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - initial version
' =================================

' =================================
' FUNCTION:     FiscalYear
' Description:  Returns the fiscal year corresponding to the input date
' Parameters:   datDate - date value to be converted to fiscal year
'               blnFourDigits - flag to use 4 digits to represent the year (default True)
'               blnAddPrefix - flag to add a prefix to the result (default True)
'               strPrefix - prefix to be added to the string
' Returns:      variant for the fiscal year string or integer (e.g., "FY2010")
' Throws:       none
' References:   none
' Source/date:  From Front-end Application Builder v1.1, Simon Kingston, date unknown
' Revisions:    John R. Boetsch, 6/17/2009 - error trapping, documentation, added prefix & digit flags
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_Time
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function FiscalYear(ByVal datDate As Date, _
    Optional ByVal blnFourDigits As Boolean = True, _
    Optional ByVal blnAddPrefix As Boolean = True, _
    Optional ByVal strPrefix As String = "FY") As Variant

    On Error GoTo Err_Handler

    Dim intYear As Integer
    Dim strYear As String

    intYear = Year(datDate)
    If Month(datDate) >= 10 Then intYear = intYear + 1

    ' Year string depending on 2 or 4 characters
    If blnFourDigits Then
        strYear = CStr(intYear)
    Else
        strYear = Right(CStr(intYear), 2)
    End If

    If blnAddPrefix Then
        FiscalYear = strPrefix & strYear
    Else
        FiscalYear = strYear
    End If

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FiscalYear[mod_Time])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     Pause
' Description:  Pauses for specified number of section
' Parameters:   NumberOfSeconds - number of seconds to pause (variant)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  G Hudson, 3/10/2005
'               http://www.access-programmers.co.uk/forums/showthread.php?t=82953
' Revisions:    BLC, 5/18/2015 - initial version
' =================================
Public Function Pause(NumberOfSeconds As Variant)
On Error GoTo Err_Handler

    Dim PauseTime As Variant, Start As Variant

    PauseTime = NumberOfSeconds
    Start = Timer
    Do While Timer < Start + PauseTime
    DoEvents
    Loop

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NumberOfSeconds[mod_Time])"
    End Select
    Resume Exit_Function
End Function