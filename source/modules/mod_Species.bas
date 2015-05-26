Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Species
' Description:  species functions & procedures
'
' Source/date:  Bonnie Campbell, 4/9/2015
' Revisions:    BLC - 4/9/2015 - initial version
' =================================

' ---------------------------------
' FUNCTION:     PopulateSpeciesPriorities
' Description:  Populate species priority values from species priority concatenation
' Assumptions:  Park priority textboxes are named tbxPARKPriority (e.g. tbxZIONPriority)
' Parameters:   parkCode - 4 character park code (string)
'               priorities - species priority string concatenation for all parks (e.g. "BLCA-1|COLM-Transect|FOBU-1")
' Returns:      Priority - value for park species priority (string)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, April 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/9/2015 - initial version
' ---------------------------------
Public Function PopulateSpeciesPriorities(parkCode As String, priorities As String) As String

On Error GoTo Err_Handler

Dim ParkPriorities As Variant
Dim i As Integer

    'check if parkCode is in priorities string
    If Len(priorities) > Len(Replace(priorities, parkCode, "")) Then
    
        'prepare the Park Priority values
        ParkPriorities = Split(priorities, "|")
        
        'set park priority values
        For i = 0 To UBound(ParkPriorities)
            'does Park have a priority value?
            If parkCode = Left(ParkPriorities(i), 4) Then
                PopulateSpeciesPriorities = Replace(ParkPriorities(i), parkCode + "-", "")
            End If
        Next
        
    Else
        'not listed
        PopulateSpeciesPriorities = "X"
    
    End If
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateSpeciesPriorities[mod_Species])"
    End Select
    Resume Exit_Function
End Function