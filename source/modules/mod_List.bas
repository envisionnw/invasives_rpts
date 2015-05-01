Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_List
' Level:        Framework module
' Version:      1.00
' Description:  Listview & listbox related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - initial version
' =================================

' =================================
' SUB:          lvwPopulateFromQuery
' Description:  populates listview control from query
' Parameters:   ctrl - listview control
'               strSQL - SQL statement to run for populating listview
'               aryHeadings - heading array for populating values
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from post comment galura.jayar, 4/26/2012
'               http://www.access-programmers.co.uk/forums/showthread.php?t=225070
'               Created 12/10/2014 blc; Last modified 12/10/2014 blc.
' Revisions:    Bonnie Campbell, Dec 10, 2014 - initial version
'               ListView requires Windows Common Control 6.0 (MSCOMCTRL.OCX from c:\windows\system32)
'                   http://support2.microsoft.com/default.aspx?scid=kb;en-us;194784
'                   http://forums.esri.com/Thread.asp?c=93&f=992&t=198775
'               BLC, 4/30/2015 - added error handling & moved from mod_Common_UI to mod_List
' =================================
Public Sub lvwPopulateFromQuery(ctrl As MSComctlLib.ListView, strSQL As String, aryFields As Variant)
On Error GoTo Err_Handler
    Dim dbs As Database
    Dim rs As Recordset
    Dim item As ListItem
    Dim i As Integer
    
    On Error Resume Next
    
    ctrl.ListItems.Clear

    Set dbs = CurrentDb
    Set rs = dbs.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            Set item = ctrl.ListItems.Add(, , rs(aryFields(i)))
            For i = 1 To UBound(aryFields)
              item.SubItems(i) = rs(aryFields(i))
            Next
            On Error Resume Next 'continue even in error
            rs.MoveNext
            Set item = Nothing
        Loop
    End If

    Set rs = Nothing

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lvwPopulateFromQuery[mod_List])"
    End Select
    Resume Exit_Procedure
End Sub


' =================================
' SUB:          lbxConditionalColor
' Description:  sets lbx text fore color
' Parameters:   ctrl - listbox control
'               tgtCol - column that determines which row(s) fore color should be set to altColor
'               normVal - determining column value for tgtCol  (if tgtCol = normVal then color is set to normColor)
'               altVal - alternate column value for tgtCol (if tgtCol = altVal then color is set to altColor)
'               normColor - string representation of normal listbox row text fore color (vbBlack, vbBlue...)
'               altColor - string representation of color to change listbox row text fore color (vbBlue, vbRed...)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from post comment, 8/2005
'               http://www.tek-tips.com/faqs.cfm?fid=6027
'               Created 12/9/2014 blc; Last modified 12/9/2014 blc.
' Revisions:    Bonnie Campbell, Dec 9, 2014 - initial version
'               ListItem requires Windows Common Control 6.0
'                   http://support2.microsoft.com/default.aspx?scid=kb;en-us;194784
'                   http://forums.esri.com/Thread.asp?c=93&f=992&t=198775
'               BLC, 4/30/2015 - added error handling & moved from mod_Common_UI to mod_List
' =================================
Public Sub lbxConditionalColor(ctrl As ListBox, tgtCol As Integer, normVal As String, altVal As String, normColor As Long, altColor As Long)
On Error GoTo Err_Handler
    Dim counter As Long
    Dim col As Integer
    
    For counter = 0 To ctrl.ListCount - 1
        With ctrl
            If CStr(.Column(tgtCol, counter)) = normVal Then
                For col = 0 To .ColumnCount - 1
                    .Column(col, counter).ForeColor = normColor
                Next col
            ElseIf CStr(.Column(tgtCol, counter)) = altVal Then
                For col = 0 To .ColumnCount - 1
                    .Column(col, counter).ForeColor = altColor
                Next col
            End If
        End With
    Next counter
    
    'ctrl.refresh

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxConditionalColor[mod_List])"
    End Select
    Resume Exit_Procedure
End Sub