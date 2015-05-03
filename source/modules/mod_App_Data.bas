Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_Data
' Level:        Application module
' Version:      1.00
' Description:  data functions & procedures
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC - 2/9/2015 - initial version
' =================================

' ---------------------------------
' SUB:          fillList
' Description:  Fill a list (or listbox like subform) from specific queries for datasheets, species or other items
' Assumptions:  Either a listbox or subform control is being populated
' Parameters:   frm - main form object
'               ctrl - either:
'                      lbx - main form listbox object (for filling a listbox control)
'                      sfrm - subform object (for populating a subform control)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/6/2015  - initial version
'   BLC, 2/18/2015 - adapted to include subform as well as listbox controls
'   BLC, 5/1/2015  - integrated into Invasives Reporting tool
' ---------------------------------
Public Sub fillList(frm As Form, ctrlSource As Control, Optional ctrlDest As Control)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strQuery As String, strSQL As String
    
    'output to form or listbox control?
   
    'determine data source
    Select Case ctrlSource.name
    
        Case "lbxDataSheets", "sfrmDatasheets" 'Datasheets
            strQuery = "qry_Active_Datasheets"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
            
        Case "lbxSpecies", "lbxTgtSpecies", "fsub_Species_Listbox" 'Species
            strQuery = "qry_Plant_Species"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
            
    End Select

    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)

    'set TempVars
    TempVars.Add "strSQL", strSQL

    If Not ctrlDest Is Nothing Then
        'populate list & headers
        PopulateList ctrlSource, rs, ctrlDest
    Else
        'populate only ctrlSource headers
        PopulateListHeaders ctrlSource, rs
    End If
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fillList[mod_App_Data])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     getParkState
' Description:  Retrieve the state associated with a park (via tlu_Parks)
' Assumptions:  Park state is properly identified in tlu_Parks
' Parameters:   parkCode - 4 character park designator
' Returns:      ParkState - 2 character state abbreviation
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015  - initial version
' ---------------------------------
Public Function getParkState(parkCode As String) As String

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim state As String, strSQL As String
   
    'handle only appropriate park codes
    If Len(parkCode) <> 4 Then
        GoTo Exit_Function
    End If
    
    'generate SQL ==> NOTE: LIMIT 1; syntax not viable for Access, use SELECT TOP x instead
    strSQL = "SELECT TOP 1 ParkState FROM tlu_Parks WHERE ParkCode LIKE '" & parkCode & "';"
            
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        state = rs.Fields("ParkState").Value
    End If
   
    'return value
    getParkState = state
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getParkState[mod_App_Data])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     MergeRecordsets
' Description:  Merge two recordsets into one (useful when the recordsets already exist vs. direct SQL union)
' Assumptions:  Recordsets have the same fields in the same order
' Parameters:   rsA - DAO recordset A
'               rsB - DAO recordset B to merge with A
' Returns:      DAO.Recordset
' Throws:       none
' References:   none
' Source/date:
' Chris Oswald, January 26, 2011
' http://www.mrexcel.com/forum/excel-questions/524214-visual-basic-applications-joining-multiple-recordets-multiple-databases.html
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015  - initial version
' ---------------------------------
Public Function MergeRecordsets(rsA As DAO.Recordset, rsB As DAO.Recordset) As DAO.Recordset

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rsOut As DAO.Recordset
    Dim iCount As Integer
    
    'handle empty recordsets
    If rsA Is Nothing Then
        'check rsB
        If rsB Is Nothing Then
            GoTo Exit_Function
        Else
            Set MergeRecordsets = rsB
            GoTo Exit_Function
        End If
    End If
    

'With rsA
    'check if rsA and rsB are both populated --> if not, exit
    If (rsA.EOF And rsA.BOF) Then
        'rsA not populated
        If (rsB.EOF And rsB.BOF) Then
            'neither is populated --> EXIT!
            GoTo Exit_Function
        Else
            'rsB populated --> return rsB
            Set MergeRecordsets = rsB
        End If
    Else
        'rsA populated --> if rsB not populated, return rsA
        If (rsB.EOF And rsB.BOF) Then Set MergeRecordsets = rsA
    End If
    
    'create output recordset vs. just adding to rsB
    Set rsOut = rsB
    
    'iterate through recordset
    'rsA.MoveFirst
    Do Until rsA.EOF
        'add rsA values as new rsOut records
        rsOut.AddNew
        For iCount = 0 To rsA.Fields.count - 1
            rsOut.Fields(iCount).Value = rsA.Fields(iCount).Value
        Next
        rsOut.Update
        rsA.MoveNext
    Loop
'End With

    Set MergeRecordsets = rsOut

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MergeRecordsets[mod_App_Data])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     DisconnectRecordset
' Description:  Create a disconnected ADO (in-memory) recordset for manipulation
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   ADO required
' Source/date:
' Danny Lesandrini, November 2, 2009
' http://www.databasejournal.com/features/msaccess/article.php/3846361/Create-In-Memory-ADO-Recordsets.htm
' Fionnuala, October 12, 2011
' http://stackoverflow.com/questions/7738811/access-listbox-based-on-value-list-sorting-on-column
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Public Function DisconnectRecordset(valueList As Variant) As Variant

On Error GoTo Err_Handler
    
'Dim rs As New adodb.Recordset
Dim rs As Recordset

'slist = "0,Standard price,1650," _
'& "14,Bookings made during Oct 2011,3770," _
'& "15,Minimum Stay 4 Nights - Special Price,2460"

With rs
 ' .ActiveConnection = Nothing
 ' .CursorLocation = adUseClient
 ' .CursorType = adOpenStatic
 ' .LockType = adLockBatchOptimistic
 ' With .Fields
 '   .Append "Field1", adInteger
 '   .Append "Field2", adVarChar, 200
 '   .Append "Field3", adInteger
 ' End With
 ' .Open

Dim Ary As Variant
Dim j As Integer, i As Integer

  Ary = Split(valueList, ",")

  For j = 0 To UBound(Ary)
      .AddNew
      For i = 0 To 2
  '        .Fields(i).Value = Ary(j)
          j = j + 1
      Next
      j = j - 1

  Next

  '.Sort = "Field3"
End With

'slist = rs.GetString(, , ",", ",")
'slist = Left(slist, Len(slist) - 1)

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CountArrayValues[mod_Lists])"
    End Select
    Resume Exit_Function
End Function

Public Function CreateDisconnectedRecordset()
Dim rstADO As ADODB.Recordset
Dim fld As ADODB.Field

Set rstADO = New ADODB.Recordset
With rstADO
    .Fields.Append "EmployeeID", adInteger, , adFldKeyColumn
    .Fields.Append "FirstName", adVarChar, 10, adFldMayBeNull
    .Fields.Append "LastName", adVarChar, 20, adFldMayBeNull
    .Fields.Append "Email", adVarChar, 64, adFldMayBeNull
    .Fields.Append "Include", adInteger, , adFldMayBeNull
    .Fields.Append "Selected", adBoolean, , adFldMayBeNull

    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .LockType = adLockPessimistic
    .Open
End With
End Function