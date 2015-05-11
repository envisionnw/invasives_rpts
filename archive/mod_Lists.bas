Attribute VB_Name = "mod_Lists"
Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Lists
' Description:  listbox functions & procedures
'
' Source/date:  Bonnie Campbell, 1/30/2015
' Revisions:    BLC - 1/30/2015 - initial version
' =================================

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
' SUB:          MoveSingleItem
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 3/5/2015 - added ability to remove from list w/o adding to target if strSourceControl = strTargetControl
' ---------------------------------
Public Sub MoveSingleItem(frm As Form, strSourceControl As String, strTargetControl As String)
    
On Error GoTo Err_Handler
    
    Dim stritem As String
    Dim intColumnCount As Integer
    
    'if source = target, just remove the item
    If strSourceControl = strTargetControl Then
        RemoveSelectedItems frm.Controls(strSourceControl)
        GoTo Exit_Sub
    End If
    
    'check for control type
    If frm.Controls(strSourceControl).ControlType = acSubform Then MsgBox frm.Controls(strSourceControl).ControlType, vbOKOnly, "ctrltype"
    
    'check for at *least* one selected item
    If frm.Controls(strSourceControl).ItemsSelected.count = 0 Then
        MsgBox "Please select at least one item.", vbExclamation, "Oops!"
        GoTo Exit_Sub
    End If
    
    If frm.Controls(strSourceControl).ItemsSelected.count > 1 Then
        MoveSelectedItems frm, strSourceControl, strTargetControl
        GoTo Exit_Sub
    End If
    
    For intColumnCount = 0 To frm.Controls(strSourceControl).ColumnCount - 1
        stritem = stritem & frm.Controls(strSourceControl).Column(intColumnCount) & ";"
    Next
    
    'remove extra semi-colon (;)
    stritem = left(stritem, Len(stritem) - 1)

    'Check the length to make sure something is selected
    ' -------------------------------------------------------------------------
    '  NOTE: ListIndex is zero based, so add 1 to remove proper item
    ' -------------------------------------------------------------------------
    If Len(stritem) > 0 Then
        frm.Controls(strTargetControl).AddItem stritem
        frm.Controls(strSourceControl).RemoveItem frm.Controls(strSourceControl).ListIndex + 1
    Else
        MsgBox "Please select an item to move."
    End If


Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MoveSingleItem[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          MoveAllItems
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 3/5/2015 - added ability to remove from list w/o adding to target if strSourceControl = strTargetControl
' ---------------------------------
Public Sub MoveAllItems(frm As Form, strSourceControl As String, strTargetControl As String)
    
On Error GoTo Err_Handler
    
    Dim stritem As String
    Dim intColumnCount As Integer, startRow As Integer
    Dim lngRowCount As Long
    
    'if source = target, just remove the items
    If strSourceControl = strTargetControl Then
        RemoveSelectedItems (frm.Controls(strSourceControl))
        GoTo Exit_Sub
    End If
        
    'check for at *least* one item
    If frm.Controls(strSourceControl).ListCount = 0 Then
        MsgBox "Your list needs at least one item to move.", vbExclamation, "Oops!"
        GoTo Exit_Sub
    End If
    
    startRow = 0 'default
    'set start row
    If frm.Controls(strSourceControl).ColumnHeads = True Then
        startRow = 1
    End If
    
    For lngRowCount = startRow To frm.Controls(strSourceControl).ListCount - 1
        For intColumnCount = 0 To frm.Controls(strSourceControl).ColumnCount - 1
            stritem = stritem & frm.Controls(strSourceControl).Column(intColumnCount, lngRowCount) & ";"
        Next
        stritem = left(stritem, Len(stritem) - 1)
        frm.Controls(strTargetControl).AddItem stritem
        stritem = ""
    Next
        
    'clear the list
    frm.Controls(strSourceControl).RowSource = ""
    
    'add back the headers
    ' -------------------------------------------------------------------------
    ' NOTE: target lbx will already have headers, so only add back to source
    ' -------------------------------------------------------------------------
    If frm.Controls(strSourceControl).ColumnHeads = True Then
        frm.Controls(strSourceControl).AddItem TempVars("lbxHdr")
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MoveAllItems[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          MoveSelectedItems
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' ManningFan, January 30,2015
' http://bytes.com/topic/access/answers/765291-populating-1-listbox-another-listbox
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 3/5/2015 - added ability to remove from list w/o adding to target if strSourceControl = strTargetControl
' ---------------------------------
Public Sub MoveSelectedItems(frm As Form, strSourceControl As String, strTargetControl As String)
    
On Error GoTo Err_Handler
    
    Dim iRow As Integer, startRow As Integer, i As Integer, x As Integer, iRemovedItems As Integer
    Dim arySelectedItems() As Integer
    Dim blnDimensioned As Boolean
    Dim stritem As String
    
    'if source = target, just remove the items
    If strSourceControl = strTargetControl Then
        RemoveSelectedItems (frm.Controls(strSourceControl))
        GoTo Exit_Sub
    End If
    
    'check for at *least* one selected item
    If frm.Controls(strSourceControl).ItemsSelected.count = 0 Then
        MsgBox "Please select at least one item.", vbExclamation, "Oops!"
        GoTo Exit_Sub
    End If
    
    startRow = 0 'default
    'set start row
    If frm.Controls(strSourceControl).ColumnHeads = True Then
        startRow = 1
    End If
    
    'add back the header if it doesn't exist
    If frm.Controls(strTargetControl).ColumnHeads = True And frm.Controls(strTargetControl).ListCount = 0 Then
       stritem = TempVars.item("lbxHdr") & stritem
       frm.Controls(strTargetControl).AddItem stritem
    End If
    
    'generate array of selected items
    For iRow = startRow To frm.Controls(strSourceControl).ListCount - 1
    
        'fetch array of selected items
        '--------------------------------------------------
        ' if > 1 item selected, other selected items
        ' deselected when first source item removed
        '--------------------------------------------------
        If frm.Controls(strSourceControl).Selected(iRow) Then
            
            'Array dimensioned?
            If blnDimensioned = True Then
                      
                'Yes ==> extend array 1 element largee than current upper bound
                '        w/o "Preserve" keyword previous elements erased w/ resizing
                ReDim Preserve arySelectedItems(0 To UBound(arySelectedItems) + 1) As Integer
                      
            Else
                      
                'No ==> dimension it and flag as dimensioned
                ReDim arySelectedItems(0 To 0) As Integer
                blnDimensioned = True
                          
            End If
                  
            'Add to last element in the array.
            arySelectedItems(UBound(arySelectedItems)) = iRow
        End If
    
    Next
    
    'set default
    iRemovedItems = 0
    
    'iterate through selected items
    For x = LBound(arySelectedItems) To UBound(arySelectedItems)
                        
        iRow = arySelectedItems(x) - iRemovedItems
            
        'clear string
        stritem = ""
        
        'add all columns
        For i = 0 To frm.Controls(strSourceControl).ColumnCount
            stritem = stritem & frm.Controls(strSourceControl).Column(i, iRow) & ";"
        Next i
        
        'add to target
        frm.Controls(strTargetControl).AddItem stritem
        
        'remove from source
        frm.Controls(strSourceControl).RemoveItem iRow
            
        'adjust list after removal
        If UBound(arySelectedItems) > 0 Then
            iRemovedItems = iRemovedItems + 1
        End If
    
    Next x

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MoveSelectedItems[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          PopulateListHeaders
' Description:  Populate the headers for listbox controls
' Assumptions:  headers are the same as recordset field names
'               sfrms acting as listboxes have static headers already present
' Parameters:   ctrl - listbox control
'               rs   - recordset containing list headers
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015  - initial version
'   BLC - 2/19/2015 - converted to generic to handle listbox-like controls & documentation update
' ---------------------------------
Public Sub PopulateListHeaders(ctrl As Control, rs As Recordset)

On Error GoTo Err_Handler

    Dim rows As Integer, cols As Integer, i As Integer, j As Integer, matches As Integer
    Dim frm As Form
    Dim stritem As String, strColHeads As String, aryColWidths() As String

    'exit if subform control (hdrs are static & present on sfrm)
    If ctrl.ControlType = 112 Then
        GoTo Exit_Sub
    End If

    Set frm = ctrl.Parent
    
    rows = rs.RecordCount
    cols = rs.Fields.count
    
    If Nz(rows, 0) = 0 Then
        MsgBox "Sorry, no records found..."
        GoTo Exit_Sub
    End If
    
    'fetch column widths
    aryColWidths = Split(ctrl.ColumnWidths, ";")
    
    'populate column names (if desired)
    If ctrl.ColumnHeads = True Then
        strColHeads = ""
        For i = 0 To cols - 1
            If CInt(aryColWidths(i)) > 0 Then
                strColHeads = strColHeads & rs.Fields(i).name & ";"
            End If
        Next i
        ctrl.AddItem strColHeads
    End If

    'save headers
    TempVars.Add "lbxHdr", strColHeads

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateListHeaders[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          PopulateList
' Description:  Populate listbox and similar controls from recordset
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' krish KM, Aug. 27, 2014
' http://stackoverflow.com/questions/25526904/populate-listbox-using-ado-recordset
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Public Sub PopulateList(ctrlSource As Control, rs As Recordset, ctrlDest As Control)

On Error GoTo Err_Handler

    Dim frm As Form
    Dim rows As Integer, cols As Integer, i As Integer, j As Integer, matches As Integer, iZeroes As Integer
    Dim stritem As String, strColHeads As String, aryColWidths() As String

    Set frm = ctrlSource.Parent
    
    rows = rs.RecordCount
    cols = rs.Fields.count
    
    'address no records
    If Nz(rows, 0) = 0 Then
        MsgBox "Sorry, no records found..."
        GoTo Exit_Sub
    End If
    
    'handle sfrm controls (acSubform = 112)
    If ctrlSource.ControlType = acSubform Then
        Set ctrlSource.Form.Recordset = rs
        
        ctrlSource.Form.Controls("tbxCode").ControlSource = "Code"
        ctrlSource.Form.Controls("tbxSpecies").ControlSource = "Species"
        ctrlSource.Form.Controls("tbxMasterCode").ControlSource = "Master_PLANT_Code"
        
        'set the initial record count (MoveLast to get full count, MoveFirst to set display to first)
        rs.MoveLast
        ctrlSource.Parent.Form.Controls("lblSfrmSpeciesCount").Caption = rs.RecordCount & " species"
        rs.MoveFirst
        
        GoTo Exit_Sub
    End If
    
    'fetch column widths array
    aryColWidths = Split(ctrlSource.ColumnWidths, ";")
    
    'count number of 0 width elements
    iZeroes = CountArrayValues(aryColWidths, "0")
        
    'clear out existing values
    ClearList ctrlSource
    
    'populate column names (if desired)
    If ctrlSource.ColumnHeads = True Then
        PopulateListHeaders ctrlSource, rs
        
        'populate second listbox headers if present
        If ctrlDest.ColumnHeads = True Then
            ClearList ctrlDest
            PopulateListHeaders ctrlDest, rs
        End If
    End If
    
    'populate data
    Select Case ctrlSource.RowSourceType
        Case "Table/Query"
            Set ctrlSource.Recordset = rs
        Case "Value List"
            
            'initialize
            i = 0
            
            Do Until rs.EOF
            
                'initialize item
                stritem = ""
                    
                'generate item
                For j = 0 To cols - 1
                    'check if column is displayed width > 0
                    If CInt(aryColWidths(j)) > 0 Then
                    
                        stritem = stritem & rs.Fields(j).Value & ";"
                    
                        'determine how many separators there are (";") --> should equal # cols
                        matches = (Len(stritem) - Len(Replace$(stritem, ";", ""))) / Len(";")
                        
                        'add item if not already in list --> # of ; should equal cols - 1
                        'but # in list should only be # of non-zero columns --> cols - iZeroes
                        If matches = cols - iZeroes Then
                            ctrlSource.AddItem stritem
                            'reset the string
                            stritem = ""
                        End If
                    
                    End If
                
                Next
                
                i = i + 1
                
                rs.MoveNext
            Loop
        Case "Field List"
    End Select

     'MsgBox ctrlSource.ListCount & " in list" & vbCrLf & rs.RecordCount & " in rs", vbOKOnly, "Num in list"
    'refresh control
    'lbx.Requery

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateList[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          RemoveListDupes
' Description:  Remove listbox duplicate values
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' matsushita, September 27, 2006
' https://social.msdn.microsoft.com/Forums/vstudio/en-US/0799668c-36dd-42d9-9599-3085a6c0581f/how-to-remove-duplicate-values-in-listbox-
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
' ---------------------------------
Public Sub RemoveListDupes(lbx As ListBox)

On Error GoTo Err_Handler

    Dim index As Integer, count As Integer
    Dim lastItem As String
    
    'sort listbox
    SortList lbx
    
    count = lbx.ListCount

    'check sorted listbox for duplicates & remove
    If count > 1 Then
    
        lastItem = lbx.ItemData(count - 1)

        For index = count - 2 To 0 Step -1
            If lbx.ItemData(index) = lastItem And Len(lbx.ItemData(index)) > 0 Then
                'duplicate
                lbx.RemoveItem (index)
            Else
                lastItem = lbx.ItemData(index)
            End If
        Next
    End If

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateList[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          SortList
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' MajP, March 22, 2012
' http://www.tek-tips.com/viewthread.cfm?qid=1677888
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
' ---------------------------------
Public Sub SortList(lbx As ListBox) ', orderCol As Integer)

On Error GoTo Err_Handler
  
  Dim strTemp As String
  Dim i As Integer
  Dim j As Integer
  
  For i = 0 To lbx.ListCount - 1
    For j = i + 1 To lbx.ListCount - 1
      If lbx.ItemData(i) > lbx.ItemData(j) Then
        strTemp = lbx.ItemData(i)
        lbx.RemoveItem (i)
        lbx.AddItem lbx.ItemData(j - 1), i
        lbx.RemoveItem (j)
        lbx.AddItem strTemp, j - 1
       End If
     Next j
   Next i

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortList[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          RemoveSelectedItems
' Description:  Removes selected items from a listbox by re-creating rowsource
' Assumptions:  lbx is a listbox control (not a continuous subform which may act as a listbox control)
' Parameters:   lbx - Listbox to remove selected items from
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' ADezii, April 13, 2010
' http://bytes.com/topic/access/answers/885569-remove-selected-items-list-box-microsoft-access
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
' ---------------------------------
Public Sub RemoveSelectedItems(lbx As ListBox)
On Error GoTo Err_Handler
  
    Dim intRow As Integer, iCol As Integer
    Dim strBuild As String
     
    With lbx
      If .ItemsSelected.count = 0 Then Exit Sub
     
      For intRow = 0 To .ListCount - 1
        If Not .Selected(intRow) Then
            For iCol = 0 To .ColumnCount
                strBuild = strBuild & .Column(iCol, intRow) & ";"
            Next
        End If
      Next
     
      strBuild = left$(strBuild, Len(strBuild) - 1)
     
      .RowSource = strBuild
    End With

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortList[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     IsListDuplicate
' Description:  Check if item is already on the list
' Assumptions:  -
' Parameters:   lbx - listbox control to check (listbox object)
'               col - column which would hold the item being checked (integer)
'               item - name of item to be checked (string)
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Public Function IsListDuplicate(lbx As ListBox, col As Integer, item As String) As Boolean
On Error GoTo Err_Handler
    
    Dim isDupe As Boolean
    Dim i As Integer
    
    'set default
    isDupe = False
    
    'iterate through listbox (use .Column(col,i) vs .ListIndex(i) which results in error 451 property let not defined, property get...)
    For i = 0 To lbx.ListCount
        'check if item exists in listbox
        If lbx.Column(col, i) = item Then
            'duplicate, so exit
            isDupe = True
            GoTo Exit_Function
        End If
    Next

Exit_Function:
    IsListDuplicate = isDupe
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsListDuplicate[mod_Lists])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:          ClearList
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Public Sub ClearList(lbx As ListBox)

On Error GoTo Err_Handler

    'clear listbox items
    lbx.RowSource = ""

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearList[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          DisableControl
' Description:  Set color scheme for labels so they appear disabled
' Assumptions:  Assumes control has BackColor and ForeColor properties
' Parameters:   ctrl - control to set color scheme for
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Public Sub DisableControl(ctrl As Control)

On Error GoTo Err_Handler
    
    ctrl.backcolor = lngLtGray
    ctrl.forecolor = lngGray
    
    If ctrl.ControlType = acCommandButton Then
        ctrl.borderColor = lngGray
    End If

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableControl[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          EnableControl
' Description:  Set color scheme for labels so they appear enabled
' Assumptions:  Assumes control has BackColor and ForeColor properties
' Parameters:   ctrl - control to set color scheme for
'               backColor - long value for desired back color
'               foreColor - long value for desired fore (text) color
'               optionally for command buttons:
'               borderColor - long value for desired border color
'               hoverColor - long value for desired hover color
'               pressedColor - long value for desired pressed button color
'               hoverForeColor - long value for desired hover fore (text) color
'               pressedForeColor - long value for desired pressed button fore (text) color
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Public Function EnableControl(ctrl As Control, backcolor As Long, forecolor As Long, _
                                Optional borderColor As Long, _
                                Optional hoverColor As Long, _
                                Optional pressColor As Long, _
                                Optional hoverForeColor As Long, _
                                Optional pressedForeColor As Long)
On Error GoTo Err_Handler
    
    ctrl.backcolor = backcolor
    ctrl.forecolor = forecolor
    
    If ctrl.ControlType = acCommandButton Then
        ctrl.borderColor = borderColor
        ctrl.hoverColor = hoverColor
        ctrl.pressedColor = pressColor
        ctrl.hoverForeColor = hoverForeColor
        ctrl.pressedForeColor = pressedForeColor
    End If

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableControl[mod_Lists])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:          SaveListToTable
' Description:  Save list items to table
' Assumptions:  Assumes control has BackColor and ForeColor properties
' Parameters:   ctrl - control to iterate through
'               tbl - table being
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 8, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/8/2015 - initial version
' ---------------------------------
Public Sub SaveListToTable(ctrl As Control, tbl As String, tblFields As Variant, blnSelectedOnly As Boolean)

On Error GoTo Err_Handler
    
    Dim strSQL As String, strFields As String
    Dim i As Integer, iRow As Integer, jCol As Integer
    
    strSQL = "INSERT INTO " & tbl & " " & tblFields & "VALUES ("
    
    ' prepare fields
    strFields = ""
    For i = 0 To UBound(tblFields)
    
        Select Case tblFields(1, i)
            Case "Integer"
            Case "VarChar"
        End Select
        strFields = strFields
    
    Next

    'iterate through items
    For iRow = 0 To ctrl.ListCount - 1
    
            For jCol = 0 To ctrl.ColumnCount - 1
            
            strSQL = strSQL & "'" & ctrl.Column(jCol, iRow) & "'"
             
            CurrentDb.Execute strSQL, dbFailOnError
            
            Next
    Next 'iRow

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableControl[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     CountArrayValues
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Public Function CountArrayValues(Ary As Variant, val As Variant) As Integer

On Error GoTo Err_Handler
    
    Dim i As Integer, numItems As Integer

    'default
    numItems = 0
    
    If IsArray(Ary) Then
    
        For i = LBound(Ary) To UBound(Ary)
            If Ary(i) = val Then
                numItems = numItems + 1
            End If
        Next
        
    End If
    
    CountArrayValues = numItems

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
