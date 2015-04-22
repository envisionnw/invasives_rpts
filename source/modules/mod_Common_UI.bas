Option Compare Database
Option Explicit
' =================================
' MODULE:       mod_Common_UI
' Description:  Standard module with User Interface (UI) related functions & subroutines
'
' Source/date:  Bonnie Campbell, May 2014
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 12/9/2014 - added lbxConditionalColor to handle listbox conditional formatting
'               --------------------------------------------------------------------------------------
'               BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
' =================================

' =================================
' SUB:     PopulateInstructions
' Description:  Sets form instruction strings
' Assumptions:  -
' Parameters:   strTab - tab for instruction string
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - XX
'               --------------------------------------------------------------------------------------
'               BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
'                                Converted QAQC to Create, Logs to View
' =================================
Public Sub PopulateInstructions(ctrl As Control, strContext As String)
    Dim strInstructions As String
    
    'MsgBox strContext
    
    Select Case strContext
        Case "Create" ' Create main
            strInstructions = "Choose the analysis you would like to run."
        Case "Outliers", "MissingData", "SuspectValues", "SuspectDO", "SuspectpH", "SuspectSC", "SuspectWT", "Duplicates" ' QA/QC main
            strInstructions = "Complete the fields to define the data set or subset you are validating. " _
                    & "Leave the fields blank if you are validating all data. Click 'Run' to validate."
        Case "View" ' View main
            strInstructions = "Log your modifications to data within the edit log. " _
                    & "Be as complete as possible to aid others in tracing data changes."
        Case "Reports" ' Reports main
            strInstructions = "Choose the report you would like to run."
        Case "Precision", "Effectiveness", "Bias", "Stage", "Flow" ' Reports > Precision etc.
            strInstructions = "Complete the fields to define the data set or subset you are reporting. " _
                    & "Leave the fields blank if you are reporting on all data. Click 'Run' to validate."
        Case "Export" ' Export main
            strInstructions = "Choose the export you would like to run."
        Case "UtahLab" ' Exports > Utah Lab etc.
            strInstructions = "Choose the export you would like to run."
        Case "DbAdmin" ' DB Admin main
            strInstructions = ""
    End Select
    
    'populate caption & display instructions
    If ctrl.ControlType = acLabel Then
        ctrl.Caption = strInstructions
        If strContext <> "DbAdmin" Then
            ctrl.Visible = True
        End If
    End If
    
End Sub

' =================================
' SUB:     PopulateInsetTitle
' Description:  Sets inset title on form
' Assumptions:
' Parameters:   frm - form holding crumb labels
'               aryCrumbs - breadcrumb array
'               separator - non-clickable value between crumbs, default = >
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - XX
'               --------------------------------------------------------------------------------------
'               BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
'                                Converted QAQC to Create, Logs to View
' =================================
Public Sub PopulateInsetTitle(ctrl As Control, strContext As String)
    Dim strTitle As String
    
    Select Case strContext
        Case "Create" ' Create main
            strTitle = "Create Species Target Lists"
        Case "Outliers", "MissingData", "SuspectValues", "SuspectDO", "SuspectpH", "SuspectSC", "SuspectWT", "Duplicates"  ' QA/QC > Outliers etc.
            strContext = Replace(Replace(strContext, "Suspect", "Suspect "), "Missing", "Missing ")
            strTitle = "Data Validation > " & strContext
        Case "Data Validation" ' QA/QC analysis project selection
            strTitle = "Data Validation > Field > Duplicates (NFV)" '<<<<< Make this so it ties back to the selected analysis
        Case "View" ' View main
            strTitle = "Data Modifications"
        Case "Reports" ' Reports main
            strTitle = "Reports"
        Case "Precision", "Effectiveness", "Bias", "Stage", "Flow" ' Reports > Precision etc.
            strTitle = "Reports > " & strContext
        Case "Export" ' Export main
            strTitle = "Export"
        Case "UtahLab" ' Exports > Utah Lab etc.
            strContext = Replace(strContext, "Lab", " Lab")
            strTitle = "Exports > " & strContext
        Case "DB Admin" ' DB Admin main
            strTitle = ""
    End Select
    
    If ctrl.ControlType = acLabel Then
        ctrl.Caption = strTitle
        If strContext <> "DbAdmin" Then
            ctrl.Visible = True
        End If
    End If
End Sub

' ---------------------------------
' SUB:          buttonHighlight
' Description:  Toggle button color to strColor or transparent if already colored
' Parameters:   btn      - name of the button to change
'                          accommodates command and label as control buttons
'               strColor - color as a string (hex)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell May 12, 2014 - NCPN tools
' Adapted:      -
' Revisions:    Bonnie Campbell, May 12, 2014 - XX
' ---------------------------------
Public Sub buttonHighlight(btn As Control, Optional solo As Boolean, Optional Toggle As Boolean, Optional intEffect As Integer, Optional strColor As String)
' Special Effects:  0 - flat, 1 - raised, 2 - sunken, 3 - etched, 4 - shadowed, 5 - chiseled
' Colors:
'   lime                   #9EFF00
'   chartreuse 1           #7FFF00 127 255 00  65407
'   dark olive green 1     #CAFF70 202 255 112 7405514
'   mint                   #BDFCC9 189 252 201 13237437
'   light lime (like)      #E6FABF 230 250 191
'   darker lt lime         #CFF583 207 245 131
On Error GoTo Err_Handler:

    'toggle button
    If Toggle Then
        buttonUnHighlight btn, Toggle
    End If
    
    'change all others to transparent if solo
    If solo Then
        buttonUnHighlight btn
    End If
    
    With btn
        If .BackStyle = 1 Then
            GoTo Transparent
        End If
        
        If (Len(strColor) <> 6) Then
            strColor = "CFF583"
        End If
    
        If intEffect > -1 Or intEffect > 6 Then
            intEffect = 0 'flat
        End If
           
        'change button background to given color
        .BackStyle = 1 'Normal - required to change color
        .BackColor = fxnHTMLConvert("#" & strColor)
        .SpecialEffect = intEffect
    End With
    
Exit_Procedure:
    Exit Sub

Transparent:
    btn.BackStyle = 0 'Transparent
    GoTo Exit_Procedure

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - buttonHighlight[mod_Common_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          buttonUnHighlight
' Description:  Toggles all other buttons to transparent if already colored
' Parameters:   btn - name of the button control to change
'                     accommodates command and label as control buttons
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell May 12, 2014 - NCPN tools
' Adapted:      -
' Revisions:    Bonnie Campbell, May 12, 2014 - initial version
' ---------------------------------
Public Sub buttonUnHighlight(btn As Control, Optional blnToggle As Boolean)
On Error GoTo Err_Handler:
Dim ctl As Control

    With btn
        'unhighlight only btn
        If blnToggle Then
            .BackStyle = 0 'transparent
            .SpecialEffect = 0 'flat
            GoTo Exit_Procedure
        End If
        
        'unhighlight all other buttons
        For Each ctl In .Parent.Controls

            If ctl.name <> btn.name And _
                ctl.ControlType = acLabel Then
                With ctl
                    .BackStyle = 0 'transparent
                End With
            End If

        Next
    
    End With
    
Exit_Procedure:
    'update display
    repaintParentForm btn
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - buttonUnHighlight[mod_Common_UI])"
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
' =================================
Public Sub lbxConditionalColor(ctrl As ListBox, tgtCol As Integer, normVal As String, altVal As String, normColor As Long, altColor As Long)
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
End Sub

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
' =================================
Public Sub lvwPopulateFromQuery(ctrl As MSComctlLib.ListView, strSQL As String, aryFields As Variant)
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
End Sub

' ---------------------------------
' SUB:          getAccessObjectType
' Description:  looks up object type in Access sys tables
' Parameters:   strName  - name of object w/in Access
' Returns:      long (type) or NULL if object doesn't exist
'                   ----------------
'                   1 = Access Table
'                   4 = OBDB-Linked Table / View
'                   5 = Access Query
'                   6 = Attached (Linked) File  (such as Excel, another Access Table or query, text file, etc.)
'                   -32768 = Access Form
'                   -32764 = Access Report
'                   -32761 = Access Module
'                   ----------------
' Throws:       none
' References:   Tom Davidson, April 8, 2011
'   http://stackoverflow.com/questions/2090578/ms-access-determine-object-type

' Source/date:  Bonnie Campbell August 20, 2014 - NCPN tools
' Adapted:      -
' Revisions:    Bonnie Campbell, August 20, 2014 - XX
' ---------------------------------
Public Function getAccessObjectType(strObject As String)
On Error GoTo Err_Handler:

    getAccessObjectType = DLookup("Type", "MSysObjects", "NAME = '" & strObject & "'")
   
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getAccessObjectType[mod_Common_UI])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnHTMLConvert
' Description:  converts HTML string value for color to RGB which can be used for control colors
' Parameters:   strHTML - HTML color (make sure you include # otherwise the color won't match)
' Returns:      HTML color as long
' Throws:       none
' References:   none
' Source/date:  Adapted from http://www.access-programmers.co.uk/forums/showthread.php?t=193353
'               by Steve R., 5/21/2010.
'               Created 05/12/2014 blc; Last modified 05/12/2014 blc.
' Revisions:    Bonnie Campbell, May 12, 2014 - XX
' =================================
Public Function fxnHTMLConvert(strHTML As String) As Long
    Rem converts a HTML color code number such as #D8B190 to an RGB value.
    fxnHTMLConvert = RGB(CInt("&H" & Mid(strHTML, 2, 2)), CInt("&H" & Mid(strHTML, 4, 2)), CInt("&H" & Mid(strHTML, 6, 2)))
End Function

' ---------------------------------
' SUB:          repaintParentForm
' Description:  Repaints the control's parent(or grandparent or great grandparent...) form
' Parameters:   ctl - control whose parent form you're looking to repaint
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell August, 2014 - NCPN tools
' Adapted:      -
' Revisions:    Bonnie Campbell, August 20, 2014 - initial version
' ---------------------------------
Public Sub repaintParentForm(ctl As Control)
On Error GoTo Err_Handler:
Dim parentControl As Object
        
    Set parentControl = ctl.Parent
    
    Do Until parentControl Is Nothing
      
        If TypeName(parentControl.name) = "String" Then
            'form? -> refresh the display
            If getAccessObjectType(parentControl.name) = -32768 Then
                parentControl.Repaint
                Exit Do
            End If
            Set parentControl = parentControl.Parent
        Else
            'form? -> refresh the display
            If CurrentProject.AllForms(parentControl.name).IsLoaded Then
                parentControl.Repaint
                Exit Do
            End If
        End If
    Loop
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - findParentForm[mod_Common_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' =================================
' SUB:          tabPageUnhide
' Description:  sets desired tab visible, all others hidden
' Parameters:   strTabName - tab page name to make visible
'               ctrl - tab control
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from Tom's post comment, 9/12/2009
'               http://www.vbdotnetforums.com/gui/36561-loop-through-tab-pages-remove.html
'               Created 06/11/2014 blc; Last modified 06/11/2014 blc.
' Revisions:    Bonnie Campbell, June 11, 2014 - XX
' =================================
Public Sub tabPageUnhide(ctrl As TabControl, strTabName As String)
    Dim pg As page
    
    For Each pg In ctrl.Pages
        If pg.name = strTabName Then
            ctrl.Pages(pg.name).Visible = True
        Else
            ctrl.Pages(pg.name).Visible = False
        End If
    Next pg
    
End Sub

' =================================
' FUNCTION:     fxnCrumbsToArray
' Description:  Prepares breadcrumb elements from Me.OpenArgs values
' Parameters:   strCrumbs - Me.OpenArgs values from form open subs
'               delimiter - delimiter used for separating string values, default = | (pipe)
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - XX
' =================================
Public Function fxnCrumbsToArray(strCrumbs As String, Optional delimiter = "|")
    
    Dim strCrumbTrail As String

    If Len(strCrumbs) > 0 Then
        Dim aryCrumbs As Variant
        
        aryCrumbs = Split(strCrumbs, delimiter)
        
    End If

    fxnCrumbsToArray = aryCrumbs
    
End Function

' =================================
' SUB:     PrepareCrumbs
' Description:  Sets breadcrumb label control captions & click events based on crumb element array
' Assumptions:  Breadcrumbs are displayed using label controls (lblCrumb01...)
'               & labels already exist on the targeted form
' Parameters:   frm - form holding crumb labels
'               aryCrumbs - breadcrumb array
'               separator - non-clickable value between crumbs, default = >
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - XX
' =================================
Public Sub PrepareCrumbs(frm As SubForm, aryCrumbs As Variant, Optional separator = ">")
    
    Dim ctrl As Control
    Dim i As Integer, intLastCtrlWidth As Integer, intLastCtrlPosition As Integer
    Dim strNum As String, strCtrlName As String, strCtrlSeparator As String
    
    'initialize
    intLastCtrlPosition = 10
    
    'avoid flicker
    'Painting = False
    
    For i = 1 To UBound(aryCrumbs)
        ' set lbl caption
        If (i < 10) Then
            strNum = 0 & i
        Else
            strNum = i
        End If
        
        strCtrlName = "lblCrumb" & strNum
        
        With frm.Controls(strCtrlName)
       
            If .ControlType = acLabel Then
                'label control
                .Caption = aryCrumbs(i)
            Else
                'hyperlink control (displaytext vs caption)
                .Value = aryCrumbs(i)
            End If
            
            'set control position
            If intLastCtrlPosition > frm.Controls(strCtrlName).Parent.Width Then
                .Left = frm.Controls(strCtrlName).Parent.Width - .Width
            Else
                .Left = intLastCtrlPosition
            End If
            
            'set control width
'            setControlWidth frm.Controls(strCtrlName), , frm.Controls(strCtrlName).Parent.Width
            
            'save new ctrl width for setting separator position
            intLastCtrlWidth = .Width
        
        End With
        
        'display the separator
        If (i < UBound(aryCrumbs)) Then
          strCtrlSeparator = "lblSep" & strNum
          With frm.Controls(strCtrlSeparator)
            .Left = intLastCtrlPosition + intLastCtrlWidth + 10
            .Caption = separator
            .Visible = True
            
            'determine position of next control
            intLastCtrlPosition = .Left + .Width + 10
          End With
        End If
        
    Next i
    
    'ready for viewing
    'Painting = True
    
End Sub

' ---------------------------------
' SUB:          initializeControls
' Description:  set initial control values
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, Sept 2014 for NCPN tools
' Adapted:      -
' Revisions:    BLC - 9/01/2014 - initial version
' ---------------------------------
Public Sub initializeControls(frm As Form)
    On Error GoTo Err_Handler
    Dim aryCtrls() As Variant
    Dim ctrlName As String, tgtCtrlName As String
    Dim i As Integer

    With frm
        Select Case .name
            Case "frm_Set_Defaults"
                'TempVars not yet populated -> use fsub_DbAdmin control defaults
                aryCtrls = Array("User", "Project", "GPS_model", "Park", "Datum", "Declination", "Timeframe", "Project")
                For i = 0 To UBound(aryCtrls)
                    ctrlName = "tbx" & aryCtrls(i)
                    If aryCtrls(i) = "Declination" Or _
                       aryCtrls(i) = "Timeframe" Or _
                       aryCtrls(i) = "Project" Then
                        tgtCtrlName = "tbx" & aryCtrls(i)
                    Else
                        tgtCtrlName = "cbx" & aryCtrls(i)
                    End If
                    .Controls(tgtCtrlName) = Forms!frm_Switchboard.fsub_DbAdmin.Form.Controls(ctrlName).Value
                Next
        End Select
    End With
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - initializeControls[mod_Common_UI])"
    End Select
    Resume Exit_Procedure
End Sub