Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_UI
' Level:        Framework module
' Version:      1.00
' Description:  User interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - initial version
' =================================

' =================================
' FUNCTION:     FormIsOpen
' Description:  Indicates whether or not the specific form is open in form view
' Parameters:   none
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 5/5/2006 as fxnSwitchboardIsOpen
' Adapted:      Bonnie Campbell, 4/30/2015 for NCPN tools
' Revisions:    BLC, 4/30/2015 - initial version
' =================================
Public Function FormIsOpen(strFormName As String) As Boolean
    On Error GoTo Err_Handler

    FormIsOpen = False    ' Default in case of error

    If CurrentProject.AllForms(strFormName).IsLoaded = True Then
        If CurrentProject.AllForms(strFormName).CurrentView = 1 Then
            FormIsOpen = True
        End If
    End If

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FormIsOpen[mod_UI])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnSwitchboardIsOpen
' Description:  Indicates whether or not the switchboard form is open in form view
' Parameters:   none
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 5/5/2006
' Revisions:    JRB, 5/5/2006 - initial version
'               BLC, 4/30/2015  - moved to mod_Db framework module from mod_Custom_Functions
' =================================
Public Function fxnSwitchboardIsOpen() As Boolean
    On Error GoTo Err_Handler

    fxnSwitchboardIsOpen = False    ' Default in case of error

    Dim strSwitchboardName As String

    strSwitchboardName = "frm_Switchboard"

    If CurrentProject.AllForms(strSwitchboardName).IsLoaded = True Then
        If CurrentProject.AllForms(strSwitchboardName).CurrentView = 1 Then
            fxnSwitchboardIsOpen = True
        End If
    End If

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnSwitchboardIsOpen[mod_UI])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnFormIsLoaded
' Description:  Returns whether the specified form is loaded in Form or Datasheet view
' Parameters:   strFormName - string for the name of the form to check
' Returns:      True if the specified form is open in Form view or Datasheet view
' Throws:       none
' References:   none
' Source/date:  From Northwind sample database, date unknown
' Revisions:    John R. Boetsch, 6/17/2009 - error trapping, documentation
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_UI
' =================================
Public Function fxnFormIsLoaded(ByVal strFormName As String) As Integer
    On Error GoTo Err_Handler
 
    ' These variables are used to test the return values of the SysCmd function
    '  and the CurrentView property of the requested form.
    Const conObjStateClosed = 0
    Const conDesignView = 0

    ' Use the SysCmd function to check the current state of the requested form.
    '  Possible states: not open or nonexistent, open, new, or changed but not saved
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        ' Checks for the current view of the requested form, assuming the previous statement
        '   found it to be open ... return True if open and not in design view
        If Forms(strFormName).CurrentView <> conDesignView Then
            fxnFormIsLoaded = True
        End If
    End If
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnFormIsLoaded[mod_UI])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     fxnHideObject
' Description:  Changes the hidden property of an object to hide / show in the database window
' Parameters:   strObjectName - name of the object (string)
'               blnHide - True to hide, False to show (default True)
'               varType - object type (default acTable)
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/25/2009
' Revisions:    JRB, 6/25/2009 - initial version
'               BLC, 4/30/2015 - move from mod_Utilities to mod_UI
' =================================
Public Function fxnHideObject(strObjectName As String, _
    Optional blnHide As Boolean = True, Optional varType As Variant = acTable)

    On Error GoTo Err_Handler

    SetHiddenAttribute varType, strObjectName, blnHide

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnHideObject[mod_UI])"
    End Select
    Resume Exit_Function
End Function


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
On Error GoTo Err_Handler

    Dim pg As page
    
    For Each pg In ctrl.Pages
        If pg.name = strTabName Then
            ctrl.Pages(pg.name).Visible = True
        Else
            ctrl.Pages(pg.name).Visible = False
        End If
    Next pg
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tabPageUnhide[mod_UI])"
    End Select
    Resume Exit_Procedure
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
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
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
            "Error encountered (#" & Err.Number & " - buttonHighlight[mod_UI])"
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
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
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
            "Error encountered (#" & Err.Number & " - buttonUnHighlight[mod_UI])"
    End Select
    Resume Exit_Procedure
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
' Revisions:    BLC, 6/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' =================================
Public Function fxnCrumbsToArray(strCrumbs As String, Optional delimiter = "|")

On Error GoTo Err_Handler

    Dim strCrumbTrail As String

    If Len(strCrumbs) > 0 Then
        Dim aryCrumbs As Variant
        
        aryCrumbs = Split(strCrumbs, delimiter)
        
    End If

    fxnCrumbsToArray = aryCrumbs
    
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnCrumbsToArray[mod_UI])"
    End Select
    Resume Exit_Procedure
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
' Revisions:    BLC, 6/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' =================================
Public Sub PrepareCrumbs(frm As SubForm, aryCrumbs As Variant, Optional separator = ">")
 On Error GoTo Err_Handler
 
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
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PrepareCrumbs[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          repaintParentForm
' Description:  Repaints the control's parent(or grandparent or great grandparent...) form
' Parameters:   ctl - control whose parent form you're looking to repaint
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell August, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 8/20/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
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
            "Error encountered (#" & Err.Number & " - findParentForm[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

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
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' =================================
Public Function fxnHTMLConvert(strHTML As String) As Long
    Rem converts a HTML color code number such as #D8B190 to an RGB value.
    fxnHTMLConvert = RGB(CInt("&H" & Mid(strHTML, 2, 2)), CInt("&H" & Mid(strHTML, 4, 2)), CInt("&H" & Mid(strHTML, 6, 2)))
End Function