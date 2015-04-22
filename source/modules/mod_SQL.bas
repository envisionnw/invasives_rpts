Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_SQL
' VERSION:      1.01
' Description:  Database/SQL properties, functions & subroutines
'
' Source/date:  Bonnie Campbell, 7/24/2014
' Revisions:    BLC, 7/24/2014 - 1.00 - initial version
'               BLC, 8/19/2014 - 1.01 - added versioning
' =================================

' ---------------------------------
' PROPERTY:     dbCurrent
' Description:  Gets a single instance of the current db to avoid multiple calls
'               to CurrentDb which can yield to Error 3048 "Cannot open any more databases" errors
'               due to multiple open db
' Parameters:   -
' Returns:      current database object
' Throws:       -
' References:   -
' Source/date:  Dirk Goldgar, MS Access MVP - May 22, 2013
'   http://social.msdn.microsoft.com/Forums/office/en-US/9993d229-8a00-4a59-a796-dfa2dad505bc/cannot-open-any-more-databases?forum=accessdev
' Adapted:      Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Revisions:    BLC, 7/23/2014 - XX
' ---------------------------------
Private m_db As DAO.Database
Public Property Get dbCurrent() As DAO.Database

    If (m_db Is Nothing) Then
        Set m_db = CurrentDb
    End If

    Set dbCurrent = m_db

End Property

' ---------------------------------
' FUNCTION:     getSQL
' Description:  Retrieve query SQL string using query name
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:
'   Albert D. Kallal  (Access MVP) Edmonton, Alberta Canada kallal@msn.com - Sept 8, 2010
'   http://social.msdn.microsoft.com/Forums/office/en-US/3a26a941-b75b-49e4-bfe8-10c152f2b6c0/sql-or-querydef-in-vba-code?forum=accessdev
'   Daniel Pineault, CARDA Consultants Inc. - June 10, 2010
'   http://www.devhut.net/2010/06/10/ms-access-vba-edit-a-querys-sql-statement/
' Adapted:      Bonnie Campbell, July, 2014 for NCPN tools
' Revisions:    BLC, 7/23/2014 - initial version
' ---------------------------------
Public Function getSQL(strQuery As String) As String
On Error GoTo Err_Handler:

   getSQL = dbCurrent.QueryDefs(strQuery).sql
   
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getSql[mod_Point_Intercept])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     getWhereSQL
' Description:  Prepare a SQL WHERE clause based on the parameters, parameter types, fields, and
'               current WHERE clause (strWhere) passed into the function
' Assumptions:  Assumes parameters passed through params will each have the parameter name, type, and field name
'                   params(x,0) = parameter value
'                   params(x,1) = parameter type
'                   params(x,2) = database field name
'               NOTE: The function does not currently handle dependent parameters which require
'                     the presence of other parameters to be included in the WHERE clause
'                     These have to be accommodated separately.
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, August, 2014 for NCPN tools
' Adapted:      Bonnie Campbell, July, 2014 for NCPN tools
' Revisions:    BLC, 8/11/2014 - initial version
' ---------------------------------
Public Function getWhereSQL(strWhere As String, params As Variant) As String
On Error GoTo Err_Handler:
Dim blnCheck As Boolean
Dim strParam As String
Dim i As Integer

    'default
    blnCheck = False

    For i = 0 To UBound(params) - 1
    
        'handle empty field values
        If Len(params(i, 2)) > 0 Then
    
            'handle when param isn't the only parameter (need ' AND ' in SQL WHERE clause)
            If Len(strWhere) > 0 Then strWhere = strWhere & " AND"
    
            'check if parameter is is non-empty (string) or non-zero (integer)
            Select Case params(i, 1)
                Case "string"
                    If Len(Trim(params(i, 0))) > 0 Then blnCheck = True
                    strParam = "'" & params(i, 0) & "'"
                Case "integer"
                    If params(i, 0) > 0 Then blnCheck = True
                    strParam = params(i, 0)
            End Select
        
            'prepare SQL
            If Not IsNull(params(i, 0)) And blnCheck Then
             strWhere = strWhere & " " & params(i, 2) & " = " & strParam
            End If
        
        Else
            Exit For 'done
        End If
    Next
    
   getWhereSQL = strWhere
   
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getSql[mod_Point_Intercept])"
    End Select
    Resume Exit_Function
End Function