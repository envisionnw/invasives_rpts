' =================================
' MODULE:       mod_Linked_Tables
' Description:  Standard module for generic functions related to linked tables
' Source/date:  John R. Boetsch, May 24, 2006
' Revisions:    JRB, 7/9/2009 - simplified by moving certain functions to another module
'               JRB, 12/30/2009 - moved fxnVerifyLinks to another module

Option Compare Database
Option Explicit

' =================================
' FUNCTION:     fxnCheckLink
' Description:  Checks the status of the link for the specified table
' Parameters:   strTableName - name of the table to check
' Returns:      True (valid link) or False
' Throws:       none
' References:   none
' Source/date:  From Access97 Developer's Handbook by Litwin, Getz and Gilbert (Sybex)
'               Copyright 1997.  All Rights Reserved
'               Created 09/13/94 pel; Last modified 07/10/96 pel.
' Revisions:    John R. Boetsch, May 17, 2006 - updated documentation, added error traps
' =================================
Public Function fxnCheckLink(strTableName As String) As Boolean
    On Error GoTo Err_Handler

    Dim varRet As Variant

    On Error Resume Next
    ' Check for failure.  If can't determine the name of
    ' the first field in the table, the link must be bad.
    varRet = CurrentDb.tabledefs(strTableName).Fields(0).name
    If Err <> 0 Then
        fxnCheckLink = False
    Else
        fxnCheckLink = True
    End If
    
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnCheckLink)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnIsODBC
' Description:  Determine whether the input table is connected by ODBC
' Parameters:   strTableName - string for the table name
' Returns:      True (if table object in collection and ODBC) or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/24/2009
' Revisions:    <name, date, desc - add lines as you go>
' =================================
Public Function fxnIsODBC(strTableName As String) As Boolean
    On Error GoTo Err_Handler

    Dim strCriteria As String

    strCriteria = "(([Name])=""" & strTableName & """) AND (([Type]) In (1, 4, 6))"
    If DLookup("Type", "MSysObjects", strCriteria) = 4 Then
        ' ODBC connection
        fxnIsODBC = True
    Else
        ' Native table or linked Access table
        fxnIsODBC = False
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnIsODBC)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnLinkedDatabase
' Description:  Returns the database file path (Access) or the database name (ODBC) for
'                   a linked table name
' Parameters:   strTableName - the name of the linked table
' Returns:      database name for the linked table, or empty string ("") if none
' Throws:       none
' References:   fxnParseConnectionStr
' Source/date:  John R. Boetsch, 6/24/2009
' Revisions:    <name, date, desc - add lines as you go>
' =================================
Public Function fxnLinkedDatabase(ByVal strTableName As String) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    strTemp = fxnParseConnectionStr(CurrentDb.tabledefs(strTableName).Connect)
    fxnLinkedDatabase = strTemp

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3265
        MsgBox "The table '" & strTableName & "' was not found in the front-end.", _
            vbCritical, "Error encountered (#" & Err.Number & " - fxnCheckLink)"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnLinkedDatabase)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnParseConnectionStr
' Description:  Return the specified portion of the linked table connection string
' Parameters:   strConnStr - linked table connection string
'               strComponent - optional string to specify the portion to return
'                   (default "DATABASE=")
'               strDelimiter - optional string delimiter (default ";")
'               blnIsFound - optional reference variable to incidate whether the
'                   specified string component is found in the connection string
' Returns:      connection string component, or empty string ("") if not found
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/24/2009
' Revisions:    <name, date, desc - add lines as you go>
' =================================
Public Function fxnParseConnectionStr(strConnStr As String, _
    Optional strComponent As String = "DATABASE=", _
    Optional strDelimiter As String = ";", _
    Optional blnIsFound As Boolean = False) As String

    On Error GoTo Err_Handler

    Dim varStartPos As Variant
    Dim varEndPos As Variant
    Dim varLength As Variant
    Dim strResult As String

    varStartPos = InStr(1, strConnStr, strComponent, vbTextCompare)
    If IsNull(varStartPos) Or IsEmpty(varStartPos) Or varStartPos = 0 Then
        ' The component is not found in the connection string
        blnIsFound = False
    Else
        blnIsFound = True
        ' Determine the end position of the database string
        varEndPos = InStr(varStartPos, strConnStr, strDelimiter, vbTextCompare)
        If varEndPos > varStartPos Then
            ' There is a delimiter following the desired string
            varStartPos = varStartPos + Len(strComponent)
            varLength = varEndPos - varStartPos
            fxnParseConnectionStr = Mid(strConnStr, varStartPos, varLength)
        Else
            varLength = Len(strConnStr) - varStartPos + 1 - Len(strComponent)
            fxnParseConnectionStr = Right(strConnStr, varLength)
        End If
    End If
    
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnParseConnectionStr)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnTestODBCConnection
' Description:  Uses a pass-through query to test an ODBC connection and to trap ODBC errors
' Parameters:   strTableName - the name of the linked table
'               strConnStr - optional linked table connection string
'               varSQL - optional SQL statement to execute
'               blnRetErrMsg - optional flag to show error msg if the test fails (default=True)
' Returns:      True if the connection returns records, otherwise False
' Throws:       none
' References:   fxnParseConnectionStr
' Source/date:  John R. Boetsch, 6/24/2009 (adapted from http://support.microsoft.com/kb/210319)
' Revisions:    <name, date, desc - add lines as you go>
' =================================
Function fxnTestODBCConnection(strTableName As String, _
    Optional ByVal strConnStr As String, _
    Optional varSQL As Variant, _
    Optional blnRetErrMsg As Boolean = True) As Boolean

    On Error GoTo Err_Handler

    fxnTestODBCConnection = False   ' Default in case of error

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strDbName As String

    ' Create a blank pass-through query
    Set db = CurrentDb()
    Set qdf = db.CreateQueryDef("")

    ' If no revised connection string was passed, use the current connection string
    If strConnStr = "" Then strConnStr = CurrentDb.tabledefs(strTableName).Connect
    strDbName = fxnParseConnectionStr(strConnStr)

    ' Update the connection string for the pass-through query, set to not return records
    qdf.Connect = strConnStr
    qdf.ReturnsRecords = False

    If IsMissing(varSQL) Then
        ' If no query statement passed, select a few records to test the connection string
        qdf.sql = "SELECT TOP 2 * FROM " & strTableName
    Else: qdf.sql = varSQL
    End If
    qdf.Execute

    ' Set to true (if no errors)
    fxnTestODBCConnection = True

Exit_Procedure:
    On Error Resume Next
    Set db = Nothing
    Set qdf = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3151       ' Connection failed
        If blnRetErrMsg Then _
        MsgBox "Cannot connect to the specified database/table:" & vbCrLf & vbCrLf & _
            "  Db: " & strDbName & vbCrLf & "  Table: " & strTableName, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnTestODBCConnection)"
      Case 3146, 208  ' Connection failed
        If blnRetErrMsg Then _
        MsgBox "Cannot find the table in the specified database:" & vbCrLf & vbCrLf & _
            "  Db: " & strDbName & vbCrLf & "  Table: " & strTableName, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnTestODBCConnection)"
      Case 3305       ' Invalid pass-through connection string
        MsgBox "Invalid pass-through query connection string ..." & vbCrLf & _
            strTableName & " may not be an ODBC-linked table.", vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnTestODBCConnection)"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnTestODBCConnection)"
    End Select
    Resume Exit_Procedure

End Function