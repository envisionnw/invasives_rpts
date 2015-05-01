Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Linked_Tables
' Level:        Framework module
' Description:  Linked table related functions & subroutines
'
' Adapted from: John R. Boetsch, May 24, 2006
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    JRB, 7/9/2009 - simplified by moving certain functions to another module
'               JRB, 12/30/2009 - moved fxnVerifyLinks to another module
'               ------------------------------------------------------------------------
'               BLC, 4/30/2015 - added fxnVerifyLinks, fxnRefreshLinks, fxnVerifyLinkTableInfo,
'                                fxnMakeBackup from mod_Custom_Functions
' =================================

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

' =================================
' FUNCTION:     fxnRefreshLinks
' Description:  Updates the link to the specified back-end database tables after first
'               verifying that the tables exist in the specified link file
' Parameters:   strDbName - name of the database to refresh
'               strNewConnStr - updated connection string
'               blnIsODBC - flag to indicate that the back-end is ODBC (default = False)
' Returns:      True (successfully relinked) or False
' Throws:       none
' References:   fxnParseConnectionStr, fxnTestODBCConnection
' Source/date:  Susan Huse, fall 2004 and Mark A. Wotawa, 02/08/2000
' Revisions:    John R. Boetsch, 5/22/2006 - combined verify and refresh functions
'                   for table links, fixed meter increment problem updated documentation
'                   and error traps
'               JRB, 7/9/2009 - updated to accommodate ODBC links and to update the table
'                   description in tsys_Link_Tables for Access tables
'               JRB, 12/30/2009 - updated to use the popup progress meter form
'               -------------------------------------------------------------------------
'               BLC, 4/30/2015 - moved to mod_Linked_Tables from mod_Custom_Functions & renamed RefreshLinks
' =================================
Public Function RefreshLinks(strDbName As String, ByVal strNewConnStr As String, _
    Optional ByVal blnIsODBC As Boolean = False) As Boolean
    On Error GoTo Err_Handler

    Dim varFileName As Variant
    Dim dbGet As DAO.Database
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim intNumTables As Integer
    Dim varReturn As Variant
    Dim intI As Integer
    Dim strTable As String
    Dim strDesc As String
    Dim strSQL As String
    Dim frm As Form             ' Reference to the progress popup form
    Dim strProgForm As String   ' Name of the progress popup form
    Dim strProgress As String   ' Progress bar string

    RefreshLinks = False   ' Default unless all tables verified

    Set db = CurrentDb
    Set rst = db.OpenRecordset("SELECT * FROM tsys_Link_Tables WHERE " & _
                "[tsys_Link_Tables]![Link_db] = """ & strDbName & """", dbOpenSnapshot)

    ' Counts the number of tables in the system table associated with this db
    rst.MoveLast    ' Need to do this to make the record count accurate
    intNumTables = rst.RecordCount

    ' Initialize the progress popup form
    strProgForm = "frm_Progress_Meter"
    DoCmd.OpenForm strProgForm
    Set frm = Forms!frm_Progress_Meter
    frm.Caption = " Update table connections"
    frm!txtPercent = 0

    If blnIsODBC = False Then   ' Access back-end
        ' Opens the target database and the current system table containing the list
        '   of tables for refreshing links
        varFileName = fxnParseConnectionStr(strNewConnStr)
        Set dbGet = DBEngine.OpenDatabase(varFileName)

        ' First pass to verify the tables in the new back-end database (avoids partial updates)
        '   Initialize the system meter to indicate progress
        varReturn = SysCmd(acSysCmdInitMeter, "Verifying tables in " & _
            strDbName, intNumTables)
        ' Update the message below the progress meter
        frm!txtMsg = "Verifying tables in " & strDbName
        intI = 0
        rst.MoveFirst
        Do Until rst.EOF
            intI = intI + 1
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            ' Update the popup progress meter
            frm!txtPercent = Round(100 * intI / intNumTables)
            ' Update the progress bar in the progress popup with sequential "�" characters
            '   which look like a bar because of the font of the control (20 characters=100%)
            strProgress = String(Round(19 * intI / intNumTables), "�")
            frm!txtProgress = strProgress
            frm.Repaint
            strTable = rst![Link_table]
            Debug.Print strTable
            varReturn = dbGet.tabledefs(strTable).Fields(0).name
            rst.MoveNext
        Loop

        ' Second pass to refresh all links now that they are validated
        '   Reinitialize the system meter to indicate progress
        varReturn = SysCmd(acSysCmdInitMeter, "Updating table links in " & _
            strDbName, intNumTables)
        ' Update the message below the progress meter
        frm!txtMsg = "Updating table links in " & strDbName
        intI = 0
        rst.MoveFirst
        Do Until rst.EOF
            intI = intI + 1
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            ' Update the popup progress meter
            frm!txtPercent = Round(100 * intI / intNumTables)
            ' Update the progress bar in the progress popup with sequential "�" characters
            '   which look like a bar because of the font of the control (20 characters=100%)
            strProgress = String(Round(19 * intI / intNumTables), "�")
            frm!txtProgress = strProgress
            frm.Repaint
            strTable = rst![Link_table]
Debug.Print strTable
            ' Update and refresh the table connection
            Set tdf = db.tabledefs(strTable)
            tdf.Connect = strNewConnStr
            tdf.RefreshLink
            ' Update the table description in tsys_Link_Tables
            ' Set default description in case there is none
            strDesc = " - no description - "
            strDesc = tdf.Properties("Description") ' Throws trapped error 3270 if none
            strSQL = "UPDATE tsys_Link_Tables " & _
                "SET tsys_Link_Tables.Description_text=""" & strDesc & _
                """ WHERE (((tsys_Link_Tables.Link_table)=""" & strTable & """));"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
            rst.MoveNext
        Loop
    Else    ' ODBC back-end
        ' First pass to verify the tables in the new back-end database (avoids partial updates)
        '   Initialize the system meter to indicate progress
        varReturn = SysCmd(acSysCmdInitMeter, "Verifying tables in " & _
            strDbName, intNumTables)
        ' Update the message below the progress meter
        frm!txtMsg = "Verifying tables in " & strDbName
        intI = 0
        rst.MoveFirst
        Do Until rst.EOF
            intI = intI + 1
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            ' Update the popup progress meter
            frm!txtPercent = Round(100 * intI / intNumTables)
            ' Update the progress bar in the progress popup with sequential "�" characters
            '   which look like a bar because of the font of the control (20 characters=100%)
            strProgress = String(Round(19 * intI / intNumTables), "�")
            frm!txtProgress = strProgress
            frm.Repaint
            strTable = rst![Link_table]
            If fxnTestODBCConnection(strTable, strNewConnStr) = False Then GoTo Exit_Procedure
            rst.MoveNext
        Loop

        ' Second pass to refresh all links now that they are validated
        '   Reinitialize the system meter to indicate progress
        varReturn = SysCmd(acSysCmdInitMeter, "Updating table links in " & _
            strDbName, intNumTables)
        ' Update the message below the progress meter
        frm!txtMsg = "Updating table links in " & strDbName
        intI = 0
        rst.MoveFirst
        Do Until rst.EOF
            intI = intI + 1
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            ' Update the popup progress meter
            frm!txtPercent = Round(100 * intI / intNumTables)
            ' Update the progress bar in the progress popup with sequential "�" characters
            '   which look like a bar because of the font of the control (20 characters=100%)
            strProgress = String(Round(19 * intI / intNumTables), "�")
            frm!txtProgress = strProgress
            frm.Repaint
            strTable = rst![Link_table]
            ' Update and refresh the table connection
            Set tdf = db.tabledefs(strTable)
            ' Use test again to trap errors
            If fxnTestODBCConnection(strTable, strNewConnStr) = True Then
                tdf.Connect = strNewConnStr
                tdf.RefreshLink
            Else
                GoTo Exit_Procedure
            End If
            rst.MoveNext
        Loop
    End If

    RefreshLinks = True    ' Links successfully updated

Exit_Procedure:
    On Error Resume Next
    DoCmd.SetWarnings True
    varReturn = SysCmd(acSysCmdRemoveMeter)
    DoCmd.Close acForm, strProgForm, acSaveNo
    Set frm = Nothing
    dbGet.Close
    Set dbGet = Nothing
    rst.Close
    Set tdf = Nothing
    Set rst = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    RefreshLinks = False
    Select Case Err.Number
      Case 3021
        MsgBox "Error #" & Err.Number & ":  There are no table links associated " & _
            "with one or more of these files." & vbCrLf & "Please contact the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshLinks[mod_Linked_Tables])"
      Case 3024
        MsgBox "Error #" & Err.Number & ":  Cannot find the following file:" & _
            vbCrLf & vbCrLf & varFileName, vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshLinks[mod_Linked_Tables])"
      Case 3078   ' Also got this error if the function call SQL string has a bad
                '   reference to the system table
        MsgBox "Error #" & Err.Number & ":  The following table is not native " & _
            "to the selected database file." & vbCrLf & "Please make sure you " & _
            "browsed to to the correct file." & vbCrLf & vbCrLf & strTable, _
            vbCritical, "Error encountered (#" & Err.Number & " - RefreshLinks[mod_Linked_Tables])"
      Case 3061   ' Bad parameters for the SQL string
        MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshLinks[mod_Linked_Tables])"
      Case 3265
        MsgBox "Error #" & Err.Number & ":  The database file is missing the " & _
            "following table:" & vbCrLf & vbCrLf & strTable, _
            vbCritical, "Error encountered (#" & Err.Number & " - RefreshLinks[mod_Linked_Tables])"
      Case 3219 ' Trying to update a link on top of an imported table
        MsgBox "Error #" & Err.Number & ":  You are trying to update a link to " & _
            "a table that has already been imported." & vbCrLf & vbCrLf & _
            strTable & vbCrLf & vbCrLf & "Please call the database " & _
            "administrator to help you relink this table manually." & vbCrLf & _
            "Afterwards you will be able to automatically update links again.", _
            vbCritical, "Error encountered (#" & Err.Number & " - RefreshLinks[mod_Linked_Tables])"
      Case 3270     ' Property not found (TableDefs description)
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshLinks[mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnVerifyLinkTableInfo
' Description:  Verifies that the information in tsys_Link_Dbs and tsys_Link_Tables is
'                   complete and matches that in MSysObjects
' Parameters:   none
' Returns:      True if the information matches and there are no problems, False otherwise
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 7/9/2009
' Revisions:    JRB, 7/27/2009 - added a do loop to update missing table descriptions
'               -------------------------------------------------------------------------
'               BLC, 4/30/2015 - moved to mod_Linked_Tables from mod_Custom_Functions
' =================================
Public Function fxnVerifyLinkTableInfo() As Boolean
    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim intNRecs As Integer
    Dim strTable As String
    Dim strDesc As String
    Dim blnHasError As Boolean
    Dim strSQL As String

    Set db = CurrentDb
    blnHasError = False             ' Flag to indicate error status

    ' First make sure that there are linked tables
    intNRecs = DCount("*", "MSysObjects", "([Type] In (4,6)) And ([Name] Not Like '~*')")
    If intNRecs = 0 Then    ' No linked tables in the recordset
        MsgBox "There are no linked tables found in the systems tables." & _
            vbCrLf & "Please contact the database administrator before " & _
            "using this application.", vbCritical, "Application error (fxnVerifyLinkTableInfo[mod_Linked_Tables])"
        GoTo Exit_Procedure
    End If

    ' Look for linked table records that no longer actually exist in the database
    intNRecs = DCount("*", "qsys_Linked_tables_not_in_MSysObjects")
    If intNRecs > 0 Then
        Set rst = db.OpenRecordset("qsys_Linked_tables_not_in_MSysObjects", _
            dbOpenSnapshot)
        Do Until rst.EOF
            ' Delete mismatched records from tsys_Link_Tables
            strSQL = "DELETE * FROM tsys_Link_Tables WHERE ([Link_table]=""" & _
                rst![Link_table] & """);"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
            rst.MoveNext
        Loop
        rst.Close
        ' Throw an error if there are still mismatched records
        If DCount("*", "qsys_Linked_tables_not_in_MSysObjects") > 0 Then
            blnHasError = True
            DoCmd.OpenQuery "qsys_Linked_tables_not_in_MSysObjects", , acReadOnly
        End If
    End If

    ' Look for linked tables that are not in the application table
    intNRecs = DCount("*", "qsys_Linked_tables_not_in_tsys_Link_Tables")
    If intNRecs > 0 Then
        DoCmd.SetWarnings False
        ' Run the append query to add databases not in tsys_Link_Dbs
        DoCmd.OpenQuery "qsys_Linked_dbs_not_in_tsys_Link_Dbs"
        ' Append missing table records to tsys_Link_Tables
        strSQL = "INSERT INTO tsys_Link_Tables " & _
            "( Link_table, Link_db ) " & _
            "SELECT qsys_Linked_tables_not_in_tsys_Link_Tables.CurrTable, " & _
            "qsys_Linked_tables_not_in_tsys_Link_Tables.CurrDb " & _
            "FROM qsys_Linked_tables_not_in_tsys_Link_Tables;"
        DoCmd.RunSQL strSQL
        DoCmd.SetWarnings True
        ' Update descriptions
        Set rst = db.OpenRecordset("SELECT * FROM tsys_Link_Tables " & _
            "WHERE tsys_Link_Tables.Description_text Is Null", dbOpenSnapshot)
        Do Until rst.EOF
            strTable = rst![Link_table]
            Set tdf = db.tabledefs(strTable)
            ' Update the table description in tsys_Link_Tables
            ' Set default description in case there is none
            strDesc = " - no description - "
            strDesc = tdf.Properties("Description") ' Throws trapped error 3270 if none
            strSQL = "UPDATE tsys_Link_Tables " & _
                "SET tsys_Link_Tables.Description_text=""" & strDesc & _
                """ WHERE (((tsys_Link_Tables.Link_table)=""" & strTable & """));"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
            rst.MoveNext
        Loop
        rst.Close
        ' Throw an error if there are still mismatched records
        If DCount("*", "qsys_Linked_tables_not_in_tsys_Link_Tables") > 0 Then
            blnHasError = True
            DoCmd.OpenQuery "qsys_Linked_tables_not_in_tsys_Link_Tables", , acReadOnly
        End If
    End If

    ' Look for linked db records without child table records
    intNRecs = DCount("*", "qsys_Linked_dbs_without_table_records")
    If intNRecs > 0 Then
        Set rst = db.OpenRecordset("qsys_Linked_dbs_without_table_records", _
            dbOpenSnapshot)
        Do Until rst.EOF
            ' Delete mismatched records from tsys_Link_Dbs
            strSQL = "DELETE * FROM tsys_Link_Dbs WHERE ([Link_db]=""" & _
                rst![Link_db] & """);"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
            rst.MoveNext
        Loop
        rst.Close
        ' Throw an error if there are still mismatched records
        If DCount("*", "qsys_Linked_dbs_without_table_records") > 0 Then
            blnHasError = True
            DoCmd.OpenQuery "qsys_Linked_dbs_without_table_records", , acReadOnly
        End If
    End If

    ' Look for records with mismatched db name, server, file path, or ODBC info
    intNRecs = DCount("*", "qsys_Linked_tables_mismatched_info")
    If intNRecs > 0 Then
        blnHasError = True
        DoCmd.OpenQuery "qsys_Linked_tables_mismatched_info"
    End If

    ' Warn the user if an error was found
    If blnHasError Then
        MsgBox "The application tables need to be updated with" & vbCrLf & _
            "correct information about the linked back-end" & vbCrLf & _
            "databases and tables before the application can" & vbCrLf & _
            "be used." & vbCrLf & vbCrLf & "Please contact the database administrator.", _
            vbCritical, "Application error (fxnVerifyLinkTableInfo[mod_Linked_Tables])"
    End If

    fxnVerifyLinkTableInfo = Not blnHasError

Exit_Procedure:
    On Error Resume Next
    DoCmd.SetWarnings True
    rst.Close
    Set tdf = Nothing
    Set rst = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3135, 3061, 3078  ' Problem with SQL syntax, or ref to nonexistent object, etc.
        MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnVerifyLinkTableInfo)"
      Case 3011, 7874   ' System table not found
         MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnVerifyLinkTableInfo)"
      Case 3265     ' Field name in the system table improperly specified
        MsgBox "Error #" & Err.Number & ":  System table field not found." & _
            vbCrLf & "Please notify the database administrator before using " & _
            "this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnVerifyLinkTableInfo)"
      Case 3270     ' Property not found (TableDefs description)
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnVerifyLinkTableInfo)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnVerifyLinks
' Description:  Loops through all of the linked tables to verify valid links
' Parameters:   none
' Returns:      True (no link errors) or False
' Throws:       none
' References:   fxnCheckLink
' Source/date:  John R. Boetsch, May 24, 2006
' Revisions:    JRB, 7/8/2009 - simplified recordset and error traps
'               JRB, 10/8/2009 - changed progress meter message
'               JRB, 12/30/2009 - updated to use the popup progress meter form
'               -------------------------------------------------------------------------
'               BLC, 4/30/2015 - moved to mod_Linked_Tables from mod_Custom_Functions & renamed VerifyLinks
' =================================
Public Function VerifyLinks() As Boolean
    On Error GoTo Err_Handler

    Dim rst As DAO.Recordset
    Dim intNumTables As Integer
    Dim intI As Integer
    Dim varReturn As Variant
    Dim strLinkTableName As String
    Dim frm As Form             ' Reference to the progress popup form
    Dim strProgForm As String   ' Name of the progress popup form
    Dim strProgress As String   ' Progress bar string

    VerifyLinks = False  ' Default unless successful

    ' Set the recordset to the system table to show all linked tables except those
    '   that have recently been deleted (which have names starting with '~'
    Set rst = CurrentDb.OpenRecordset("SELECT MSysObjects.Name, MSysObjects.Database " & _
        "FROM MSysObjects " & _
        "WHERE ((MSysObjects.Name) Not Like '~*') AND ((MSysObjects.Type) In (4,6)) " & _
        "ORDER BY MSysObjects.Name;", dbOpenSnapshot)

    ' Counts the number of linked tables in the recordset
    rst.MoveLast    ' Need to do this to make the record count accurate
    intNumTables = rst.RecordCount

    ' Initialize the progress popup form
    strProgForm = "frm_Progress_Meter"
    DoCmd.OpenForm strProgForm
    Set frm = Forms!frm_Progress_Meter
    frm.Caption = " Verifying table connections"
    frm!txtPercent = 0
    ' Initialize the message below the progress meter
    frm!txtMsg = " ... Please wait ..."

    '   Initialize the system meter to indicate progress
    varReturn = SysCmd(acSysCmdInitMeter, "Verifying table connections", intNumTables)
    intI = 0
    rst.MoveFirst

    ' Loop through each record and check for bad links
    '   Send to error handler if a bad link is encountered
    Do Until rst.EOF
        intI = intI + 1
        ' Update the status bar progress meter
        varReturn = SysCmd(acSysCmdUpdateMeter, intI)
        ' Update the popup progress meter
        frm!txtPercent = Round(100 * intI / intNumTables)
        ' Update the progress bar in the progress popup with sequential "�" characters
        '   which look like a bar because of the font of the control (20 characters=100%)
        strProgress = String(Round(19 * intI / intNumTables), "�")
        frm!txtProgress = strProgress
        frm.Repaint
        strLinkTableName = rst![name]
        ' Make sure the linked table opens properly
        If fxnCheckLink(strLinkTableName) = False Then
            ' Unable to open a linked table (not a critical error)
            MsgBox "Unable to open the following table:" & vbCrLf & vbCrLf & _
                strLinkTableName, vbExclamation, "Broken table links"
            GoTo Exit_Procedure
        Else
        ' Table link is valid
            rst.MoveNext
        End If
    Loop

    ' If no bad links were encountered
    VerifyLinks = True

Exit_Procedure:
    On Error Resume Next
    varReturn = SysCmd(acSysCmdRemoveMeter)
    DoCmd.Close acForm, strProgForm, acSaveNo
    Set frm = Nothing
    rst.Close
    Set rst = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3135, 3061, 3078  ' Problem with SQL syntax, or ref to nonexistent object, etc.
        MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyLinks[mod_Linked_Tables])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyLinks[mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnMakeBackup
' Description:  Creates a backup of linked Access back-end database files
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   fxnCreateFolder, fxnFolderExists, fxnParsePath, fxnParseFileName,
'                   fxnParseFileExt, fxnSaveFile, fxnZipFiles, fxnFileExists, fxnPause
' Source/date:  Susan Huse, fall 2004
' Revisions:    John R. Boetsch, May 2005 - minor revisions and documentation
'               JRB, 5/16/2006 - updated documentation, error traps; modified date/time
'                   stamp to be appended to the file name; changed varCopyFile to a Variant
'                   to accommodate nulls from the procedure call
'               JRB, 1/8/2009 - streamlined to use zip files capability
'               JRB, 6/29/2009 - additional updates to accommodate multiple back-ends and
'                   revised system table structure
'               JRB, 10/8/2009 - inserted a pause in zip file creation to avoid closing
'                   before large back-end files are fully zipped
'               -------------------------------------------------------------------------
'               BLC, 4/30/2015 - moved to mod_Linked_Tables from mod_Custom_Functions
' =================================
Public Function fxnMakeBackup()
    On Error GoTo Err_Handler

    ' Prompt the user to confirm before backing up ... if no, exit function
    If MsgBox("Would you like to make a backup copy of the data?", vbYesNo, _
        "Create Backup?") = vbNo Then
        GoTo Exit_Procedure
    End If

    Dim rst As DAO.Recordset
    Dim intNRecs As Integer
    Dim strDbFile As String
    Dim fs As Variant
    Dim varCopyFile As Variant
    Dim arrFile() As String
    Dim strNewFile As String
    Dim strPath As String
    Dim strBackupDate As String
    Dim blnZipped As Boolean
    Dim strBackupFolder As String

    strBackupFolder = "Db_backups"
    strBackupDate = Format$(Now, "YYYYMMDD_HHNN")

    ' Set the recordset to the systems table, grouped by linked Access databases
    Set rst = CurrentDb.OpenRecordset("SELECT Database " & _
        "FROM MSysObjects " & _
        "WHERE ((MSysObjects.Type) = 6) And ((MSysObjects.Name) Not Like '~*') " & _
        "GROUP BY MSysObjects.Database;", dbOpenSnapshot)

    ' Counts the number of linked Access back-end files in the database
    rst.MoveLast    ' Need to do this to make the record count accurate
    intNRecs = rst.RecordCount
    If intNRecs = 0 Then    ' No linked databases in the recordset
        MsgBox "There are no Access back-end files to back up ...", , _
            "No back-end file to back up"
        GoTo Exit_Procedure
    End If

    ' Loop through the recordset and back up each file as indicated in the system file
    rst.MoveFirst
    Do Until rst.EOF
        strDbFile = rst![Database]
        ' If the string is not empty and backups are indicated for this back-end ...
        If strDbFile <> "" And _
            DLookup("[Backups]", "tsys_Link_Dbs", "[File_path]=""" & strDbFile & """") Then

            ' Remove the file name from the path
            strPath = fxnParsePath(strDbFile)
            ' Remove the right-most back slash if present
            If Right(strPath, 1) = "\" Then strPath = Left(strPath, Len(strPath) - 1)
            ' Update the backup folder string unless it is already the current folder
            arrFile() = Split(strPath, "\")
            If strBackupFolder <> arrFile(UBound(arrFile)) Then _
                strPath = strPath & "\" & strBackupFolder
            ' Verify the existence of the backup folder (and create it if needed)
            If fxnFolderExists(strPath) = False Then fxnCreateFolder (strPath)
            If fxnFolderExists(strPath) = False Then
                MsgBox "Unable to find/create the backup folder.", , "No Backup Made"
                GoTo Exit_Procedure
            End If
            ' Create the new file string by adding the current file name to the new path
            strNewFile = strPath & "\" & fxnParseFileName(strDbFile)
            ' Remove the current file extension
            strNewFile = Left(strNewFile, Len(strNewFile) - Len(fxnParseFileExt(strDbFile)))
            ' Append the backup date/time
            strNewFile = strNewFile & "_" & strBackupDate
            ' Zip the file to the new destination file name plus the ".zip" extension
            blnZipped = fxnZipFiles(strDbFile, strNewFile & ".zip")
            If blnZipped Then
                Dim intCounter As Integer
                intCounter = 0
                Call fxnPause(1000)
                Do While intCounter < 120
                    intCounter = intCounter + 1
                    If fxnFileExists(strNewFile & ".zip") Then
                        Exit Do
                    Else
                        ' Pause for 1000 ms before trying again
                        Call fxnPause(1000)
                    End If
                Loop
                MsgBox "Backup file successfully created: " & vbCrLf & vbCrLf & _
                    strNewFile & ".zip", vbOKOnly
            Else
                ' Zip operation unsuccessful, so try to make an outright copy
                ' Open the save file dialog and update to the actual name given by the user
                varCopyFile = fxnSaveFile(strNewFile, _
                    "Microsoft Access (*.mdb, *.accdb)", "*.mdb;*.accdb")
                If IsNull(varCopyFile) Then
                    ' User canceled save operation
                    MsgBox "No backup made", vbOKOnly
                Else
                    ' Perform the actual file copy
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    fs.CopyFile strDbFile, varCopyFile
                    MsgBox "Backup file successfully created: " & vbCrLf & vbCrLf & _
                        varCopyFile, vbOKOnly
                End If
            End If
            
        End If
        rst.MoveNext
    Loop    ' To next back-end

Exit_Procedure:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnMakeBackup)"
    End Select
    Resume Exit_Procedure

End Function