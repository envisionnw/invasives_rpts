Option Compare Database
Option Explicit
' =================================
' MODULE:       mod_dbTemplates
' Description:  Template related functions & subroutines pulling templates from the database
'
' Source/date:  Bonnie Campbell, May 2014
' Revisions:    BLC, 6/13/2014 - initial version
' =================================

' ---------------------------------
' SUB:     GetSQLTemplate
' Description:  loads SQL templates (queries as SQL string) into memory as a dictionary object
'               with query SQL strings available without querying the db tsys_SQL_templates table
' Parameters:
' Returns:      dictionary object stored in tempVars.Item("SQL")
' Assumptions:  placing
' Throws:       none
' References:   tsys_SQL_templates, Microsoft Scripting Runtime (dictionary object)
' Source/date:  Bonnie Campbell, June 2014
' Revisions:    BLC, 6/16/2014 - XX
' ---------------------------------
Public Sub GetSQLTemplates(Optional strVersion As String = "")

    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String, strSQLWhere As String, key As String, Value As String
    
    'handle default
    strSQLWhere = " WHERE Is_Supported > 0"
    
    If Len(strVersion) > 0 Then
        strSQLWhere = " AND LCase(versionID) = LCase(" & strVersion & " )"
    End If
    
    'sql
    strSQL = "SELECT * FROM tsys_Db_Templates" & strSQLWhere
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset(strSQL)
    
    'handle no records
    If rst.EOF Then
        MsgBox "Sorry, no templates were found for this database version.", vbExclamation, _
            "Linked Database Templates Not Found"
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If
    
    'prepare dictionary
    Dim dict As New Scripting.Dictionary
    Dim ary(1 To 4) As String
    Dim i As Integer
    
    'prepare the dictionary key array
    ary(1) = "context"
    ary(2) = "template_Name"
    ary(3) = "SQLstring" 'template
    ary(4) = "var_list"
    
    rst.MoveFirst
    Do Until rst.EOF
        'populate the dictionary
        For i = 1 To UBound(ary)
            key = ary(i)
            If (ary(i) = "SQLstring") Then
                Value = rst!template
            Else
                Value = rst.Fields(ary(i))
            End If
            If Not dict.Exists(key) Then
                dict.Add key, Value
            End If
        Next
        rst.MoveNext
    Loop
    
    TempVars.Add "SQL", dict

    'cleanup
    Set dict = Nothing
    
Exit_Procedure:
End Sub

' ---------------------------------
' SUB:     GetDbQuerySQL
' Description:  gets SQL from existing database queries via QueryDef object
' Parameters:   strQueryName - name of the Access query
' Returns:      SQL string
' Assumptions:  -
' Throws:       none
' References:   -
' Source/date:  sphinney, 7/13/2009 comment on
'               http://bytes.com/topic/access/answers/871500-getting-sql-string-query
' Adapted:      Bonnie Campbell, June 2014
' Revisions:    BLC, 6/16/2014 - XX
' ---------------------------------
Private Function GetDbQuerySQL(strQueryName As String) As String
Dim QD As DAO.QueryDef
 
Set QD = CurrentDb.QueryDefs(strQueryName)
GetDbQuerySQL = QD.sql
 
End Function

' ---------------------------------
' SUB:          qryExists
' Description:  Checks if query exists in database as a permanent query(QueryDefs)
' Parameters:   strQueryName - query name as a string
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Nick Vans, January 31, 2008
'               http://bytes.com/topic/access/answers/765384-determine-if-query-x-exists
' Adapted:      Bonnie Campbell, June 17, 2014
' Revisions:    6/17/2014 - BLC - XX
' ---------------------------------
Public Function qryExists(strQueryName As String) As Boolean

    Dim qdf As DAO.QueryDef
    
    'default
    qryExists = False
  
    For Each qdf In CurrentDb.QueryDefs
'        Debug.Print qdf.Name
        If qdf.name = strQueryName Then
            qryExists = True
            Exit For
        End If
    Next
End Function