Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =9
    ItemSuffix =13
    Left =840
    Top =5835
    Right =8040
    Bottom =9195
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3d34192b53bbe340
    End
    Caption ="Species Cover"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =3600
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1320
                    Top =240
                    Width =4575
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="lblHeader"
                    Caption ="Species Cover by Route"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4440
                    Top =2580
                    Width =1350
                    Height =299
                    Name ="btnClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =2820
                    Top =1080
                    Width =2520
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Park_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Parks.ParkCode, tlu_Parks.ParkName FROM tlu_Parks; "
                    ColumnWidths ="0;2565"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1560
                            Top =1080
                            Width =1140
                            Height =245
                            FontWeight =700
                            Name ="Park_Code_Label"
                            Caption ="Select Park:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =720
                    Left =2820
                    Top =1680
                    Width =1200
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Visit_Year"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Distinct Visit_Year FROM Select_Cover_Year WHERE [Unit_Code] = 'DINO' ORD"
                        "ER BY Visit_Year"
                    ColumnWidths ="2820"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1380
                            Top =1680
                            Width =1320
                            Height =245
                            FontWeight =700
                            Name ="Plot_ID_Label"
                            Caption ="Select Year:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =2580
                    Width =1350
                    Height =299
                    TabIndex =3
                    Name ="btnReport"
                    Caption ="Create Table"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' MODULE:       frm_Species_Cover_by_Route
' Level:        Form module
' Version:      1.04
' Description:  File and directory related functions & subroutines
'
' Source/date:  Unknown
' Adapted:      Bonnie Campbell, May 2017
' Revisions:    Unknown        - 1.00 - initial version
'               BLC, 5/10/2017 - 1.01 - documentation, added Form_Open(), Visit_Year_AfterUpdate()
'               BLC, 6/15/2017 - 1.02 - revised to pull year from Select_Cover_Year vs qry_sel_cover_year
'               BLC, 6/20/2017 - 1.03 - cleared form fields after click
'               BLC, 6/22/2017 - 1.04 - revised to construct result from queries (Route_SpeciesCover_Crosstab_*,
'                                       where * = TCount, PctCover, SE)
' =================================

' ---------------------------------
'  Methods
' ---------------------------------

' ---------------------------------
' SUB:          Form_Open
' Description:  Form opening actions
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, May 2017 - initial version
' Adapted:      -
' Revisions:    BLC - 5/10/2017 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'initialize (year & create table button disabled until park selection)
    Me.Visit_Year.Enabled = False
    Me.btnReport.Enabled = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Species_Cover_by_Route form])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          Park_Code_AfterUpdate
' Description:  Sets park code/visit yeas filtering
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, May 2017 - initial version
' Revisions:    Unknown         - initial version
'               BLC - 5/10/2017 - added documentation, enabled Visit_Year
'               BLC - 6/15/2017 - revised to pull year from Select_Cover_Year vs qry_sel_cover_year
' ---------------------------------
Private Sub Park_Code_AfterUpdate()
On Error GoTo Err_Handler

  If Not IsNull(Me!Park_Code) Then
    Me!Visit_Year.Enabled = True
    Me!Visit_Year.RowSource = "SELECT Distinct Visit_Year FROM Select_Cover_Year WHERE [Unit_Code] = '" & Me!Park_Code & "' ORDER BY Visit_Year"
    Me.Refresh
  End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Park_Code_AfterUpdate[frm_Species_Cover_by_Route form])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          Visit_Year_AfterUpdate
' Description:  Sets park code/visit yeas filtering
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, May 2017 - initial version
' Adapted:      -
' Revisions:    BLC - 5/11/2017 - initial version
' ---------------------------------
Private Sub Visit_Year_AfterUpdate()
On Error GoTo Err_Handler

  If Not IsNull(Me!Visit_Year) Then
    Me!btnReport.Enabled = True
  End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Visit_Year_AfterUpdate[frm_Species_Cover_by_Route form])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          btnReport_Click
' Description:  Builds work table for species cover by route
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Russ DenBleyker - January 2010
' Adapted:      Bonnie Campbell, May 2017 - initial version
' Revisions:    RDB - 1/2010    - initial version
'               BLC - 5/10/2017 - added documentation, removed error message for Visit_Year (revised
'                                 to be disabled instead)
'               BLC - 6/20/2017 - cleared form fields after click
'               BLC - 6/22/2017 - revised to construct result from queries (Route_SpeciesCover_Crosstab_*,
'                                       where * = TCount, PctCover, SE)
' ---------------------------------
Private Sub btnReport_Click()
On Error GoTo Err_Handler

    If IsNull(Me!Park_Code) Or IsNull(Me!Visit_Year) Then
      MsgBox "You must select both park and year.", , "Monitoring Transect Data"
      Exit Sub
    End If

    'notify user system is busy
    DoCmd.Hourglass True
    Dim msg As String
    msg = "Generating " & Me.Park_Code & " " & Me.Visit_Year & " results..."
    SysCmd acSysCmdSetStatus, msg

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim tbl As DAO.TableDef
    Dim strFilter As String
    Dim strWHERE As String
    Dim origSQL As String
    Dim strSQL As String
    Dim strQuery As String
    Dim qry As Variant
    Dim strTableSQL As String
    Dim strComponent As String
    Dim strNewTable As String
    Dim strTable As String
    Dim strResult As String
    Dim strResultOld As String
    Dim ary As Variant
    
    'prepare filter clause
    strFilter = " WHERE sc.Unit_Code = '" & Me!Park_Code & _
                "' AND sc.Visit_Year = " & Me!Visit_Year & " "
    
    'prepare table name component
    strComponent = Me!Park_Code & "_" & Me!Visit_Year & "_"
    
    'prepare db
    Set db = CurrentDb
    
    'set array
    ary = Array("TCount", "PctCover", "SE")
    
    For Each qry In ary
        
        msg = "Generating " & Me.Park_Code & " " & Me.Visit_Year & " " & qry & "..."
        SysCmd acSysCmdSetStatus, msg
        
        'construct query, table names
        strQuery = "Route_SpeciesCover_Crosstab_" & qry
        strNewTable = strComponent & qry
        
        'use the existing queries, but filter based on Unit_Code & Visit_Year
        Set qdf = db.QueryDefs(strQuery)
        
        'save original SQL
        origSQL = qdf.SQL
        
        'remove any existing WHERE clause (in case user saved query w/ park & year filter)
        '& reset the query def to the original SQL w/o this filtering clause
        If InStr(origSQL, "WHERE") Then
            origSQL = ReplaceTextBetween(origSQL, "WHERE", "GROUP")
            qdf.SQL = origSQL
        End If
                
        'add filter
        strSQL = Replace(Replace(qdf.SQL, ";", ""), "GROUP", strFilter & "GROUP") & ";"
        qdf.SQL = strSQL
        
        'delete table if it exists
        If TableExists(strNewTable) Then _
            DoCmd.DeleteObject acTable, strNewTable
        
        'save query to table
        strTableSQL = "SELECT * INTO " & strComponent & qry & _
                       " FROM " & strQuery & ";"
        
        DoCmd.SetWarnings False
        db.Execute strTableSQL
        DoCmd.SetWarnings True
        
        'DoCmd.OpenQuery strQuery, acViewNormal, acReadOnly
        
        msg = "Combining " & Me.Park_Code & " " & Me.Visit_Year & " results ..."
        SysCmd acSysCmdSetStatus, msg
        
        'add the table to the overall species table
        strTable = strComponent & "SpeciesCover_by_Route"
        strResult = strTable & "_Result"
        strResultOld = strResult & "_OLD"
        
        If Not TableExists(strTable) Then
            'copy the table to new table
            DoCmd.CopyObject , strTable, acTable, strNewTable
        Else
            'delete only on the 1st (TCount) iteration
            If qry = "TCount" Then
                'remove existing prior results table of same name
                If TableExists(strResult) Then _
                    DoCmd.DeleteObject acTable, strResult
                'add table to existing table
                db.Execute CombineTableSQL(strTable, strNewTable, strResult)
            Else
                'remove any previous "OLD" table
                'If TableExists(strResultOld) Then DoCmd.DeleteObject acTable, strResultOld
                
                'rename strResult to strResultOld
                DoCmd.Rename strResultOld, acTable, strResult
                
                'add table to existing table
                db.Execute CombineTableSQL(strResultOld, strNewTable, strResult)
                
                'remove old result
                DoCmd.DeleteObject acTable, strResultOld
            End If
            
        End If
        
        'remove the filter (revert to original SQL) for the next iteration
        qdf.SQL = origSQL
        
    Next
                    
    'cleanup & clear memory
    Set ary = Nothing
    Set qdf = Nothing
    Set tbl = Nothing
    Set db = Nothing
    
    msg = "Shifting " & Me.Park_Code & " " & Me.Visit_Year & " results columns..."
    SysCmd acSysCmdSetStatus, msg
    
    'definitions
    Dim tdf As DAO.TableDef
    Dim col As DAO.field
    Dim strCol As String
    Dim strRoute As String
    Dim strPrevRoute As String
    Dim strTCount As String, strPctCover As String, strSE As String
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(strResult)
    strPrevRoute = ""
    
    With tdf
    'order result columns (fields)
    For Each col In tdf.Fields
    
        'get column name
        strCol = col.Name
        strRoute = Left(strCol, InStr(col.Name, ") ") + 1)
        
        'ignore 1-5 (static Park, Year, Species, Master Common Name, IsDead)
        If col.OrdinalPosition > 5 Then
        
            'if the routes differ we're onto another set to move
            If strRoute <> strPrevRoute Then
                strTCount = strRoute & "TCount"
                strPctCover = strRoute & "PctCover"
                strSE = strRoute & "SE"
                
                'move route's PctCover after TCount, SE after PctCover
                SetColumnOrdinalPosition tdf, strPctCover, strTCount
                SetColumnOrdinalPosition tdf, strSE, strPctCover
            
                strPrevRoute = strRoute
            End If
            
        End If
    Next
    End With
    
    msg = Me.Park_Code & " " & Me.Visit_Year & " results complete..."
    SysCmd acSysCmdSetStatus, msg
    
Exit_Procedure:
    'clear fields
    Me.Park_Code = ""
    Me.Visit_Year = ""

    'clear notification
    DoCmd.Hourglass False
    SysCmd acSysCmdClearStatus
    
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReport_Click[frm_Species_Cover_by_Route])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          btnClose_Click
' Description:  Form closing actions
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Russ DenBleyker - January 2010
' Adapted:      Bonnie Campbell, May 2017 - initial version
' Revisions:    RDB - 1/2010    - initial version
'               BLC - 5/10/2017 - added documentation
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Species_Cover_by_Route])"
    End Select
    Resume Exit_Procedure
End Sub

'Private Sub Visit_Year_Change()
' Debug.Print Me.Visit_Year.RowSource
'End Sub
'
'Private Sub Visit_Year_GotFocus()
' Debug.Print Me.Visit_Year.RowSource
'End Sub

'' ---------------------------------
'' SUB:          CollapseRows
'' Description:  Collapses TCount, PctCover, SE for one species/IsDead into one row
'' Parameters:   -
'' Returns:      -
'' Throws:       -
'' References:   -
''   R.Hicks, Sept 15, 2002
''   http://www.utteraccess.com/forum/copy-table-structure-vb-t117555.html
'' Source/date:  Bonnie Campbell, June 22 2017
'' Adapted:      -
'' Revisions:    BLC - 6/22/2017 - initial version
'' ---------------------------------
'Public Sub CollapseRows(tbl As TableDef)
'On Error GoTo Err_Handler
'
'    Dim db As DAO.Database
'    Dim tdf As DAO.TableDef
'    Dim col As field
'    Dim rs As DAO.Recordset
'    Dim rsPctCover As DAO.Recordset
'    Dim rsSE As DAO.Recordset
'    Dim strNewTable As String
'    Dim Park As String, VisitYear As String, Species As String, CommonName As String, _
'        IsDead As String, Route As String
'    Dim PrevPark As String, PrevVisitYear As String, PrevSpecies As String, _
'        PrevCommonName As String, PrevIsDead As String, PrevRoute As String
'    Dim TCount As Integer
'    Dim PctCover As String
'    Dim SE As Double
'
'    strNewTable = tbl.Name & "_NEW"
'
'    PrevPark = ""
'    PrevVisitYear = ""
'    PrevSpecies = ""
'    PrevCommonName = ""
'    PrevIsDead = ""
'
'    Set db = CurrentDb
'    Set tdf = db.TableDefs(tdf)
'
'    Set rs = db.OpenRecordset(tbl)
'
'    'create empty table w/ same columns to fill into
'    DoCmd.TransferDatabase acExport, "Microsoft Access", db.Name, acTable, tbl.Name, strTableNew, True
'
'    With tdf
'        'iterate through ALL
'        Do Until rs.BOF And rs.EOF
'
'            'order result columns (fields)
'            For Each col In tdf.Fields
'
'                'get park, visit year, species, common name & isdead
'                If col.OrdinalPosition = 1 Then Park = col.Value
'                If col.OrdinalPosition = 2 Then VisitYear = col.Value
'                If col.OrdinalPosition = 3 Then Species = col.Value
'                If col.OrdinalPosition = 4 Then CommonName = col.Value
'                If col.OrdinalPosition = 5 Then IsDead = col.Value
'
'                'ignore 1-5 (static Park, Year, Species, Master Common Name, IsDead)
'                If col.OrdinalPosition > 5 Then
'
'                    'get column & route name
'                    strCol = col.Name
'                    Route = Left(strCol, InStr(col.Name, ") ") + 1)
'
'                    Select Case Replace(strCol, strRoute)
'                        Case "TCount"
'                            TCount = col.Value
'                        Case "AvgCover"
'                            AvgCover = col.Value
'                        Case "SE"
'                            SE = col.Value
'                    End Select
'
'                    'concatenate for comparison
'                    Concat = Park & VisitYear & Species & CommonName & IsDead & Route
'
'                    If Concat <> PrevConcat Then
'
'                    'add these to the new table if they don't already exist
'
'
'                    If PrevSpecies = rs!Species And PrevIsDead = rs!IsDead Then
'
'                    End If
'                End If
'
'                'capture the previous values
'                PrevConcat = Concat
'                PrevPark = Park
'                PrevVisitYear = VisitYear
'                PrevSpecies = Species
'                PrevCommonName = CommonName
'                PrevIsDead = IsDead
'
'            Next
'
'
'        Loop
'    End With
'
'Exit_Procedure:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - CollapseRows[frm_Species_Cover_by_Route])"
'    End Select
'    Resume Exit_Procedure
'End Sub
