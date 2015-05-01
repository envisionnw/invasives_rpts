Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =11
    ItemSuffix =22
    Right =15975
    Bottom =11760
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xc1f3db6ed487e440
    End
    RecordSource ="tbl_Target_Areas"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    SplitFormSplitterBar =0
    SplitFormSplitterBar =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =4140
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =2355
                    Height =375
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblListSelectionHdr"
                    Caption ="Target List Selection"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =435
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4380
                    Top =3600
                    Width =2220
                    ForeColor =16711680
                    Name ="btnLoadList"
                    Caption ="Load List >>"
                    StatusBarText ="Continue to choose activities"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =3600
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =3960
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =6750156
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =52377
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =3407769
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =52224
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =2375487
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =6750156
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ListBox
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    Left =1320
                    Top =1260
                    Width =1320
                    Height =2100
                    ColumnOrder =0
                    TabIndex =1
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxParks"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Target_Species.Park_Code FROM tbl_Target_Species ORDER BY tb"
                        "l_Target_Species.Park_Code; "
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1320
                    LayoutCachedTop =1260
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =3360
                End
                Begin Label
                    OverlapFlags =85
                    Left =600
                    Top =720
                    Width =4875
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblInstructions"
                    Caption ="Choose the park(s) and year(s) you'd like to include."
                    GridlineColor =10921638
                    LayoutCachedLeft =600
                    LayoutCachedTop =720
                    LayoutCachedWidth =5475
                    LayoutCachedHeight =1035
                End
                Begin ListBox
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    Left =3540
                    Top =1260
                    Width =1320
                    Height =2100
                    ColumnOrder =1
                    TabIndex =2
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxYears"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Target_Species.Target_Year FROM tbl_Target_Species ORDER BY "
                        "tbl_Target_Species.Target_Year; "
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3540
                    LayoutCachedTop =1260
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =3360
                End
            End
        End
        Begin Section
            Height =0
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' MODULE:       Form_frmLoadList
' Description:  Load species list to target species list functions and routines
'
' Source/date:  Bonnie Campbell, 3/5/2015
' Revisions:    BLC - 3/5/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler

    'minimize the opening form
    Forms(Form.OpenArgs).SetFocus
    DoCmd.Minimize
    Me.SetFocus
    
    'save the form reference
    TempVars.Add "originForm", Form.OpenArgs
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[form_frmLoadList])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxParks_Click
' Description:  Determine selected parks
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Seth Schrock, March 12, 2013
' http://bytes.com/topic/access/answers/947721-taking-last-comma-off-text-string
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
' ---------------------------------
Private Sub lbxParks_Click()
On Error GoTo Err_Handler
Dim strParks As String, strComma As String
Dim item As Variant

    'determine the selected park(s)
    For Each item In lbxParks.ItemsSelected
        
        strParks = strParks & "'" & lbxParks.ItemData(item) & "',"

    Next
    
    'trim last comma
    strParks = IIf(Right(strParks, 1) = ",", Left(strParks, Len(strParks) - 1), strParks)
    
    TempVars.Add "parks", strParks
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxParks_Click[Form_frmSelectList])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxYears_Click
' Description:  Determine selected Years
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Seth Schrock, March 12, 2013
' http://bytes.com/topic/access/answers/947721-taking-last-comma-off-text-string
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
' ---------------------------------
Private Sub lbxYears_Click()
On Error GoTo Err_Handler
Dim strYears As String, strComma As String
Dim item As Variant

    'determine the selected year(s)
    For Each item In lbxYears.ItemsSelected
        
        strYears = strYears & lbxYears.ItemData(item) & ","
        
    Next
        
    'trim last comma
    strYears = IIf(Right(strYears, 1) = ",", Left(strYears, Len(strYears) - 1), strYears)
    
    TempVars.Add "years", strYears
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxYears_Click[Form_frmSelectList])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnLoadList_Click
' Description:  Load the target list species into frmTgtSpecies.lbxTgtSpecies
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
' ---------------------------------
Private Sub btnLoadList_Click()
On Error GoTo Err_Handler
    
    Dim strSQL As String, strWhere As String
    
    'determine the selected park(s) & year(s)
    If Len(TempVars.item("parks")) > 0 And Len(TempVars.item("years")) > 0 Then
        strWhere = "WHERE Park_Code IN (" & TempVars.item("parks") & ") " _
                 & "AND Target_Year IN (" & TempVars.item("years") & ")"
    End If
    
    'prep WHERE clause
    If Len(Replace(strWhere, "WHERE", "")) = 0 Then strWhere = ""
    
    'build SQL statement
    strSQL = "SELECT DISTINCT Master_Plant_Code_FK AS Code, Species_Name AS Species " _
            & "FROM tbl_Target_Species " _
            & strWhere & ";"
    
    'run search
    Dim rs As DAO.Recordset
    Dim rsNew As DAO.Recordset
      
    'fetch data
    Set rs = CurrentDb.OpenRecordset(strSQL)

    'merge existing listbox recordset w/ new SQL recordset
    Set rsNew = MergeRecordsets(Forms("frmTgtSpecies").lbxTgtSpecies.Recordset, rs)

    'load listbox
    PopulateList Forms("frmTgtSpecies").lbxTgtSpecies, rsNew, Forms("frmTgtSpecies").lbxTgtSpecies
    
    'remove dupes
    RemoveListDupes Forms("frmTgtSpecies").lbxTgtSpecies
    
    'cleanup
    TempVars.Remove ("parks")
    TempVars.Remove ("years")
    
    'return to species form
    Dim originForm As String
    
    originForm = Me.name
    
    'open species search form
    DoCmd.OpenForm "frmTgtSpecies", acNormal, , , , acWindowNormal, originForm
    
    'close & return to frmTgtSpecies
    If Forms("frmTgtSpecies").Minimized Then DoCmd.Restore
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLoadList_Click[Form_frmSelectList])"
    End Select
    Resume Exit_Sub
End Sub
