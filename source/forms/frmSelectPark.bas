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
    ItemSuffix =14
    Right =15720
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
            Height =2460
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =1632
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblParkSelectionHdr"
                    Caption ="Park Selection"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1692
                    LayoutCachedHeight =432
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4380
                    Top =1740
                    Width =2220
                    ForeColor =16711680
                    Name ="btnContinue"
                    Caption ="Continue >>"
                    StatusBarText ="Continue to choose activities"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =1740
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =2100
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
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5040
                    Left =1140
                    Top =1260
                    Width =3480
                    Height =300
                    ColumnOrder =0
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"8\""
                    Name ="cbxPark"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [tlu_Parks].[ParkCode], [tlu_Parks].[ParkName] FROM tlu_Parks ORDER BY [P"
                        "arkName]; "
                    ColumnWidths ="1080;3960"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="Choose a park."
                    GridlineColor =10921638

                    LayoutCachedLeft =1140
                    LayoutCachedTop =1260
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =600
                            Top =720
                            Width =3468
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPark"
                            Caption ="Choose the park you wish to work on."
                            GridlineColor =10921638
                            LayoutCachedLeft =600
                            LayoutCachedTop =720
                            LayoutCachedWidth =4068
                            LayoutCachedHeight =1020
                        End
                    End
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
            Height =492
            Name ="FormFooter"
            AutoHeight =1
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
' MODULE:       Form_frmTgtSpeciesList
' Description:  Target species functions & procedures
'
' Source/date:  Bonnie Campbell, 2/11/2015
' Revisions:    BLC - 2/11/2015 - initial version
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
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    
    Initialize

    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[form_frmTgtAreas])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxPark_Change
' Description:  Actions to take when a park is selected
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Private Sub cbxPark_Change()
On Error GoTo Err_Handler

    If cbxPark.ItemsSelected.count > 0 Then
        'set park
        TempVars.item("park") = cbxPark.ItemsSelected.item(0)
        
        'enable the continue button
        EnableControl btnContinue, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
        btnContinue.Enabled = True
    'Else
        'disable the continue button"
     '   btnContinue.Enabled = False
     '   DisableControl btnContinue
    End If
    
    If Len(cbxPark.Value) > 0 Then
        'set park
        TempVars.item("park") = cbxPark.Value
        'enable the continue button
        EnableControl btnContinue, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
        btnContinue.Enabled = True
    Else
        'disable the continue button"
        btnContinue.Enabled = False
        DisableControl btnContinue
    End If
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPark_Change[Form_frmSelectPark])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnContinue_Click
' Description:  Continue to park actions
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Private Sub btnContinue_Click()
On Error GoTo Err_Handler
    
    'open activity list
    DoCmd.OpenForm "frmActions", acNormal
        
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnContinue_Click[Form_frmSelectPark])"
    End Select
    Resume Exit_Sub
End Sub
