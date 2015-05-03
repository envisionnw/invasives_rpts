Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10935
    DatasheetFontHeight =11
    ItemSuffix =19
    Right =20208
    Bottom =9408
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x72574db34b86e440
    End
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =11220
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =8340
                    Top =5220
                    Width =1618
                    Height =373
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =29
                    RightMargin =29
                    BottomMargin =29
                    BackColor =6750105
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblSaveList"
                    Caption ="Save List"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Print sheets"
                    GridlineColor =10921638
                    LayoutCachedLeft =8340
                    LayoutCachedTop =5220
                    LayoutCachedWidth =9958
                    LayoutCachedHeight =5593
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =360
                    Top =1080
                    Width =3960
                    Height =4032
                    FontSize =10
                    BoundColumn =2
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxSpecies"
                    RowSourceType ="Value List"
                    RowSource ="Park;File_Code;Datasheet;File_Description;Sort_Order;CANY;Photo;CANY_Big_Rivers_"
                        "Photo_Data_Sheet_bc20150128a.pdf;photo data sheet;1;CANY;VegPlot;CANY_Big_Rivers"
                        "_Veg_Plots_Data_Sheet_bc20150128a.pdf;veg plots data sheet;2;CANY;VegCont;CANY_B"
                        "ig_Rivers_Veg_Plots_Continuation_Data_Sheet_bc20150128a.pdf;veg continuation dat"
                        "a sheet;3;CANY;VegWalk;CANY_Big_Rivers_Veg_Walk_Data_Sheet_bc20150128a.pdf;veg w"
                        "alk data sheet;4"
                    ColumnWidths ="1440;2520;14"
                    OnDblClick ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =5112
                    DatasheetCaption ="Species"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =720
                            Width =1260
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSpecies"
                            Caption ="Species"
                            ControlTipText ="Species within park"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =720
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =1035
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5760
                    Top =1080
                    Width =3960
                    Height =4032
                    FontSize =10
                    TabIndex =1
                    BoundColumn =2
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxTgtSpecies"
                    RowSourceType ="Value List"
                    RowSource ="Park;File_Code;Datasheet;File_Description;Sort_Order"
                    ColumnWidths ="1440;2520;14"
                    OnDblClick ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Target species"
                    GridlineColor =10921638

                    LayoutCachedLeft =5760
                    LayoutCachedTop =1080
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =5112
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =720
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblTgtSpecies"
                            Caption ="Target Species"
                            GridlineColor =10921638
                            LayoutCachedLeft =5580
                            LayoutCachedTop =720
                            LayoutCachedWidth =7020
                            LayoutCachedHeight =1035
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =3780
                    Width =493
                    Height =463
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =101
                    RightMargin =29
                    BottomMargin =29
                    BackColor =13882323
                    BorderColor =8355711
                    ForeColor =8224125
                    Name ="lblRemove"
                    Caption ="<"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Remove selected"
                    GridlineColor =10921638
                    LayoutCachedLeft =4757
                    LayoutCachedTop =3780
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =4243
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =1200
                    Width =493
                    Height =463
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =101
                    RightMargin =29
                    BottomMargin =29
                    BackColor =6750105
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblAddAll"
                    Caption =">>"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add all"
                    GridlineColor =10921638
                    LayoutCachedLeft =4757
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =1663
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =4560
                    Width =493
                    Height =463
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =101
                    RightMargin =29
                    BottomMargin =29
                    BackColor =52479
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblRemoveAll"
                    Caption ="<<"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Remove all"
                    GridlineColor =10921638
                    LayoutCachedLeft =4757
                    LayoutCachedTop =4560
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =5023
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =2040
                    Width =493
                    Height =463
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =101
                    RightMargin =29
                    BottomMargin =29
                    BackColor =13882323
                    BorderColor =8355711
                    ForeColor =8224125
                    Name ="lblAdd"
                    Caption =">"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add selected"
                    GridlineColor =10921638
                    LayoutCachedLeft =4757
                    LayoutCachedTop =2040
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =2503
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4410
                    Top =5220
                    Width =1264
                    Height =358
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =29
                    RightMargin =29
                    BottomMargin =29
                    BackColor =6750105
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblReset"
                    Caption ="Reset Lists"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Reset lists to their original state"
                    GridlineColor =10921638
                    LayoutCachedLeft =4410
                    LayoutCachedTop =5220
                    LayoutCachedWidth =5674
                    LayoutCachedHeight =5578
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =1425
                    Top =5220
                    Width =1468
                    Height =373
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =29
                    RightMargin =29
                    BottomMargin =29
                    BackColor =6750105
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblSearch"
                    Caption ="Find Species"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find a species..."
                    GridlineColor =10921638
                    LayoutCachedLeft =1425
                    LayoutCachedTop =5220
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =5593
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =840
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblParkHdr"
                    Caption ="Park"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =432
                End
                Begin Subform
                    OverlapFlags =85
                    Left =480
                    Top =6780
                    Width =3960
                    Height =4032
                    TabIndex =2
                    BorderColor =10921638
                    Name ="sfrmSpeciesListbox"
                    SourceObject ="Form.sfrmSpeciesListbox"
                    GridlineColor =10921638

                    LayoutCachedLeft =480
                    LayoutCachedTop =6780
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =10812
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =6360
                            Width =1200
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSpeciesListbox"
                            Caption ="Species"
                            GridlineColor =10921638
                            LayoutCachedLeft =240
                            LayoutCachedTop =6360
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =6675
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =2940
                    Top =780
                    Width =1440
                    Height =276
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesCount"
                    Caption ="species"
                    ControlTipText ="Number of species in the current list"
                    GridlineColor =10921638
                    LayoutCachedLeft =2940
                    LayoutCachedTop =780
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1056
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =8340
                    Top =780
                    Width =1440
                    Height =276
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTgtSpeciesCount"
                    Caption ="species"
                    ControlTipText ="Number of species in the current list"
                    GridlineColor =10921638
                    LayoutCachedLeft =8340
                    LayoutCachedTop =780
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =1056
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =3060
                    Top =6480
                    Width =1440
                    Height =276
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSfrmSpeciesCount"
                    Caption ="species"
                    ControlTipText ="Number of species in the current list"
                    GridlineColor =10921638
                    LayoutCachedLeft =3060
                    LayoutCachedTop =6480
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =6756
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
' MODULE:       Form_frmTgtSpecies
' Description:  Species selction functions & procedures
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC - 2/9/2015 - initial version
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
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    
    Initialize
    
    'prep headers
    lblParkHdr.Caption = TempVars.item("park")
    lblSpecies.Caption = TempVars.item("state") & " Species"
    lblSpeciesListbox.Caption = TempVars.item("state") & " Species"
    
    'clear headers
    lbxSpecies.RowSource = ""
    lbxTgtSpecies.RowSource = ""
    
    'initial listbox fill
    fillList Me, lbxSpecies, lbxTgtSpecies  '.RowSourceType = "Value List"

    'Enable move items lbls (or not)
    If lbxSpecies.ListCount > 0 Then
        lblAddAll.Visible = True
        lblRemoveAll.Visible = True
    End If
    
    'Set counts
    lblSpeciesCount.Caption = lbxSpecies.ListCount & " species"
    lblTgtSpeciesCount.Caption = lbxTgtSpecies.ListCount & " species"
'    lblLbxSpeciesCount.Caption = Count(sfrm.ListCount) & " species"

    'enable > or < *only* if at least one item selected
    'check background / text color, if gray then exit_sub
'Forms!MasterForm!Label1550.SpecialEffect = vbraised
'Forms!MasterForm!Label1550.SpecialEffect = vbetched
'Forms!MasterForm!Label1550.SpecialEffect = vbsunken
    
    DisableControl lblAdd
    DisableControl lblRemove
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblReset_Click
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
Private Sub lblReset_Click()
On Error GoTo Err_Handler

    'go back to initial state
    Form_Load

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblReset_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxSpecies_Click
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
Private Sub lbxSpecies_Click()
On Error GoTo Err_Handler

    Dim varItem As Variant

    'deselect items in target control (lbxTgtSpecies)
    For Each varItem In lbxTgtSpecies.ItemsSelected
        lbxTgtSpecies.Selected(varItem) = False
    Next

    'check for selected items --> if present, enable lblAdd
    If lbxSpecies.ItemsSelected.count > 0 Then
        If lblAdd.backcolor <> TempVars.item("ctrlAddEnabled") Then
            EnableControl lblAdd, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
        End If
    Else
        DisableControl lblAdd
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxSpecies_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxSpecies_DblClick
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Public Sub lbxSpecies_DblClick(Cancel As Integer)
On Error GoTo Err_Handler:

    MoveSingleItem Me, "lbxSpecies", "lbxTgtSpecies"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxSpecies_DblClick[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxSpecies_KeyUp
' Description:  Deselects items after control update
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Allen Browne, June 2006
' http://allenbrowne.com/func-12.html
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Private Sub lbxSpecies_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

    If lbxSpecies.ItemsSelected.count > 0 And lblAdd.backcolor <> TempVars.item("ctrlAddEnabled") Then
        EnableControl lblAdd, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxSpecies_KeyUp[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Allen Browne, June 2006
' http://allenbrowne.com/func-12.html
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Private Sub lbxTgtSpecies_Click()
On Error GoTo Err_Handler
    
    Dim varItem As Variant
    
    'deselect items in source control (lbxSpecies)
    For Each varItem In lbxSpecies.ItemsSelected
        lbxSpecies.Selected(varItem) = False
    Next

    'check for selected items --> if present, enable lblRemove
    If lbxTgtSpecies.ItemsSelected.count > 0 Then
        If lblRemove.backcolor <> TempVars.item("ctrlRemoveEnabled") Then
            EnableControl lblRemove, TempVars.item("ctrlRemoveEnabled"), TempVars.item("textEnabled")
        End If
    Else
        DisableControl lblRemove
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_DblClick
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub lbxTgtSpecies_DblClick(Cancel As Integer)
    
On Error GoTo Err_Handler

    MoveSingleItem Me, "lbxTgtSpecies", "lbxSpecies"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_DblClick[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_KeyUp
' Description:  Deselects items after control update
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
Private Sub lbxTgtSpecies_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

    If lbxSpecies.ItemsSelected.count > 0 And lblRemove.backcolor <> TempVars.item("ctrlRemoveEnabled") Then
        EnableControl lblRemove, TempVars.item("ctrlRemoveEnabled"), TempVars.item("textEnabled")
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_KeyUp[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblAdd_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub lblAdd_Click()
On Error GoTo Err_Handler
    
    'ignore if 'disabled'
    If lblAdd.backcolor = lngGray Then GoTo Exit_Sub
    
    MoveSingleItem Me, "lbxSpecies", "lbxTgtSpecies"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblAdd_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblRemove_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub lblRemove_Click()
On Error GoTo Err_Handler
    
    'ignore if 'disabled'
    If lblRemove.backcolor = TempVars.item("ctrlDisabled") Then GoTo Exit_Sub
    
    MoveSingleItem Me, "lbxTgtSpecies", "lbxSpecies"
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblRemove_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblAddAll_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub lblAddAll_Click()
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'fetch recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(TempVars.item("strSQL"))
    
    MoveAllItems Me, "lbxSpecies", "lbxTgtSpecies"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblAddAll_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblRemoveAll_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub lblRemoveAll_Click()
On Error GoTo Err_Handler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'fetch recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(TempVars.item("strSQL"))
    
    MoveAllItems Me, "lbxTgtSpecies", "lbxSpecies"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblRemoveAll_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblSaveList_Click
' Description:  Save list items
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
Private Sub lblSaveList_Click()
On Error GoTo Err_Handler

    Dim iRow As Integer
    Dim strSpecies As String, strSQL As String
    
    strSQL = "INSERT INTO tbl_Target_Species" _
            & "(Master_Plant_Code_FK, Park_Code, Target_Year, Species_Name)" _
            & "VALUES "
    
    For iRow = 0 To lbxTgtSpecies.ListCount - 1
       
       ' ---------------------------------------------------
       '  NOTE: listbox column MUST have a non-zero width to retrieve its value
       ' ---------------------------------------------------
        strSpecies = lbxTgtSpecies.Column(0, iRow) 'column 0 = Master_PLANT_Code
        
        'set statusbar notice
        Dim varReturn As Variant
        varReturn = SysCmd(acSysCmdSetStatus, "Saving " & strSpecies & "...")
        
        'save it
        
    
    Next

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblSaveList_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblSearch_Click
' Description:  Opens species search to find species for populating target list
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015 - initial version
' ---------------------------------
Private Sub lblSearch_Click()
On Error GoTo Err_Handler
    Dim originForm As String
    
    originForm = Me.name
    
    'open species search form
    DoCmd.OpenForm "frmSpeciesSearch", acNormal, , , , , originForm

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblSearch_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub
