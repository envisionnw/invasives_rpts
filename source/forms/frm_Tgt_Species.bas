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
    ItemSuffix =28
    Left =4104
    Top =300
    Right =15288
    Bottom =6348
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x72574db34b86e440
    End
    Caption ="Create Target Species List"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnActivate ="[Event Procedure]"
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
            Height =6060
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =5760
                    Top =1080
                    Width =4320
                    Height =4032
                    FontSize =10
                    BoundColumn =2
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxTgtSpecies"
                    RowSourceType ="Value List"
                    RowSource ="Code;Species;;;ABRONIA;Abronia sp.;;;ACAMPTOPAPPUS;Acamptopappus sp.;;;ARTARB;Ar"
                        "temisia arbuscula;;;ERICOMPOSITUS;Erigeron compositus;;;STIARI;Achnatherum aridu"
                        "m;;;ACAMPTOPAPPUS;Acamptopappus sp.;ACAMP;;ERICOMPOSITUS;Erigeron compositus;ERC"
                        "O4;;ERICER;Eriogonum cernuum;ERCE;"
                    ColumnWidths ="1440;2520;720;288;288"
                    OnDblClick ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Target species"
                    GridlineColor =10921638

                    LayoutCachedLeft =5760
                    LayoutCachedTop =1080
                    LayoutCachedWidth =10080
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
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =840
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblParkHdr"
                    Caption ="CURE"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =432
                End
                Begin Subform
                    OverlapFlags =85
                    Left =420
                    Top =1080
                    Width =3960
                    Height =4032
                    TabIndex =1
                    BorderColor =10921638
                    Name ="fsub_Species_Listbox"
                    SourceObject ="Form.fsub_Species_Listbox"
                    GridlineColor =10921638

                    LayoutCachedLeft =420
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =5112
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =660
                            Width =1200
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSpeciesListbox"
                            Caption ="CO Species"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =660
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =975
                        End
                    End
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
                    Caption ="8 species"
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
                    Left =3000
                    Top =780
                    Width =1440
                    Height =276
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSfrmSpeciesCount"
                    Caption ="3195 species"
                    ControlTipText ="Number of species in the current list"
                    GridlineColor =10921638
                    LayoutCachedLeft =3000
                    LayoutCachedTop =780
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =1056
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =2880
                    Top =120
                    Width =4320
                    Height =315
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear"
                    Caption ="Target Species List for "
                    GridlineColor =10921638
                    LayoutCachedLeft =2880
                    LayoutCachedTop =120
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =435
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7560
                    Top =180
                    Width =1560
                    Height =405
                    TabIndex =2
                    ForeColor =16711680
                    Name ="btnLoad"
                    Caption ="Load List"
                    StatusBarText ="Continue to choose activities"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =7560
                    LayoutCachedTop =180
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =585
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
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8400
                    Top =5220
                    Width =1560
                    Height =405
                    TabIndex =3
                    ForeColor =16711680
                    Name ="btnSaveList"
                    Caption ="Save List"
                    StatusBarText ="Save the current list"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =8400
                    LayoutCachedTop =5220
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =5625
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
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1380
                    Top =5220
                    Width =1560
                    Height =405
                    TabIndex =4
                    ForeColor =16711680
                    Name ="btnSearch"
                    Caption ="Find Species"
                    StatusBarText ="Find a species..."
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1380
                    LayoutCachedTop =5220
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =5625
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
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4380
                    Top =5220
                    Width =1320
                    Height =405
                    TabIndex =5
                    ForeColor =16711680
                    Name ="btnReset"
                    Caption ="Reset List"
                    StatusBarText ="Reset lists to their original state"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =5220
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =5625
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
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =4740
                    Top =1200
                    Width =540
                    Height =480
                    FontWeight =600
                    TabIndex =6
                    ForeColor =8224125
                    Name ="btnAddAll"
                    Caption =">>"
                    StatusBarText ="Reset lists to their original state"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =1680
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =13882323
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =8224125
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
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =4560
                    Width =540
                    Height =480
                    FontWeight =600
                    TabIndex =7
                    ForeColor =16711680
                    Name ="btnRemoveAll"
                    Caption ="<<"
                    StatusBarText ="Remove all"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =4560
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =5040
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =52479
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =0
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =3960
                    Width =540
                    Height =480
                    FontWeight =600
                    TabIndex =8
                    ForeColor =16711680
                    Name ="btnRemove"
                    Caption ="<"
                    StatusBarText ="Remove all"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Remove selected"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =3960
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =4440
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =52479
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderColor =0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =0
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =3
                    WebImagePaddingBottom =3
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =1860
                    Width =540
                    Height =480
                    FontWeight =600
                    TabIndex =9
                    ForeColor =16711680
                    Name ="btnAdd"
                    Caption =">"
                    StatusBarText ="Remove all"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add selected"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =1860
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =2340
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =6750156
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
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
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =3
                    WebImagePaddingBottom =3
                    Overlaps =1
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
' MODULE:       Form_frm_Tgt_Species
' Description:  Species selction functions & procedures
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC, 2/9/2015 - initial version
'               BLC, 4/30/2015 - integrated into Invasives Reporting tool
' =================================

'=================================================================
'  Properties
'=================================================================
' ---------------------------------
' PROPERTY:     Maximized
' Description:  Indicates if form is maximized or not by checking IsZoomed()
' Assumptions:  none
' Parameters:   N/A
' Returns:      True(1) - form is maximized
'               False(0) - form is not maximized
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Get Maximized() As Boolean
     Maximized = IsZoomed(Me.hwnd) * 1
End Property

' ---------------------------------
' PROPERTY:     Minimized
' Description:  Indicates if form is minimized or not by checking IsIconic()
' Assumptions:  none
' Parameters:   N/A
' Returns:      True(1) - form is minimized
'               False(0) - form is not minimized
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Get Minimized() As Boolean
     Minimized = IsIconic(Me.hwnd) * 1
End Property

' ---------------------------------
' PROPERTY LET: Maximized
' Description:  Sets custom form property 'Maximized'
' Assumptions:
' Note:         The IsMax argument must be defined as the same data type
'               returned by the corresponding Property Get procedure for
'               the same custom property.
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Let Maximized(IsMax As Boolean)
     If IsMax Then
         Me.SetFocus
         DoCmd.Maximize
     Else
         Me.SetFocus
         DoCmd.Restore
     End If
End Property

' ---------------------------------
' PROPERTY LET: Minimized
' Description:  Sets custom form property 'Minimized'
' Assumptions:
' Note:         The IsMin argument must be defined as the same data type
'               returned by the corresponding Property Get procedure for
'               the same custom property.
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Let Minimized(IsMin As Boolean)
     If IsMin Then
         Me.SetFocus
         DoCmd.Minimize
     Else
         Me.SetFocus
         DoCmd.Restore
     End If
End Property

'Public WithEvents sfrm As Access.Form

'=================================================================
'  Subroutines & Functions
'=================================================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/9/2015 - initial version
'   BLC, 5/1/2015 - integrated into Invasives Reporting tool, removed frmSelectYear closure since that form
'                   is no longer needed, added check for species number to ensure >= 0
'   BLC, 5/13/2015 - disabled Remove All button to start & recaptioned btnReset to "Reset List" vs. "Reset Lists"
'                    set btnAdd to enabled to start vs disabled
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    Dim intSpecies As Integer
    
    Initialize
       
    'set state
    TempVars.item("state") = getParkState(TempVars.item("park"))
    
    'set year
    TempVars.item("TgtYear") = Form.OpenArgs
    
    'prep headers
    lblParkHdr.Caption = TempVars.item("park")
    lblYear.Caption = "Target Species List for " & Form.OpenArgs
    lblSpeciesListbox.Caption = TempVars.item("state") & " Species"
    
    'clear headers
    lbxTgtSpecies.RowSource = ""
    
    'initial listbox fill
     fillList Me, lbxTgtSpecies

    'Enable move items lbls (or not)
    btnAddAll.Enabled = False
    
    'Set counts
    lblTgtSpeciesCount.Caption = GetListCount(lbxTgtSpecies, True) & " species"
    
    'Set Reset button label (reset list vs. lists)
    btnReset.Caption = "Reset List"
    
    DisableControl btnAddAll
    EnableControl btnAdd, lngLtLime, lngBlue, lngDkLime, lngBrtLime, lngLtGreen, lngDkGray, lngLtLime
    DisableControl btnRemove
    DisableControl btnRemoveAll
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Form_Activate
' Description:  Sets tbxTgtSpecies value
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
'   BLC - 5/1/2015 - added check for no species to prevent # = -1
'   BLC - 5/10/2015 - revised to include generic count function
' ---------------------------------
Private Sub Form_Activate()

On Error GoTo Err_Handler
    
    'set species count
    lblTgtSpeciesCount.Caption = GetListCount(lbxTgtSpecies, True) & " species"
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Activate[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnLoad_Click
' Description:  Load list from previous year
' Assumptions:  -
' Parameters:   none
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC, 3/5/2015 - initial version
'   BLC, 5/1/2015 - updated frmSelectList to frm_Select_List to conform to standards
' ---------------------------------
Private Sub btnLoad_Click()

On Error GoTo Err_Handler

    GetListboxRecordset lbxTgtSpecies



    'open tgt species list form
    DoCmd.OpenForm "frm_Select_List", acNormal, , , , acWindowNormal, Me.name

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLoad_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnReset_Click
' Description:  Reset lists to their original state
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
' ---------------------------------
Private Sub btnReset_Click()
On Error GoTo Err_Handler

    'go back to initial state
    Form_Load

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReset_Click[Form_frm_Tgt_Species])"
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
'   BLC - 5/10/2015 - added species count update
'   BLC - 5/13/2015 - revised to use global constants vs. tempvars for enabled control
' ---------------------------------
Private Sub lbxTgtSpecies_Click()
On Error GoTo Err_Handler
    
    Dim varItem As Variant
    
   'check for selected items --> if present, enable btnRemove
    If lbxTgtSpecies.ItemsSelected.count > 0 Then
        If btnRemove.backcolor <> CTRL_REMOVE_ENABLED Then
            EnableControl btnRemove, CTRL_REMOVE_ENABLED, TEXT_ENABLED
            EnableControl btnRemoveAll, CTRL_REMOVE_ENABLED, TEXT_ENABLED
        End If
    Else
        DisableControl btnRemove
    End If
    
    'set species count
    lblTgtSpeciesCount.Caption = GetListCount(lbxTgtSpecies, True) & " species"
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_DblClick
' Description:  Removes items from lbxTgSpecies
' Assumptions:  -
' Parameters:   Cancel - if true cancels action, false runs removal
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
'   BLC - 5/10/2015 - changed from MoveSingleItem to RemoveSelectedItems to handle
'                     removing species versus populating them back to the original species list
'                     added count update
' ---------------------------------
Private Sub lbxTgtSpecies_DblClick(Cancel As Integer)
    
On Error GoTo Err_Handler

    'MoveSingleItem Me, "lbxTgtSpecies", "lbxTgtSpecies"
    RemoveSelectedItems lbxTgtSpecies

    'set species count
    lblTgtSpeciesCount.Caption = GetListCount(lbxTgtSpecies, True) & " species"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_DblClick[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_KeyUp
' Description:  Deselects items after control update
' Assumptions:  -
' Parameters:   KeyCode - keystroke code
'               Shift - if shift key has been pressed
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 5/13/2015 - revised to use global constants vs. tempvars for enabled control
' ---------------------------------
Private Sub lbxTgtSpecies_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

'    If lbxSpecies.ItemsSelected.Count > 0 And lblRemove.backcolor <> TempVars.item("ctrlRemoveEnabled") Then
    If btnRemove.backcolor <> CTRL_REMOVE_ENABLED Then
        EnableControl btnRemove, CTRL_REMOVE_ENABLED, TEXT_ENABLED
        EnableControl btnRemoveAll, CTRL_REMOVE_ENABLED, TEXT_ENABLED
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_KeyUp[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnAdd_Click
' Description:  Add selected items to list
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
'   BLC - 5/10/2015 - added update for species count
' ---------------------------------
Private Sub btnAdd_Click()
On Error GoTo Err_Handler
    
    'ignore if 'disabled'
    'If lblAdd.backcolor = lngGray Then GoTo Exit_Sub
    If btnAdd.backcolor = lngGray Then GoTo Exit_Sub
    
    'MoveSingleItem Me, "lbxSpecies", "lbxTgtSpecies"
    MoveSingleItem Me, "fsub_Species_Listbox", "lbxTgtSpecies"

    'update count
    lblTgtSpeciesCount.Caption = GetListCount(lbxTgtSpecies, True) & " species"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAdd_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnRemove_Click
' Description:  Remove selected items from list
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
'   BLC - 5/10/2015 - changed from MoveSingleItem to RemoveSelectedItems to handle
'                     removing species from list vs. adding back to original list
'                     added update for species count
'   BLC - 5/13/2015 - revised to use global constants vs. tempvars for disabled control
'                     disabled btnRemove, btnRemoveAll when target species count = 0
' ---------------------------------
Private Sub btnRemove_Click()
On Error GoTo Err_Handler
    
    'ignore if 'disabled'
    If btnRemove.backcolor = CTRL_DISABLED Then GoTo Exit_Sub
    
    'MoveSingleItem Me, "lbxTgtSpecies", "fsub_Species_Listbox"
    RemoveSelectedItems lbxTgtSpecies
    
    'update count
    Dim iCount As Integer
    iCount = GetListCount(lbxTgtSpecies, True)
    lblTgtSpeciesCount.Caption = iCount & " species"
    
    'turn off Remove buttons if iCount = 0
    If iCount = 0 Then
        DisableControl btnRemove
        DisableControl btnRemoveAll
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRemove_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnAddAll_Click
' Description:  Add all items to list
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
'   BLC - 5/10/2015 - added update for species count
' ---------------------------------
Private Sub btnAddAll_Click()
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'fetch recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(TempVars.item("strSQL"))
    
    'MoveAllItems Me, "lbxSpecies", "lbxTgtSpecies"
    MoveAllItems Me, "", "lbxTgtSpecies"

    'update count
    lblTgtSpeciesCount.Caption = GetListCount(lbxTgtSpecies, True) & " species"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddAll_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnRemoveAll_Click
' Description:  Remove all items from list
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
'   BLC - 5/10/2015 - changed from MoveAllItems to Form_Load to handle
'                     removing all species vs. moving them to original listbox
' ---------------------------------
Private Sub btnRemoveAll_Click()
On Error GoTo Err_Handler
    
    'MoveAllItems Me, "lbxTgtSpecies", "fsub_Species_Listbox"
    'go back to initial state
    Form_Load

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRemoveAll_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnSaveList_Click
' Description:  Save list items
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/3/2015 - initial version
'   BLC - 5/13/2015 - added LU code and fixed MasterCode bug which substituted LU_Code as Master_Code from tlu_NPCN_Plants
'   BLC - 5/20/2015 - reverted to LUCode for col 2, Code (Master_Plant_Code) in col 0
' ---------------------------------
Private Sub btnSaveList_Click()
On Error GoTo Err_Handler

    Dim iRow As Integer, i As Integer, iTransectOnly, iTgtAreaID
    Dim strMasterCode As String, strSpecies As String, strLUCode As String
    Dim strSQL As String, strInsert As String
    Dim varReturn As Variant
    
    'start @ row 1 (headers = row 0)
    For iRow = 1 To lbxTgtSpecies.ListCount - 1
       
       ' ---------------------------------------------------
       '  NOTE: listbox column MUST have a non-zero width to retrieve its value
       ' ---------------------------------------------------
        strMasterCode = lbxTgtSpecies.Column(0, iRow) 'column 0 = Master_PLANT_Code
        strSpecies = lbxTgtSpecies.Column(1, iRow) 'column 1 = Species name
        strLUCode = lbxTgtSpecies.Column(2, iRow) 'column 2 = LU_Code
        iTransectOnly = Nz(lbxTgtSpecies.Column(3, iRow), 0) 'column 3 = Transect_Only
        iTgtAreaID = Nz(lbxTgtSpecies.Column(4, iRow), 0) 'column 4 = Target_Area_ID
        
       ' ---------------------------------------------------
       '  Check if item exists in tbl_TgtSpecies for Park, Year, Species combo
       ' ---------------------------------------------------
        strSQL = "SELECT * FROM tbl_Target_Species " & _
                 "WHERE Master_PLANT_Code_FK = '" & strMasterCode & _
                 "' AND Park_Code = '" & TempVars.item("park") & _
                 "' AND Target_Year = " & TempVars.item("TgtYear") & ";"
        
        Dim db As DAO.Database
        Dim rs As DAO.Recordset

        Set rs = CurrentDb.OpenRecordset(strSQL) 'CurrentDb.Execute(strSQL, dbFailOnError) >> doesn't compile expected function or variable
      
        'check if there are no records (rs.BOF & rs.EOF are both true)
        If rs.BOF And rs.EOF Then
            
            'set statusbar notice
            varReturn = SysCmd(acSysCmdSetStatus, "Saving " & strSpecies & "...")
            
            'prepare SQL
            strSQL = "INSERT INTO tbl_Target_Species" _
                    & "(Master_Plant_Code_FK, Park_Code, Target_Year, Species_Name, LU_Code, " _
                    & "Transect_Only, Target_Area_ID)" _
                    & "VALUES "
    
            'prepare insert value
            strInsert = "('" & strMasterCode & "','" & TempVars.item("park") & "'," _
                        & TempVars.item("tgtYear") & ",'" & strSpecies & "','" & strLUCode _
                        & "'," & iTransectOnly & "," & iTgtAreaID & ");"
            
            'add comma if more than one row to insert
            'If (lbxTgtSpecies.ListCount - 1) > 1 And iRow < (lbxTgtSpecies.ListCount - 1) Then strInsert = strInsert & ","
            
            'finalize SQL
            strSQL = strSQL & strInsert
            
            'save full target list (insert value) [NOTE: MS Access does not support multiple insert statements, must go 1 @ a time]
            CurrentDb.Execute strSQL, dbFailOnError
            
        End If
        
    Next

        ' check for temp query & clear if it exists
        If QueryExists("tempTgtSpecies") Then
            CurrentDb.QueryDefs.Delete "tempTgtSpecies"
        End If
        
        'open target list
        Dim qdf As QueryDef
        
        Set qdf = CurrentDb.QueryDefs("qry_Tgt_Species_List")
        
        'qdf.Parameters("park") = TempVars.item("park")
        'qdf.Parameters("tgtYear") = CInt(TempVars.item("tgtYear"))
        
        strSQL = qdf.sql
        
        'Call SetValue
        'Set rs = qdf.OpenRecordset
        
        strSQL = "SELECT tbl_Target_Species.Park_Code AS Park, " & _
                 "tbl_Target_Species.Target_Year AS TgtYear, " & _
                 "Master_Plant_Code_FK, Species_Name, LU_Code, " & _
                 "Priority, Transect_Only, Target_Area_ID " & _
                 "FROM tbl_Target_Species " & _
                 "WHERE (((tbl_Target_Species.Target_Year) = CInt(tgtYear)) " & _
                 "And ((LCase([tbl_Target_Species].[Park_Code])) = LCase(park))) " & _
                 "ORDER BY tbl_Target_Species.Species_Name;"
        
        'replace values
        strSQL = Replace(strSQL, "(park)", "('" & TempVars.item("park") & "')")
        strSQL = Replace(strSQL, "(tgtYear)", "(" & TempVars.item("TgtYear") & ")")
        
        'DoCmd.OpenQuery "qryTgtSpeciesList", acViewNormal, acReadOnly
        'DoCmd.RunSQL strSQL <=== NO! not on a SELECT...
        
        CurrentDb.CreateQueryDef("tempTgtSpecies").sql = strSQL
        DoCmd.OpenQuery "tempTgtSpecies"
    
    'set statusbar notice
    varReturn = SysCmd(acSysCmdSetStatus, "Targetlist save complete.")
    
    'pause to view status bar
    For i = 0 To 10000
        i = i
    Next i
    
    'reset status bar
    varReturn = SysCmd(acSysCmdSetStatus, " ")

    'close form
    DoCmd.Close acForm, Me.name

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSaveList_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnSearch_Click
' Description:  Opens species search to find species for populating target list
' Description:  Reset lists to their original state
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 3, 2015 - for NCPN tools
' Revisions:
'   BLC, 3/3/2015  - initial version
'   BLC, 4/30/2015 - integrated into Invasives Reporting tool & updated form naming
' ---------------------------------
Private Sub btnSearch_Click()
On Error GoTo Err_Handler
    Dim originForm As String
    
    originForm = Me.name
    
    'open species search form
    DoCmd.OpenForm "frm_Species_Search", acNormal, , , , acWindowNormal, originForm
    If Forms("frm_Species_Search").Minimized Then DoCmd.Restore

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSearch_Click[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Form_Close
' Description:  Actions for closing form
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015 - initial version
'   BLC - 3/4/2015  - closed species search form
'   BLC, 4/30/2015 - integrated into Invasives Reporting tool & updated form naming
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'clear tempvars
    TempVars.Remove ("park")
    TempVars.Remove ("state")

    'close frmSpeciesSearch if open
    DoCmd.Close acForm, "frm_Species_Search"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[Form_frm_Tgt_Species])"
    End Select
    Resume Exit_Sub
End Sub
