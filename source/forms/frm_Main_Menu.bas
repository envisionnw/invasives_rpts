Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8580
    DatasheetFontHeight =9
    ItemSuffix =10
    Left =135
    Top =300
    Right =8715
    Bottom =6045
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x81b16fedaecae340
    End
    Caption ="Exotic Invasives Reports"
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
        Begin Section
            Height =5760
            BackColor =12902115
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1440
                    Top =360
                    Width =5760
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Invasive Plant Report Menu"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3360
                    Top =1020
                    Width =1740
                    FontSize =10
                    FontWeight =700
                    Name ="ButtonLink"
                    Caption ="Link Data Tables"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =1740
                    Width =2370
                    Height =300
                    TabIndex =1
                    Name ="ButtonInfestation"
                    Caption ="Infestation Report"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4440
                    Top =1740
                    Width =2370
                    Height =300
                    TabIndex =2
                    Name ="ButtonInfestRoute"
                    Caption ="Infestations by Route"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3840
                    Top =4260
                    Width =1035
                    Height =300
                    TabIndex =3
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =2280
                    Width =2370
                    Height =300
                    TabIndex =4
                    Name ="ButtonInfestSize"
                    Caption ="Infestations by Size Class"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4440
                    Top =2280
                    Width =2370
                    Height =300
                    TabIndex =5
                    Name ="ButtonInfestGrowth"
                    Caption ="Infestations by Growth Stage"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =2820
                    Width =2370
                    Height =300
                    TabIndex =6
                    Name ="ButtonMonitoringTransect"
                    Caption ="Monitoring Transect Data"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4440
                    Top =2820
                    Width =2370
                    Height =300
                    TabIndex =7
                    Name ="ButtonSpeciesCoover"
                    Caption ="Species Cover by Route"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =3360
                    Width =2370
                    Height =300
                    TabIndex =8
                    Name ="ButtonTransectCount"
                    Caption ="Transect Count by Route"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4440
                    Top =3300
                    Width =2580
                    FontSize =11
                    TabIndex =9
                    ForeColor =16711680
                    Name ="btnLaunchTgtTool"
                    Caption ="Launch Tgt Species Tool >>"
                    StatusBarText ="Launch target species tool"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =4440
                    LayoutCachedTop =3300
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =3660
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    UseTheme =1
                    Shape =1
                    BackColor =6750156
                    BorderColor =52377
                    ThemeFontIndex =1
                    HoverColor =3407769
                    PressedColor =52224
                    HoverForeColor =2375487
                    PressedForeColor =6750156
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

' ---------------------------------
' SUB:          btnLaunchTgtTool_Click
' Description:  open the species target list tool
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, April 2015 for NCPN tools
' Adapted:      -
' Revisions:    BLC - 4/21/2015 - initial version
' ---------------------------------
Private Sub btnLaunchTgtTool_Click()
On Error GoTo Err_Handler

    'minimize main form
    DoCmd.Minimize

    DoCmd.OpenForm "frm_Tgt_List_Tool", acNormal
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLaunchTgtTool_Click[frm_Main_Menu])"
    End Select
    Resume Exit_Procedure
End Sub

Private Sub ButtonLink_Click()
On Error GoTo Err_ButtonLink_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Connect_Tables"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonLink_Click:
    Exit Sub

Err_ButtonLink_Click:
    MsgBox Err.Description
    Resume Exit_ButtonLink_Click
    
End Sub
Private Sub ButtonInfestation_Click()
On Error GoTo Err_ButtonInfestation_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Select_Infest_Data"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonInfestation_Click:
    Exit Sub

Err_ButtonInfestation_Click:
    MsgBox Err.Description
    Resume Exit_ButtonInfestation_Click
    
End Sub
Private Sub ButtonInfestRoute_Click()
On Error GoTo Err_ButtonInfestRoute_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Select_Infest_by_Route"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonInfestRoute_Click:
    Exit Sub

Err_ButtonInfestRoute_Click:
    MsgBox Err.Description
    Resume Exit_ButtonInfestRoute_Click
    
End Sub
Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub
Private Sub ButtonInfestSize_Click()
On Error GoTo Err_ButtonInfestSize_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Select_Infest_by_Size"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonInfestSize_Click:
    Exit Sub

Err_ButtonInfestSize_Click:
    MsgBox Err.Description
    Resume Exit_ButtonInfestSize_Click
    
End Sub
Private Sub ButtonInfestGrowth_Click()
On Error GoTo Err_ButtonInfestGrowth_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Select_Infest_by_Growth"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonInfestGrowth_Click:
    Exit Sub

Err_ButtonInfestGrowth_Click:
    MsgBox Err.Description
    Resume Exit_ButtonInfestGrowth_Click
    
End Sub
Private Sub ButtonMonitoringTransect_Click()
On Error GoTo Err_ButtonMonitoringTransect_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Monitoring_Transect"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonMonitoringTransect_Click:
    Exit Sub

Err_ButtonMonitoringTransect_Click:
    MsgBox Err.Description
    Resume Exit_ButtonMonitoringTransect_Click
    
End Sub
Private Sub ButtonSpeciesCoover_Click()
On Error GoTo Err_ButtonSpeciesCoover_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Species_Cover_by_Route"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonSpeciesCoover_Click:
    Exit Sub

Err_ButtonSpeciesCoover_Click:
    MsgBox Err.Description
    Resume Exit_ButtonSpeciesCoover_Click
    
End Sub

Private Sub ButtonTransectCount_Click()
On Error GoTo Err_ButtonTransectCount_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Select_Transect_Counts"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonTransectCount_Click:
    Exit Sub

Err_ButtonTransectCount_Click:
    MsgBox Err.Description
    Resume Exit_ButtonTransectCount_Click
End Sub
