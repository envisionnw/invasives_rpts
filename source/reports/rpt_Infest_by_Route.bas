Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =126
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =9
    ItemSuffix =57
    Left =1440
    Top =-120
    Right =17976
    Bottom =8844
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7060c1afafcbe340
    End
    RecordSource ="tbl_wrk_Infest_Route"
    Caption ="rpt_Infest_by_Route"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xd0020000d0020000d0020000d002000000000000302a00008601000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnActivate ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    RibbonName ="Export"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =9
            FontWeight =700
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =18
            BorderLineStyle =0
            BackStyle =0
            FontName ="Times New Roman"
            AsianLineBreak =255
        End
        Begin ListBox
            TextFontFamily =18
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Times New Roman"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =18
            BorderLineStyle =0
            BackStyle =0
            FontName ="Times New Roman"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="RouteType"
        End
        Begin BreakLevel
            ControlSource ="PlotID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1500
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =3240
                    Top =240
                    Width =4320
                    Height =540
                    FontSize =24
                    FontWeight =400
                    Name ="Label28"
                    Caption ="Infestations by Route"
                End
                Begin Line
                    Top =60
                    Width =10800
                    Name ="Line31"
                End
                Begin Line
                    Top =90
                    Width =10800
                    Name ="Line32"
                End
                Begin Line
                    Top =1380
                    Width =10800
                    Name ="Line33"
                End
                Begin Line
                    Top =1410
                    Width =10800
                    Name ="Line34"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2880
                    Top =900
                    Width =3720
                    Height =360
                    ColumnOrder =0
                    FontSize =16
                    Name ="Park_Name"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6660
                    Top =900
                    Width =1260
                    Height =360
                    ColumnOrder =1
                    FontSize =16
                    TabIndex =1
                    Name ="Visit_Year"

                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1260
            Name ="GroupHeader0"
            Begin
                Begin Rectangle
                    BackStyle =1
                    OldBorderStyle =0
                    Top =60
                    Width =10800
                    Height =420
                    BackColor =14277081
                    Name ="rctRouteType"
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =480
                    BackThemeColorIndex =1
                    BackShade =85.0
                End
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =1440
                    Top =120
                    Width =2310
                    Height =300
                    FontSize =10
                    Name ="RouteType"
                    ControlSource ="RouteType"
                    StatusBarText ="Route type for grouping"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =120
                    LayoutCachedWidth =3750
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            Left =60
                            Top =120
                            Width =1320
                            Height =300
                            FontSize =10
                            Name ="RouteType_Label"
                            Caption ="Type of Route"
                            LayoutCachedLeft =60
                            LayoutCachedTop =120
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =420
                        End
                    End
                End
                Begin Label
                    Left =60
                    Top =720
                    Width =1440
                    Height =270
                    FontSize =10
                    Name ="PlotID_Label"
                    Caption ="Route"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =60
                    LayoutCachedTop =720
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =990
                End
                Begin Label
                    TextAlign =2
                    Left =3420
                    Top =600
                    Width =1200
                    Height =540
                    FontSize =10
                    Name ="RouteLength_Label"
                    Caption ="Route Length (m)"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3420
                    LayoutCachedTop =600
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =1140
                End
                Begin Label
                    TextAlign =2
                    Left =6120
                    Top =900
                    Width =900
                    Height =270
                    FontSize =10
                    Name ="InfestTot_Label"
                    Caption ="Total"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6120
                    LayoutCachedTop =900
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =1170
                End
                Begin Label
                    TextAlign =2
                    Left =7080
                    Top =900
                    Width =1020
                    Height =270
                    FontSize =10
                    Name ="PriorityTot_Label"
                    Caption ="Priority 1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7080
                    LayoutCachedTop =900
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =1170
                End
                Begin Label
                    TextAlign =2
                    Left =8700
                    Top =900
                    Width =900
                    Height =270
                    FontSize =10
                    Name ="TotPct_Label"
                    Caption ="Total"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8700
                    LayoutCachedTop =900
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =1170
                End
                Begin Label
                    TextAlign =2
                    Left =9660
                    Top =900
                    Width =1020
                    Height =270
                    FontSize =10
                    Name ="PriorityPct_Label"
                    Caption ="Priority 1"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9660
                    LayoutCachedTop =900
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =1170
                End
                Begin Line
                    BorderWidth =1
                    Top =480
                    Width =10800
                    Name ="Line35"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =480
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =480
                End
                Begin Line
                    BorderWidth =1
                    Top =540
                    Width =10800
                    Name ="Line36"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =540
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =540
                End
                Begin Line
                    BorderWidth =1
                    Top =1200
                    Width =10800
                    Name ="Line37"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =1200
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1200
                End
                Begin Line
                    BorderWidth =1
                    Top =1200
                    Width =10800
                    Name ="Line38"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =1200
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1200
                End
                Begin Label
                    TextAlign =2
                    Left =6120
                    Top =600
                    Width =1980
                    Height =270
                    FontSize =10
                    Name ="Label44"
                    Caption ="Infestations"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6120
                    LayoutCachedTop =600
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =870
                End
                Begin Line
                    BorderWidth =1
                    Left =6120
                    Top =900
                    Width =1980
                    Name ="Line45"
                    LayoutCachedLeft =6120
                    LayoutCachedTop =900
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =900
                End
                Begin Label
                    TextAlign =2
                    Left =8700
                    Top =600
                    Width =1995
                    Height =270
                    FontSize =10
                    Name ="Label46"
                    Caption ="Infestations/ha"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8700
                    LayoutCachedTop =600
                    LayoutCachedWidth =10695
                    LayoutCachedHeight =870
                End
                Begin Label
                    TextAlign =2
                    Left =4800
                    Top =900
                    Width =1200
                    Height =255
                    Name ="Label47"
                    Caption ="Area (ha)"
                    LayoutCachedLeft =4800
                    LayoutCachedTop =900
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =1155
                End
                Begin Line
                    BorderWidth =1
                    Top =60
                    Width =10800
                    Name ="Line56"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =60
                End
                Begin Line
                    BorderWidth =1
                    Width =10800
                    Name ="Line55"
                    Tag ="DetachedLabel"
                    LayoutCachedWidth =10800
                End
                Begin Line
                    BorderWidth =1
                    Left =8700
                    Top =900
                    Width =1980
                    Name ="Line50"
                    LayoutCachedLeft =8700
                    LayoutCachedTop =900
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =900
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =3240
                    Height =270
                    ColumnWidth =2505
                    Name ="PlotID"
                    ControlSource ="PlotID"
                    StatusBarText ="Route"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3420
                    Top =60
                    Width =1200
                    Height =270
                    ColumnWidth =1230
                    TabIndex =1
                    Name ="RouteLength"
                    ControlSource ="RouteLength"
                    StatusBarText ="Length of route in meters"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =60
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6120
                    Top =60
                    Width =900
                    Height =270
                    ColumnWidth =1905
                    TabIndex =2
                    Name ="InfestTot"
                    ControlSource ="InfestTot"
                    StatusBarText ="Total infestations detected"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =60
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Top =60
                    Width =1020
                    Height =270
                    TabIndex =3
                    Name ="PriorityTot"
                    ControlSource ="PriorityTot"
                    StatusBarText ="Total priority 1 infestations detected"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =60
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8700
                    Top =60
                    Width =900
                    Height =270
                    TabIndex =4
                    Name ="TotPct"
                    ControlSource ="TotPct"
                    StatusBarText ="Infestations per 100 m2"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =60
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9660
                    Top =60
                    Width =990
                    Height =270
                    TabIndex =5
                    Name ="PriorityPct"
                    ControlSource ="PriorityPct"
                    StatusBarText ="Priority 1 infestations per 100 m2"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =60
                    LayoutCachedWidth =10650
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4800
                    Top =60
                    Width =1199
                    Height =270
                    TabIndex =6
                    Name ="RouteArea"
                    ControlSource ="RouteArea"
                    StatusBarText ="Area of route in hectares"

                    LayoutCachedLeft =4800
                    LayoutCachedTop =60
                    LayoutCachedWidth =5999
                    LayoutCachedHeight =330
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =720
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =10560
                    FontWeight =700
                    Name ="Text2"
                    ControlSource ="=\"Summary for \" & \"'RouteType' = \" & \" \" & [RouteType] & \" (\" & Count(*)"
                        " & \" \" & IIf(Count(*)=1,\"detail record\",\"detail records\") & \")\""

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =300
                End
                Begin Label
                    Left =60
                    Top =420
                    Width =774
                    Height =240
                    Name ="Label3"
                    Caption ="Sum"
                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =834
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3420
                    Top =420
                    Width =1200
                    FontWeight =700
                    TabIndex =1
                    Name ="Sum Of RouteLength"
                    ControlSource ="=Sum([RouteLength])"
                    EventProcPrefix ="Sum_Of_RouteLength"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =420
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6120
                    Top =420
                    Width =900
                    FontWeight =700
                    TabIndex =2
                    Name ="Sum Of InfestTot"
                    ControlSource ="=Sum([InfestTot])"
                    EventProcPrefix ="Sum_Of_InfestTot"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =420
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Top =420
                    Width =1020
                    FontWeight =700
                    TabIndex =3
                    Name ="Sum Of PriorityTot"
                    ControlSource ="=Sum([PriorityTot])"
                    EventProcPrefix ="Sum_Of_PriorityTot"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =420
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8700
                    Top =420
                    Width =900
                    FontWeight =700
                    TabIndex =4
                    Name ="Route Type Total Infestations/ha"
                    ControlSource ="=IIf([RouteArea]>0,(Sum([InfestTot]))/Sum([RouteArea]),\"N/A\")"
                    EventProcPrefix ="Route_Type_Total_Infestations_ha"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =420
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9660
                    Top =420
                    Width =1050
                    FontWeight =700
                    TabIndex =5
                    Name ="Route Type Priority Infestations/ha"
                    ControlSource ="=IIf([RouteArea]>0,(Sum([PriorityTot]))/Sum([RouteArea]),\"N/A\")"
                    EventProcPrefix ="Route_Type_Priority_Infestations_ha"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =420
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =660
                End
                Begin Line
                    Left =60
                    Width =9240
                    Name ="Line39"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4800
                    Top =420
                    Width =1200
                    FontWeight =700
                    TabIndex =6
                    Name ="Text48"
                    ControlSource ="=Sum([RouteArea])"

                    LayoutCachedLeft =4800
                    LayoutCachedTop =420
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =660
                End
                Begin Line
                    BorderWidth =1
                    Width =10800
                    Name ="Line51"
                    Tag ="DetachedLabel"
                    LayoutCachedWidth =10800
                End
                Begin Line
                    BorderWidth =1
                    Width =10800
                    Name ="lnRouteTypeFooter"
                    Tag ="DetachedLabel"
                    LayoutCachedWidth =10800
                End
            End
        End
        Begin PageFooter
            Height =510
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =4560
                    Height =270
                    Name ="Text29"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4740
                    Top =240
                    Width =4560
                    Height =270
                    TabIndex =1
                    Name ="Text30"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                End
                Begin Line
                    Width =10800
                    Name ="Line40"
                End
                Begin Line
                    Top =30
                    Width =10800
                    Name ="Line41"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =420
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3420
                    Top =120
                    Width =1200
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="RouteLength Grand Total Sum"
                    ControlSource ="=Sum([RouteLength])"
                    EventProcPrefix ="RouteLength_Grand_Total_Sum"

                    LayoutCachedLeft =3420
                    LayoutCachedTop =120
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =390
                End
                Begin Label
                    Left =60
                    Top =120
                    Width =1152
                    Height =276
                    FontSize =10
                    Name ="lblGrandTotal"
                    Caption ="Grand Total"
                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =1212
                    LayoutCachedHeight =396
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6120
                    Top =120
                    Width =900
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="InfestTot Grand Total Sum"
                    ControlSource ="=Sum([InfestTot])"
                    EventProcPrefix ="InfestTot_Grand_Total_Sum"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =120
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Top =120
                    Width =1020
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="PriorityTot Grand Total Sum"
                    ControlSource ="=Sum([PriorityTot])"
                    EventProcPrefix ="PriorityTot_Grand_Total_Sum"

                    LayoutCachedLeft =7080
                    LayoutCachedTop =120
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8700
                    Top =120
                    Width =900
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="Overall Total Infestations/ha"
                    ControlSource ="=IIf([RouteArea]>0,(Sum([InfestTot]))/Sum([RouteArea]),\"N/A\")"
                    EventProcPrefix ="Overall_Total_Infestations_ha"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =120
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9660
                    Top =120
                    Width =1050
                    Height =270
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="Overall Priority Infestations/ha"
                    ControlSource ="=IIf([RouteArea]>0,(Sum([PriorityTot]))/Sum([RouteArea]),\"N/A\")"
                    EventProcPrefix ="Overall_Priority_Infestations_ha"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =120
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4800
                    Top =120
                    Width =1200
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    Name ="Text49"
                    ControlSource ="=Sum([RouteArea])"

                    LayoutCachedLeft =4800
                    LayoutCachedTop =120
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =360
                End
                Begin Line
                    BorderWidth =1
                    Top =60
                    Width =10800
                    Name ="lnReportFooter"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =60
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

Private Sub Report_Activate()
  Me!Park_Name = DLookup("[ParkName]", "tlu_Parks", "[ParkCode]= '" & Left(OpenArgs, 4) & "'")
  Me!Visit_Year = Right(OpenArgs, 4)
End Sub
