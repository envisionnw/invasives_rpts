Version =20
VersionRequired =20
Begin Report
    AllowFilters = NotDefault
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15264
    DatasheetFontHeight =11
    ItemSuffix =141
    Top =900
    Right =14844
    Bottom =6948
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x3d3ba481b6a4e440
    End
    RecordSource ="qry_Tgt_Species_List_Park_Summary"
    Caption ="INVASIVE LIST"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0xf0000000630100001e0100006d01000000000000a03b0000d002000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    RibbonName ="Export"
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
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            ControlSource ="Family"
        End
        Begin BreakLevel
            ControlSource ="utah_species"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =780
            BackColor =15849926
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =4260
                    Height =528
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblReportHdr"
                    Caption ="INVASIVES TARGET LIST"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =588
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10380
                    Width =4680
                    Height =528
                    ColumnOrder =0
                    FontSize =20
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxPark"
                    ControlSource ="=TempVars(\"Park\")+\" SUMMARY\""
                    StatusBarText ="Park for list"
                    GridlineColor =10921638

                    LayoutCachedLeft =10380
                    LayoutCachedWidth =15060
                    LayoutCachedHeight =528
                    ForeTint =50.0
                End
            End
        End
        Begin PageHeader
            Height =1380
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Width =15264
                    Height =480
                    BackColor =15849926
                    BorderColor =10921638
                    Name ="rectPageHdr"
                    GridlineColor =10921638
                    LayoutCachedWidth =15264
                    LayoutCachedHeight =480
                    BackThemeColorIndex =2
                    BackTint =20.0
                End
                Begin Label
                    TextAlign =2
                    Left =1320
                    Top =960
                    Width =1800
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNameUT"
                    Caption ="UT"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =1320
                    LayoutCachedTop =960
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =3360
                    Top =960
                    Width =1980
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNameCO"
                    Caption ="CO"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3360
                    LayoutCachedTop =960
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =5580
                    Top =960
                    Width =1380
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPlantCode"
                    Caption ="Plant Code"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5580
                    LayoutCachedTop =960
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =60
                    Top =960
                    Width =1200
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFamily"
                    Caption ="Family"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =960
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =7080
                    Top =960
                    Width =1680
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCommonName"
                    Caption ="Common Name"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7080
                    LayoutCachedTop =960
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =1320
                    Top =600
                    Width =3720
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNames"
                    Caption ="Species Names"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =1320
                    LayoutCachedTop =600
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =900
                End
                Begin Line
                    Left =1320
                    Top =924
                    Width =3720
                    Name ="lnSpecies"
                    GridlineColor =10921638
                    LayoutCachedLeft =1320
                    LayoutCachedTop =924
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =924
                End
                Begin Line
                    BorderWidth =2
                    Top =1320
                    Width =15264
                    Name ="lnHeader"
                    GridlineColor =10921638
                    LayoutCachedTop =1320
                    LayoutCachedWidth =15264
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10080
                    Top =60
                    Width =5040
                    Height =312
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPage"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =10080
                    LayoutCachedTop =60
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =372
                End
                Begin Label
                    TextAlign =2
                    Left =14340
                    Top =660
                    Width =840
                    Height =540
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPri1Parks"
                    Caption ="# Priority 1 Parks"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =14340
                    LayoutCachedTop =660
                    LayoutCachedWidth =15180
                    LayoutCachedHeight =1200
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6360
                    Width =2880
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxListName"
                    ControlSource ="=IIf([Page]>1,\"Invasives List for \" & TempVars(\"Park\"),\"\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =312
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1440
                    Top =60
                    Width =3300
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDate"
                    ControlSource ="=Format(Now(),\"mmmm d\"\", \"\"yyyy h:nn ampm\")"
                    Format ="Medium Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedTop =60
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =375
                End
                Begin Label
                    Left =120
                    Top =60
                    Width =1320
                    Height =300
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="lblPrinted"
                    Caption ="Printed:"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =360
                    ForeTint =75.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9420
                    Top =600
                    Width =299
                    Height =660
                    TabIndex =3
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear1"
                    ControlSource ="=[MinYear]"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =9420
                    LayoutCachedTop =600
                    LayoutCachedWidth =9719
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10020
                    Top =600
                    Width =300
                    Height =660
                    TabIndex =4
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear2"
                    ControlSource ="=[MinYear]+1"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedTop =600
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Top =600
                    Width =300
                    Height =660
                    TabIndex =5
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear3"
                    ControlSource ="=[MinYear]+2"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =10620
                    LayoutCachedTop =600
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11208
                    Top =600
                    Width =300
                    Height =660
                    TabIndex =6
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear4"
                    ControlSource ="=[MinYear]+3"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =11208
                    LayoutCachedTop =600
                    LayoutCachedWidth =11508
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11820
                    Top =600
                    Width =300
                    Height =660
                    TabIndex =7
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear5"
                    ControlSource ="=[MinYear]+4"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =11820
                    LayoutCachedTop =600
                    LayoutCachedWidth =12120
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12480
                    Top =600
                    Width =300
                    Height =660
                    TabIndex =8
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear6"
                    ControlSource ="=[MinYear]+5"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =12480
                    LayoutCachedTop =600
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13080
                    Top =600
                    Width =300
                    Height =660
                    TabIndex =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear7"
                    ControlSource ="=[MinYear]+6"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =13080
                    LayoutCachedTop =600
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13740
                    Top =600
                    Width =300
                    Height =660
                    TabIndex =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear8"
                    ControlSource ="=[MinYear]+7"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =13740
                    LayoutCachedTop =600
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =2770
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Top =2280
                    Width =15264
                    Height =490
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDetail"
                    GridlineColor =10921638

                    LayoutCachedTop =2280
                    LayoutCachedWidth =15264
                    LayoutCachedHeight =2770
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14160
                    Width =660
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumSpeciesPri1"
                    ControlSource ="=CountInString([ParkYearPriorities],\"-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =14160
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9240
                    Top =24
                    Width =677
                    Height =432
                    FontSize =7
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear1Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([MinYear]))"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220042004c00430041002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =9240
                    LayoutCachedTop =24
                    LayoutCachedWidth =9917
                    LayoutCachedHeight =456
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x42004c00430041002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1320
                    Height =312
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFamily"
                    ControlSource ="Family"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5580
                    Top =60
                    Width =1380
                    Height =312
                    ColumnWidth =2655
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLUCode"
                    ControlSource ="LU_Code"
                    GridlineColor =10921638

                    LayoutCachedLeft =5580
                    LayoutCachedTop =60
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3420
                    Top =60
                    Width =1980
                    Height =312
                    ColumnWidth =1170
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSpeciesCO"
                    ControlSource ="Co_Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =60
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7020
                    Top =60
                    Width =2160
                    Height =312
                    ColumnWidth =2400
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCommon"
                    ControlSource ="Master_Common_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =7020
                    LayoutCachedTop =60
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Top =480
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear1"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & [lblYear1] & \"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =480
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =480
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear2"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & [lblYear2] & \"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =480
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3180
                    Top =480
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear3"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & [lblYear3] & \"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =3180
                    LayoutCachedTop =480
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3480
                    Top =480
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear4"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & [lblYear4] & \"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =3480
                    LayoutCachedTop =480
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3780
                    Top =480
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear5"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & [lblYear5] & \"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =3780
                    LayoutCachedTop =480
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4080
                    Top =480
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear6"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & [lblYear6] & \"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =4080
                    LayoutCachedTop =480
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4380
                    Top =480
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear7"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & [lblYear7] & \"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =480
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4680
                    Top =480
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear8"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & [lblYear8] & \"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =4680
                    LayoutCachedTop =480
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1380
                    Top =840
                    Width =1980
                    Height =312
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSpeciesUT"
                    ControlSource ="utah_species"
                    GridlineColor =10921638

                    LayoutCachedLeft =1380
                    LayoutCachedTop =840
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =1152
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1920
                    Top =480
                    Width =660
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxRunSumPri1"
                    ControlSource ="=CountInString([ParkYearPriorities],\"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =480
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =780
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4860
                    Top =840
                    Width =5280
                    Height =300
                    FontSize =9
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxAll"
                    ControlSource ="ParkYearPriorities"
                    GridlineColor =10921638

                    LayoutCachedLeft =4860
                    LayoutCachedTop =840
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =1140
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =3960
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10020
                    Top =960
                    Width =1140
                    Height =312
                    FontSize =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumPriority1"
                    ControlSource ="=[tbxRunSumPri1]"
                    StatusBarText ="Standard park code (CANY, FOBU, etc.)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedTop =960
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =1272
                End
                Begin Label
                    TextAlign =3
                    Left =7200
                    Top =960
                    Width =2700
                    Height =324
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTotalNum"
                    Caption ="Total # Priority 1 Species ="
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7200
                    LayoutCachedTop =960
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =1284
                End
                Begin Line
                    BorderWidth =2
                    Width =15264
                    Name ="lnPageFooter"
                    GridlineColor =10921638
                    LayoutCachedWidth =15264
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9480
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumBLCA"
                    ControlSource ="=[tbxYear1]"
                    StatusBarText ="=\"Total # priority 1 (\"&[lblYear1]&\")\""
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =60
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =3
                    Left =5760
                    Top =60
                    Width =3480
                    Height =324
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblParkYearPriorities"
                    Caption ="Total # Priority 1 Species by Park =>"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5760
                    LayoutCachedTop =60
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =384
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9480
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueYear1"
                    ControlSource ="=Sum(IIf(CountInString([ParkYearPriorities],\"1\")=1,CountInString([ParkYearPrio"
                        "rities],[TempVars]![Park] & \"-\" & [lblYear1] & \"-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (Year 1)"
                    ConditionalFormat = Begin
                        0x010000002e010000010000000100000000000000000000006600000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4900490066002800530075006d002800490049006600280043006f0075006e00 ,
                        0x740049006e0053007400720069006e00670028005b005000610072006b005000 ,
                        0x720069006f007200690074006900650073005d002c0022003100220029003d00 ,
                        0x31002c0043006f0075006e00740049006e0053007400720069006e0067002800 ,
                        0x5b005000610072006b005000720069006f007200690074006900650073005d00 ,
                        0x2c00220042004c00430041002d003100220029002c003000290029003e003000 ,
                        0x2c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =420
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =720
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff00650000004900 ,
                        0x490066002800530075006d002800490049006600280043006f0075006e007400 ,
                        0x49006e0053007400720069006e00670028005b005000610072006b0050007200 ,
                        0x69006f007200690074006900650073005d002c0022003100220029003d003100 ,
                        0x2c0043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220042004c00430041002d003100220029002c003000290029003e0030002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    BackStyle =1
                    TextAlign =3
                    Left =6852
                    Top =420
                    Width =2388
                    Height =288
                    FontSize =10
                    BackColor =16777164
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblUniquePri1"
                    Caption ="Unique Priority 1 Species =>"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6852
                    LayoutCachedTop =420
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =708
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9480
                    Top =1440
                    Width =300
                    Height =2160
                    FontSize =8
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear1"
                    ControlSource ="=getListLastModifiedDate([lblYear1],TempVars(\"Park\"))"
                    StatusBarText ="=\"List Last Modification Date (\"& [lblYear1] &\")\""
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =3600
                End
                Begin Label
                    TextAlign =3
                    Left =7860
                    Top =1440
                    Width =1260
                    Height =960
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLastModDate"
                    Caption ="Last      Modified  =>\015\012Date      "
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7860
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =2400
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2040
                    Top =1320
                    Width =1380
                    Height =585
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text135"
                    ControlSource ="=TempVars(\"Park\") & \"-\" & [lblYear1] & \"-1\""
                    GridlineColor =10921638

                    LayoutCachedLeft =2040
                    LayoutCachedTop =1320
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =1905
                    Begin
                        Begin Label
                            Left =240
                            Top =1320
                            Width =810
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label136"
                            Caption ="Text135"
                            GridlineColor =10921638
                            LayoutCachedLeft =240
                            LayoutCachedTop =1320
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =1635
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2100
                    Top =2100
                    Width =1380
                    Height =585
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text137"
                    ControlSource ="=[MinYear]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =2100
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =2685
                    Begin
                        Begin Label
                            Left =300
                            Top =2100
                            Width =810
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label138"
                            Caption ="Text135"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =2100
                            LayoutCachedWidth =1110
                            LayoutCachedHeight =2415
                        End
                    End
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
' MODULE:       Report_rpt_Tgt_Species_List_Park_Summary
' Description:  Load species list to target species list functions and routines
'
' Source/date:  Bonnie Campbell, 9/21/2015
' Revisions:    BLC - 9/21/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Report_Open
' Description:  Actions for when report opens
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Notes:
'   Consider references for performance improvements/user cues that report is still being generated
'   http://stackoverflow.com/questions/11477297/giving-an-alias-to-a-subquery-containing-a-join-in-access
' Source/date:
' Adapted:      Bonnie Campbell, September 21, 2015 - for NCPN tools
' Revisions:
'   BLC - 9/21/2015 - initial version
'   BLC - 9/30/2015 - set report data source SQL to update to TempVars("Park")
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)

On Error GoTo Err_Handler

    'get report data source & alter it using target year to reduce query time?
    Dim i As Integer
    
    Screen.MousePointer = 11 'Hour Glass

    DoCmd.OpenForm "frm_Progress_Bar", acNormal
    
    For i = 1 To 10
        
        Forms("frm_Progress_Bar").Increment i * 10, "Preparing report..."
    
    Next

    'update data source
    Dim strSQL As String
    Dim qdf As DAO.QueryDef
    
    Set qdf = CurrentDb.QueryDefs("qry_Tgt_Species_List_Park_Summary")
    strSQL = qdf.sql
Debug.Print strSQL
    strSQL = Replace(strSQL, "PARKNAME", TempVars!Park)
Debug.Print strSQL
    qdf.sql = strSQL
    
'    strSQL = "SELECT DISTINCT Master_Plant_Code_FK, LU_Code, Family, Species_Name, utah_species, Co_Species, Wy_Species, & " _
'                & "Master_Common_Name, ConcatRelated("ParkYearPriority","qry_Annual_Complete_Tgt_Species_Lists"," _
'                & ""Park = '" & TempVars!Park & "' AND Species_Name='"+Species_Name+"'",'',"|") AS ParkYearPriorities, " _
'                & "(SELECT Min(TgtYear) FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park = '" & TempVars!Park & "') AS MinYear, " _
'                & "(SELECT Max(TgtYear) FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park = '" & TempVars!Park & "') AS MaxYear " _
'                & "FROM (SELECT * FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park = '" & TempVars!Park & "')  AS [%$##@_Alias] " _
'                & "GROUP BY Park, Master_Plant_Code_FK, LU_Code, Family, Species_Name, Priority, Transect_Only, " _
'                & "Target_Area_ID, Tgt_Area, utah_species, Co_Species, Wy_Species, Master_Common_Name, PriorityTarget, " _
'                & "ParkYearPriority, SpeciesYear;"

    If Len(Me.OpenArgs) > 0 Then
        ' Bob Larsen, January 28, 2012
        ' https://social.msdn.microsoft.com/Forums/office/en-US/3e126484-112f-4854-a5c0-2e9ef48e02bc/how-to-change-recordsource-for-a-report-with-vba?forum=accessdev
        'set recordset to passed in SQL via OpenArgs
        'If Me.OpenArgs <> vbNullString Then
        'Me.Recordset = Me.OpenArgs
        ' dyDMA, Sept 8, 2008
        ' http://www.utteraccess.com/forum/Run-time-error-32585-t1710296.html
        '==> Run-time Error 32585: This feature is only available in an ADP
        '==> Only Access ADP's can use this method (assign report recordset @ run-time)
        '==> Not available for *.mdb or *.accdb's
        
        'set orderby
        Me.OrderBy = Me.OpenArgs
    End If
    'sPercentage

If ReportIsLoaded("rpt_Tgt_Species_List_Park_Summary") Then
     DoEvents
     Pause (5) 'was 15
     DoCmd.Close acForm, "frm_Progress_Bar"
     DoEvents
    
    Pause (10) 'was 30
    ' clear statusbar note running report
    SysCmd acSysCmdSetStatus, "Calculations complete! Fetching report..."
End If

Screen.MousePointer = 1 'Standard Cursor
    
    'reset the query
    Set qdf = CurrentDb.QueryDefs("qry_Tgt_Species_List_Park_Summary")
    strSQL = qdf.sql
Debug.Print strSQL
    strSQL = Replace(strSQL, TempVars!Park, "PARKNAME")
Debug.Print strSQL
    qdf.sql = strSQL

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[Report_rpt_Tgt_Species_List_Park_Summary])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Report_Load
' Description:  Actions for when report is loaded
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, September 21, 2015 - for NCPN tools
' Revisions:
'   BLC - 9/21/2015 - initial version
' ---------------------------------
Private Sub Report_Load()
On Error GoTo Err_Handler
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Load[Report_rpt_Tgt_Species_List_Park_Summary])"
    End Select
    Resume Exit_Sub
End Sub
