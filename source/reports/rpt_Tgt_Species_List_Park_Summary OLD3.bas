Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15264
    DatasheetFontHeight =11
    Left =45
    Top =1155
    Right =13650
    Bottom =11010
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xcd59eeceb7a4e440
    End
    RecordSource ="qry_Tgt_Species_List_Park_Summary"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
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
                    Left =10320
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

                    LayoutCachedLeft =10320
                    LayoutCachedWidth =15000
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
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13740
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =7
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
                    ColumnOrder =6
                    TabIndex =1
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
                    Left =12480
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =5
                    TabIndex =2
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
                    Left =11820
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =4
                    TabIndex =3
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
                    Left =11208
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =3
                    TabIndex =4
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
                    Left =10620
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =2
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
                    Left =10020
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =1
                    TabIndex =6
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
                    Left =9420
                    Top =600
                    Width =299
                    Height =660
                    ColumnOrder =0
                    TabIndex =7
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
                Begin Line
                    Left =1260
                    Top =924
                    Width =3720
                    Name ="lnSpecies"
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedTop =924
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =924
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
                    ColumnOrder =10
                    TabIndex =8
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
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6360
                    Width =2880
                    Height =312
                    ColumnOrder =9
                    TabIndex =9
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
                    TabStop = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10080
                    Top =60
                    Width =5040
                    Height =312
                    ColumnOrder =8
                    TabIndex =10
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
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =960
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
                    Width =15264
                    Height =490
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDetail"
                    GridlineColor =10921638

                    LayoutCachedWidth =15264
                    LayoutCachedHeight =490
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Top =60
                    Width =1320
                    Height =312
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFamily"
                    ControlSource ="Family"
                    GridlineColor =10921638

                    LayoutCachedTop =60
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5520
                    Top =60
                    Width =1380
                    Height =312
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLUCode"
                    ControlSource ="LU_Code"
                    GridlineColor =10921638

                    LayoutCachedLeft =5520
                    LayoutCachedTop =60
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3360
                    Top =60
                    Width =1980
                    Height =312
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSpeciesCO"
                    ControlSource ="Co_Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =3360
                    LayoutCachedTop =60
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6960
                    Top =60
                    Width =2160
                    Height =312
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCommon"
                    ControlSource ="Master_Common_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =6960
                    LayoutCachedTop =60
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9240
                    Top =60
                    Width =677
                    Height =432
                    FontSize =7
                    TabIndex =5
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
                    LayoutCachedTop =60
                    LayoutCachedWidth =9917
                    LayoutCachedHeight =492
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
                    Left =1380
                    Top =60
                    Width =1980
                    Height =312
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSpeciesUT"
                    ControlSource ="utah_species"
                    GridlineColor =10921638

                    LayoutCachedLeft =1380
                    LayoutCachedTop =60
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4860
                    Top =540
                    Width =5280
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxAll"
                    ControlSource ="ParkYearPriorities"
                    GridlineColor =10921638

                    LayoutCachedLeft =4860
                    LayoutCachedTop =540
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2340
                    Top =540
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear1"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear1] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2340
                    LayoutCachedTop =540
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2640
                    Top =540
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear2"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear2] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =540
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2940
                    Top =540
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear3"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear3] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2940
                    LayoutCachedTop =540
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3240
                    Top =540
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear4"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear4] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =540
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3540
                    Top =540
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear5"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear5] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =3540
                    LayoutCachedTop =540
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3840
                    Top =540
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear6"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear6] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =3840
                    LayoutCachedTop =540
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4140
                    Top =540
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear7"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear7] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =4140
                    LayoutCachedTop =540
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4440
                    Top =540
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear8"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear8] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =4440
                    LayoutCachedTop =540
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Top =540
                    Width =660
                    Height =300
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxRunSumPri1"
                    ControlSource ="=CountInString([ParkYearPriorities],\"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedTop =540
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14148
                    Top =36
                    Width =660
                    Height =300
                    FontSize =9
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumSpeciesPri1"
                    ControlSource ="=CountInString([ParkYearPriorities],\"-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =14148
                    LayoutCachedTop =36
                    LayoutCachedWidth =14808
                    LayoutCachedHeight =336
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13548
                    Top =60
                    Width =677
                    Height =432
                    FontSize =7
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear8Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([lblYear8]))"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x22005a0049004f004e002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =13548
                    LayoutCachedTop =60
                    LayoutCachedWidth =14225
                    LayoutCachedHeight =492
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x5a0049004f004e002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12888
                    Top =60
                    Width =677
                    Height =432
                    FontSize =7
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear7Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([lblYear7]))"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220047004f00530050002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =12888
                    LayoutCachedTop =60
                    LayoutCachedWidth =13565
                    LayoutCachedHeight =492
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x47004f00530050002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12288
                    Top =60
                    Width =677
                    Height =432
                    FontSize =7
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear6Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([lblYear6]))"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220046004f00420055002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =12288
                    LayoutCachedTop =60
                    LayoutCachedWidth =12965
                    LayoutCachedHeight =492
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x46004f00420055002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11628
                    Top =60
                    Width =677
                    Height =432
                    FontSize =7
                    TabIndex =21
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear5Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([lblYear5]))"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x2200440049004e004f002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =11628
                    LayoutCachedTop =60
                    LayoutCachedWidth =12305
                    LayoutCachedHeight =492
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x440049004e004f002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11028
                    Top =60
                    Width =677
                    Height =432
                    FontSize =7
                    TabIndex =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear4Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([lblYear4]))"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220043005500520045002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =11028
                    LayoutCachedTop =60
                    LayoutCachedWidth =11705
                    LayoutCachedHeight =492
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x43005500520045002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10428
                    Top =60
                    Width =677
                    Height =432
                    FontSize =7
                    TabIndex =23
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear3Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([lblYear3]))"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220043004f004c004d002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =10428
                    LayoutCachedTop =60
                    LayoutCachedWidth =11105
                    LayoutCachedHeight =492
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x43004f004c004d002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9828
                    Top =60
                    Width =677
                    Height =432
                    FontSize =7
                    TabIndex =24
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear2Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([lblYear2]))"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220043004100520045002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =9828
                    LayoutCachedTop =60
                    LayoutCachedWidth =10505
                    LayoutCachedHeight =492
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x43004100520045002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
            End
        End
        Begin PageFooter
            Height =360
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =3720
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9420
                    Width =300
                    Height =300
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumBLCA"
                    ControlSource ="=[tbxYear1]"
                    StatusBarText ="=\"Total # priority 1 (\"&[lblYear1]&\")\""
                    GridlineColor =10921638

                    LayoutCachedLeft =9420
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCARE"
                    ControlSource ="=[tbxYear2]"
                    StatusBarText ="Total # priority 1 (CARE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =9960
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10560
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCOLM"
                    ControlSource ="=[tbxYear3]"
                    StatusBarText ="Total # priority 1 (COLM)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10560
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11160
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCURE"
                    ControlSource ="=[tbxYear4]"
                    StatusBarText ="Total # priority 1 (CURE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11160
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11760
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumDINO"
                    ControlSource ="=[tbxYear5]"
                    StatusBarText ="Total # priority 1 (DINO)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11760
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12360
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumFOBU"
                    ControlSource ="=[tbxYear6]"
                    StatusBarText ="Total # priority 1 (FOBU)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12360
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13020
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumGOSP"
                    ControlSource ="=[tbxYear7]"
                    StatusBarText ="Total # priority 1 (GOSP)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13020
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13680
                    Width =300
                    Height =270
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumZION"
                    ControlSource ="=[tbxYear8]"
                    StatusBarText ="Total # priority 1 (ZION)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13680
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =270
                End
                Begin Label
                    TextAlign =3
                    Left =5700
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
                    LayoutCachedLeft =5700
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =324
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9420
                    Top =360
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueYear1"
                    ControlSource ="=Sum(IIf(CountInString([ParkYearPriorities],\"1\")=1,CountInString([ParkYearPrio"
                        "rities],[TempVars]![Park] & \"-\" & [MinYear] & \"-1\"),0))"
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

                    LayoutCachedLeft =9420
                    LayoutCachedTop =360
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =660
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
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Top =360
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueYear2"
                    ControlSource ="=Sum(IIf(CountInString([ParkYearPriorities],\"1\")=1,CountInString([ParkYearPrio"
                        "rities],[TempVars].[Park] & \"-\" & [MinYear]+1 & \"-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (Year 2)"
                    ConditionalFormat = Begin
                        0x010000002e010000010000000100000000000000000000006600000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4900490066002800530075006d002800490049006600280043006f0075006e00 ,
                        0x740049006e0053007400720069006e00670028005b005000610072006b005000 ,
                        0x720069006f007200690074006900650073005d002c0022003100220029003d00 ,
                        0x31002c0043006f0075006e00740049006e0053007400720069006e0067002800 ,
                        0x5b005000610072006b005000720069006f007200690074006900650073005d00 ,
                        0x2c00220043004100520045002d003100220029002c003000290029003e003000 ,
                        0x2c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =9960
                    LayoutCachedTop =360
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =660
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff00650000004900 ,
                        0x490066002800530075006d002800490049006600280043006f0075006e007400 ,
                        0x49006e0053007400720069006e00670028005b005000610072006b0050007200 ,
                        0x69006f007200690074006900650073005d002c0022003100220029003d003100 ,
                        0x2c0043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220043004100520045002d003100220029002c003000290029003e0030002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10560
                    Top =360
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueYear3"
                    ControlSource ="=Sum(IIf(CountInString([ParkYearPriorities],\"1\")=1,CountInString([ParkYearPrio"
                        "rities],[TempVars]![Park] & \"-\" & [MinYear]+2 & \"-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (Year 3)"
                    ConditionalFormat = Begin
                        0x010000002e010000010000000100000000000000000000006600000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4900490066002800530075006d002800490049006600280043006f0075006e00 ,
                        0x740049006e0053007400720069006e00670028005b005000610072006b005000 ,
                        0x720069006f007200690074006900650073005d002c0022003100220029003d00 ,
                        0x31002c0043006f0075006e00740049006e0053007400720069006e0067002800 ,
                        0x5b005000610072006b005000720069006f007200690074006900650073005d00 ,
                        0x2c00220043004f004c004d002d003100220029002c003000290029003e003000 ,
                        0x2c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =10560
                    LayoutCachedTop =360
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =660
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff00650000004900 ,
                        0x490066002800530075006d002800490049006600280043006f0075006e007400 ,
                        0x49006e0053007400720069006e00670028005b005000610072006b0050007200 ,
                        0x69006f007200690074006900650073005d002c0022003100220029003d003100 ,
                        0x2c0043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220043004f004c004d002d003100220029002c003000290029003e0030002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11160
                    Top =360
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueYear4"
                    ControlSource ="=Sum(IIf(CountInString([ParkYearPriorities],\"1\")=1,CountInString([ParkYearPrio"
                        "rities],[TempVars].[Park] & \"-\" & [MinYear]+3 & \"-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (Year 4)"
                    ConditionalFormat = Begin
                        0x010000002e010000010000000100000000000000000000006600000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4900490066002800530075006d002800490049006600280043006f0075006e00 ,
                        0x740049006e0053007400720069006e00670028005b005000610072006b005000 ,
                        0x720069006f007200690074006900650073005d002c0022003100220029003d00 ,
                        0x31002c0043006f0075006e00740049006e0053007400720069006e0067002800 ,
                        0x5b005000610072006b005000720069006f007200690074006900650073005d00 ,
                        0x2c00220043005500520045002d003100220029002c003000290029003e003000 ,
                        0x2c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =11160
                    LayoutCachedTop =360
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =660
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff00650000004900 ,
                        0x490066002800530075006d002800490049006600280043006f0075006e007400 ,
                        0x49006e0053007400720069006e00670028005b005000610072006b0050007200 ,
                        0x69006f007200690074006900650073005d002c0022003100220029003d003100 ,
                        0x2c0043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220043005500520045002d003100220029002c003000290029003e0030002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11760
                    Top =360
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueYear5"
                    ControlSource ="=Sum(IIf(CountInString([ParkYearPriorities],\"1\")=1,CountInString([ParkYearPrio"
                        "rities],[TempVars].[Park] & \"-\" & [MinYear]+4 & \"-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (Year 5)"
                    ConditionalFormat = Begin
                        0x010000002e010000010000000100000000000000000000006600000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4900490066002800530075006d002800490049006600280043006f0075006e00 ,
                        0x740049006e0053007400720069006e00670028005b005000610072006b005000 ,
                        0x720069006f007200690074006900650073005d002c0022003100220029003d00 ,
                        0x31002c0043006f0075006e00740049006e0053007400720069006e0067002800 ,
                        0x5b005000610072006b005000720069006f007200690074006900650073005d00 ,
                        0x2c002200440049004e004f002d003100220029002c003000290029003e003000 ,
                        0x2c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =11760
                    LayoutCachedTop =360
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =660
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff00650000004900 ,
                        0x490066002800530075006d002800490049006600280043006f0075006e007400 ,
                        0x49006e0053007400720069006e00670028005b005000610072006b0050007200 ,
                        0x69006f007200690074006900650073005d002c0022003100220029003d003100 ,
                        0x2c0043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x2200440049004e004f002d003100220029002c003000290029003e0030002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12360
                    Top =360
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueYear6"
                    ControlSource ="=Sum(IIf(CountInString([ParkYearPriorities],\"1\")=1,CountInString([ParkYearPrio"
                        "rities],[TempVars].[Park] & \"-\" & [MinYear]+5 & \"-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (Year 6)"
                    ConditionalFormat = Begin
                        0x010000002e010000010000000100000000000000000000006600000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4900490066002800530075006d002800490049006600280043006f0075006e00 ,
                        0x740049006e0053007400720069006e00670028005b005000610072006b005000 ,
                        0x720069006f007200690074006900650073005d002c0022003100220029003d00 ,
                        0x31002c0043006f0075006e00740049006e0053007400720069006e0067002800 ,
                        0x5b005000610072006b005000720069006f007200690074006900650073005d00 ,
                        0x2c00220046004f00420055002d003100220029002c003000290029003e003000 ,
                        0x2c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =12360
                    LayoutCachedTop =360
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =660
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff00650000004900 ,
                        0x490066002800530075006d002800490049006600280043006f0075006e007400 ,
                        0x49006e0053007400720069006e00670028005b005000610072006b0050007200 ,
                        0x69006f007200690074006900650073005d002c0022003100220029003d003100 ,
                        0x2c0043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220046004f00420055002d003100220029002c003000290029003e0030002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13020
                    Top =360
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueYear7"
                    ControlSource ="=Sum(IIf(CountInString([ParkYearPriorities],\"1\")=1,CountInString([ParkYearPrio"
                        "rities],[TempVars].[Park] & \"-\" & [MinYear]+6 & \"-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (Year 7)"
                    ConditionalFormat = Begin
                        0x010000002e010000010000000100000000000000000000006600000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4900490066002800530075006d002800490049006600280043006f0075006e00 ,
                        0x740049006e0053007400720069006e00670028005b005000610072006b005000 ,
                        0x720069006f007200690074006900650073005d002c0022003100220029003d00 ,
                        0x31002c0043006f0075006e00740049006e0053007400720069006e0067002800 ,
                        0x5b005000610072006b005000720069006f007200690074006900650073005d00 ,
                        0x2c00220047004f00530050002d003100220029002c003000290029003e003000 ,
                        0x2c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =13020
                    LayoutCachedTop =360
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =660
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff00650000004900 ,
                        0x490066002800530075006d002800490049006600280043006f0075006e007400 ,
                        0x49006e0053007400720069006e00670028005b005000610072006b0050007200 ,
                        0x69006f007200690074006900650073005d002c0022003100220029003d003100 ,
                        0x2c0043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220047004f00530050002d003100220029002c003000290029003e0030002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13680
                    Top =360
                    Width =300
                    Height =270
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueYear8"
                    ControlSource ="=Sum(IIf(CountInString([ParkYearPriorities],\"1\")=1,CountInString([ParkYearPrio"
                        "rities],[TempVars].[Park] & \"-\" & [MinYear]+7 & \"-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (Year 8)"
                    ConditionalFormat = Begin
                        0x010000002e010000010000000100000000000000000000006600000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4900490066002800530075006d002800490049006600280043006f0075006e00 ,
                        0x740049006e0053007400720069006e00670028005b005000610072006b005000 ,
                        0x720069006f007200690074006900650073005d002c0022003100220029003d00 ,
                        0x31002c0043006f0075006e00740049006e0053007400720069006e0067002800 ,
                        0x5b005000610072006b005000720069006f007200690074006900650073005d00 ,
                        0x2c0022005a0049004f004e002d003100220029002c003000290029003e003000 ,
                        0x2c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =13680
                    LayoutCachedTop =360
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =630
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff00650000004900 ,
                        0x490066002800530075006d002800490049006600280043006f0075006e007400 ,
                        0x49006e0053007400720069006e00670028005b005000610072006b0050007200 ,
                        0x69006f007200690074006900650073005d002c0022003100220029003d003100 ,
                        0x2c0043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x22005a0049004f004e002d003100220029002c003000290029003e0030002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    BackStyle =1
                    TextAlign =3
                    Left =6792
                    Top =360
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
                    LayoutCachedLeft =6792
                    LayoutCachedTop =360
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =648
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    TabStop = NotDefault
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10020
                    Top =900
                    Width =1140
                    Height =312
                    FontSize =12
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumPriority1"
                    ControlSource ="=[tbxRunSumPri1]"
                    StatusBarText ="Standard park code (CANY, FOBU, etc.)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedTop =900
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =1212
                End
                Begin Label
                    TextAlign =3
                    Left =7200
                    Top =900
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
                    LayoutCachedTop =900
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =1224
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9420
                    Top =1440
                    Width =300
                    Height =2160
                    FontSize =8
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear1"
                    ControlSource ="=getListLastModifiedDate([lblYear1],TempVars(\"Park\"))"
                    StatusBarText ="=\"List Last Modification Date (\"& [lblYear1] &\")\""
                    GridlineColor =10921638

                    LayoutCachedLeft =9420
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =3600
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Top =1440
                    Width =300
                    Height =2160
                    FontSize =8
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear2"
                    ControlSource ="=getListLastModifiedDate([lblYear2],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (CARE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =9960
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =3600
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10560
                    Top =1440
                    Width =300
                    Height =2160
                    FontSize =8
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear3"
                    ControlSource ="=getListLastModifiedDate([lblYear3],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (COLM)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10560
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =3600
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11160
                    Top =1440
                    Width =300
                    Height =2160
                    FontSize =8
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear4"
                    ControlSource ="=getListLastModifiedDate([lblYear4],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (CURE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11160
                    LayoutCachedTop =1440
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =3600
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11760
                    Top =1440
                    Width =300
                    Height =2160
                    FontSize =8
                    TabIndex =21
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear5"
                    ControlSource ="=getListLastModifiedDate([lblYear5],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (DINO)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11760
                    LayoutCachedTop =1440
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =3600
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12360
                    Top =1440
                    Width =300
                    Height =2160
                    FontSize =8
                    TabIndex =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear6"
                    ControlSource ="=getListLastModifiedDate([lblYear6],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (FOBU)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12360
                    LayoutCachedTop =1440
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =3600
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13020
                    Top =1440
                    Width =300
                    Height =2160
                    FontSize =8
                    TabIndex =23
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear7"
                    ControlSource ="=getListLastModifiedDate([lblYear6],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (GOSP)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13020
                    LayoutCachedTop =1440
                    LayoutCachedWidth =13320
                    LayoutCachedHeight =3600
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13680
                    Top =1440
                    Width =300
                    Height =2160
                    FontSize =8
                    TabIndex =24
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear8"
                    ControlSource ="=getListLastModifiedDate([lblYear8],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (ZION)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13680
                    LayoutCachedTop =1440
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =3600
                End
                Begin Label
                    TextAlign =3
                    Left =7800
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
                    LayoutCachedLeft =7800
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =2400
                End
            End
        End
    End
End
