Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15480
    DatasheetFontHeight =11
    ItemSuffix =125
    Left =84
    Right =15924
    Bottom =7536
    DatasheetGridlinesColor =14806254
    Filter ="TgtYear=2013"
    RecSrcDt = Begin
        0xbdcb1e56bd91e440
    End
    RecordSource ="qry_Tgt_Species_List_Annual_Summary"
    Caption ="INVASIVE LIST"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0xa8000000630100001e0100006d01000000000000df3d0000ea01000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
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
            ControlSource ="Species_Name"
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
                    Width =3912
                    Height =528
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblReportHdr"
                    Caption ="INVASIVES TARGET LIST"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3972
                    LayoutCachedHeight =588
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Width =4680
                    Height =528
                    ColumnOrder =0
                    FontSize =20
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxYear"
                    ControlSource ="=TempVars(\"TgtYear\")+\" ANNUAL SUMMARY\""
                    StatusBarText ="Park and year for list"
                    GridlineColor =10921638

                    LayoutCachedLeft =10620
                    LayoutCachedWidth =15300
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
                    Width =15480
                    Height =480
                    BackColor =15849926
                    BorderColor =10921638
                    Name ="rectPageHdr"
                    GridlineColor =10921638
                    LayoutCachedWidth =15480
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
                    Left =7260
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
                    LayoutCachedLeft =7260
                    LayoutCachedTop =960
                    LayoutCachedWidth =8940
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
                    Width =15480
                    Name ="lnHeader"
                    GridlineColor =10921638
                    LayoutCachedTop =1320
                    LayoutCachedWidth =15480
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =5040
                    Height =312
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDate"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10320
                    Top =60
                    Width =5040
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPage"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =10320
                    LayoutCachedTop =60
                    LayoutCachedWidth =15360
                    LayoutCachedHeight =372
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =9420
                    Top =600
                    Width =299
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblBLCA"
                    Caption ="BLCA"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =9420
                    LayoutCachedTop =600
                    LayoutCachedWidth =9719
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =9960
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCARE"
                    Caption ="CARE"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =9960
                    LayoutCachedTop =600
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =10500
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCOLM"
                    Caption ="COLM"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10500
                    LayoutCachedTop =600
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =11028
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCURE"
                    Caption ="CURE"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =11028
                    LayoutCachedTop =600
                    LayoutCachedWidth =11328
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =11580
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblDINO"
                    Caption ="DINO"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =11580
                    LayoutCachedTop =600
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =12060
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFOBU"
                    Caption ="FOBU"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =12060
                    LayoutCachedTop =600
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =12600
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblGOSP"
                    Caption ="GOSP"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =12600
                    LayoutCachedTop =600
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =13140
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblZION"
                    Caption ="ZION"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =13140
                    LayoutCachedTop =600
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =13620
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
                    LayoutCachedLeft =13620
                    LayoutCachedTop =660
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =1200
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6360
                    Width =2880
                    Height =312
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxListName"
                    ControlSource ="=IIf([Page]>1,\"Invasives List for \" & TempVars(\"TgtYear\"),\"\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =312
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =490
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Width =15480
                    Height =490
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDetail"
                    GridlineColor =10921638

                    LayoutCachedWidth =15480
                    LayoutCachedHeight =490
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13440
                    Width =660
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumSpeciesPri1"
                    ControlSource ="=CountInString([ParkPriorities],1)"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =13440
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12960
                    Top =24
                    Width =655
                    Height =432
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxZIONPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"ZION\",[tbxAll])"
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

                    LayoutCachedLeft =12960
                    LayoutCachedTop =24
                    LayoutCachedWidth =13615
                    LayoutCachedHeight =456
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12420
                    Top =24
                    Width =655
                    Height =432
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxGOSPPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"GOSP\",[tbxAll])"
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

                    LayoutCachedLeft =12420
                    LayoutCachedTop =24
                    LayoutCachedWidth =13075
                    LayoutCachedHeight =456
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12000
                    Top =24
                    Width =655
                    Height =432
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFOBUPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"FOBU\",[tbxAll])"
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

                    LayoutCachedLeft =12000
                    LayoutCachedTop =24
                    LayoutCachedWidth =12655
                    LayoutCachedHeight =456
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11400
                    Top =24
                    Width =655
                    Height =432
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDINOPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"DINO\",[tbxAll])"
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

                    LayoutCachedLeft =11400
                    LayoutCachedTop =24
                    LayoutCachedWidth =12055
                    LayoutCachedHeight =456
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10860
                    Top =24
                    Width =655
                    Height =432
                    FontSize =9
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCUREPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"CURE\",[tbxAll])"
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

                    LayoutCachedLeft =10860
                    LayoutCachedTop =24
                    LayoutCachedWidth =11515
                    LayoutCachedHeight =456
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10320
                    Top =24
                    Width =655
                    Height =432
                    FontSize =9
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCOLMPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"COLM\",[tbxAll])"
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

                    LayoutCachedLeft =10320
                    LayoutCachedTop =24
                    LayoutCachedWidth =10975
                    LayoutCachedHeight =456
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9780
                    Top =24
                    Width =655
                    Height =432
                    FontSize =9
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCAREPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"CARE\",[tbxAll])"
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

                    LayoutCachedLeft =9780
                    LayoutCachedTop =24
                    LayoutCachedWidth =10435
                    LayoutCachedHeight =456
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
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9240
                    Top =24
                    Width =655
                    Height =432
                    FontSize =9
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxBLCAPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"BLCA\",[tbxAll])"
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
                    LayoutCachedWidth =9895
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
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1320
                    Height =312
                    FontSize =9
                    TabIndex =21
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFamily"
                    ControlSource ="Family"
                    StatusBarText ="Target Species name (ITIS species name from tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5580
                    Top =60
                    Width =1380
                    Height =312
                    ColumnWidth =2655
                    FontSize =9
                    TabIndex =23
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Species_Name"
                    ControlSource ="Master_Plant_Code_FK"
                    StatusBarText ="Target Species name (ITIS species name from tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =5580
                    LayoutCachedTop =60
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3420
                    Top =60
                    Width =1980
                    Height =312
                    ColumnWidth =1170
                    FontSize =9
                    TabIndex =24
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbl_Target_Species.Park_Code"
                    ControlSource ="Co_Species"
                    StatusBarText ="Standard park code (CANY, FOBU, etc.)"
                    EventProcPrefix ="tbl_Target_Species_Park_Code"
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =60
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7200
                    Top =60
                    Width =1980
                    Height =312
                    ColumnWidth =2400
                    FontSize =9
                    TabIndex =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCommon"
                    ControlSource ="Master_Common_Name"
                    StatusBarText ="FK to plant master code (tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =7200
                    LayoutCachedTop =60
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2580
                    Top =60
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxBLCA"
                    ControlSource ="=CountInString([ParkPriorities],\"BLCA-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =60
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =60
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCARE"
                    ControlSource ="=CountInString([ParkPriorities],\"CARE-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3180
                    Top =60
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCOLM"
                    ControlSource ="=CountInString([ParkPriorities],\"COLM-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =3180
                    LayoutCachedTop =60
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3480
                    Top =60
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCURE"
                    ControlSource ="=CountInString([ParkPriorities],\"CURE-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =3480
                    LayoutCachedTop =60
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3780
                    Top =60
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDINO"
                    ControlSource ="=CountInString([ParkPriorities],\"DINO-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =3780
                    LayoutCachedTop =60
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4080
                    Top =60
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFOBU"
                    ControlSource ="=CountInString([ParkPriorities],\"FOBU-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =4080
                    LayoutCachedTop =60
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4380
                    Top =60
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxGOSP"
                    ControlSource ="=CountInString([ParkPriorities],\"GOSP-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =60
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4680
                    Top =60
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxZION"
                    ControlSource ="=CountInString([ParkPriorities],\"ZION-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =4680
                    LayoutCachedTop =60
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1380
                    Top =60
                    Width =1980
                    Height =312
                    FontSize =9
                    TabIndex =25
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbl_Target_Species.Target_Year"
                    ControlSource ="utah_species"
                    StatusBarText ="Year (4-digit)"
                    EventProcPrefix ="tbl_Target_Species_Target_Year"
                    GridlineColor =10921638

                    LayoutCachedLeft =1380
                    LayoutCachedTop =60
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Top =60
                    Width =1320
                    Height =312
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTgtYr"
                    ControlSource ="TgtYear"
                    StatusBarText ="Target Species name (ITIS species name from tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedTop =60
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =660
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxRunSumPri1"
                    ControlSource ="=CountInString([ParkPriorities],1)"
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4860
                    Top =120
                    Width =5280
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxAll"
                    ControlSource ="ParkPriorities"
                    GridlineColor =10921638

                    LayoutCachedLeft =4860
                    LayoutCachedTop =120
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =420
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
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
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
                    Width =15480
                    Name ="lnPageFooter"
                    GridlineColor =10921638
                    LayoutCachedWidth =15480
                End
                Begin TextBox
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
                    ControlSource ="=[tbxBLCA]"
                    StatusBarText ="Total # priority 1 (BLCA)"
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =60
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCARE"
                    ControlSource ="=[tbxCARE]"
                    StatusBarText ="Total # priority 1 (CARE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =9960
                    LayoutCachedTop =60
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10500
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCOLM"
                    ControlSource ="=[tbxCOLM]"
                    StatusBarText ="Total # priority 1 (COLM)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10500
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11040
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCURE"
                    ControlSource ="=[tbxCURE]"
                    StatusBarText ="Total # priority 1 (CURE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11040
                    LayoutCachedTop =60
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11580
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumDINO"
                    ControlSource ="=[tbxDINO]"
                    StatusBarText ="Total # priority 1 (DINO)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11580
                    LayoutCachedTop =60
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12120
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumFOBU"
                    ControlSource ="=[tbxFOBU]"
                    StatusBarText ="Total # priority 1 (FOBU)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12120
                    LayoutCachedTop =60
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12600
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumGOSP"
                    ControlSource ="=[tbxGOSP]"
                    StatusBarText ="Total # priority 1 (GOSP)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12600
                    LayoutCachedTop =60
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13140
                    Top =60
                    Width =300
                    Height =270
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumZION"
                    ControlSource ="=[tbxZION]"
                    StatusBarText ="Total # priority 1 (ZION)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13140
                    LayoutCachedTop =60
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =330
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
                    Name ="lblParkPriorities"
                    Caption ="Total # Priority 1 Species by Park ="
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5760
                    LayoutCachedTop =60
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =384
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9480
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueBLCA"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"BLCA-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (BLCA)"
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
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9960
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueCARE"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"CARE-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (CARE)"
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
                    LayoutCachedTop =420
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =720
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10500
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueCOLM"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"COLM-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (COLM)"
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

                    LayoutCachedLeft =10500
                    LayoutCachedTop =420
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =720
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11040
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueCURE"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"CURE-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (CURE)"
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

                    LayoutCachedLeft =11040
                    LayoutCachedTop =420
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =720
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11580
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueDINO"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"DINO-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (DINO)"
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

                    LayoutCachedLeft =11580
                    LayoutCachedTop =420
                    LayoutCachedWidth =11880
                    LayoutCachedHeight =720
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12120
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueFOBU"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"FOBU-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (FOBU)"
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

                    LayoutCachedLeft =12120
                    LayoutCachedTop =420
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =720
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
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12600
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueGOSP"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"GOSP-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (GOSP)"
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

                    LayoutCachedLeft =12600
                    LayoutCachedTop =420
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =720
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
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13140
                    Top =420
                    Width =300
                    Height =270
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueZION"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"ZION-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (ZION)"
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

                    LayoutCachedLeft =13140
                    LayoutCachedTop =420
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =690
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
                    Left =6948
                    Top =420
                    Width =2292
                    Height =288
                    FontSize =10
                    BackColor =16777164
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblUniquePri1"
                    Caption ="Unique Priority 1 Species ="
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6948
                    LayoutCachedTop =420
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =708
                    BackThemeColorIndex =-1
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
' MODULE:       Report_rptTgtSpeciesListAnnual
' Description:  Load species list to target species list functions and routines
'
' Source/date:  Bonnie Campbell, 4/7/2015
' Revisions:    BLC - 4/7/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Report_Open
' Description:  Actions for when report opens
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, April 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/7/2015 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)

On Error GoTo Err_Handler
'http://stackoverflow.com/questions/11477297/giving-an-alias-to-a-subquery-containing-a-join-in-access

'Dim ParkPriorities As Variant

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
    
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[Report_rptTgtSpeciesListAnnual])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Report_Load
' Description:  Actions for when report is loaded
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, April 8, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/8/2015 - initial version
' ---------------------------------
Private Sub Report_Load()

On Error GoTo Err_Handler

    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Load[Report_rptTgtSpeciesListAnnual])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Detail_Format
' Description:  Actions for when report detail is formatted
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, April 8, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/8/2015 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

On Error GoTo Err_Handler

    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[Report_rptTgtSpeciesListAnnual])"
    End Select
    Resume Exit_Sub
End Sub
