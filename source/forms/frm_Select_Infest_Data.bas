﻿Version =20
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
    ItemSuffix =14
    Left =4080
    Top =2535
    Right =11025
    Bottom =6660
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3d34192b53bbe340
    End
    Caption ="Select Infestation Data"
    DatasheetFontName ="Arial"
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
                    Left =1440
                    Top =240
                    Width =4335
                    Height =420
                    FontSize =16
                    FontWeight =700
                    Name ="Label3"
                    Caption ="Infestation Listing"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4560
                    Top =2580
                    Width =1350
                    Height =299
                    Name ="ButtonClose"
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
                    RowSource ="SELECT qry_sel_Infest_Year.Visit_Year FROM qry_sel_Infest_Year; "
                    ColumnWidths ="2820"

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
                    Left =1200
                    Top =2580
                    Width =1350
                    Height =299
                    TabIndex =3
                    Name ="ButtonReport"
                    Caption ="Preview Report"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2880
                    Top =2580
                    Width =1350
                    Height =300
                    TabIndex =4
                    Name ="ButtonQuery"
                    Caption ="Run as Query"
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

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub


Private Sub Park_Code_AfterUpdate()
  If Not IsNull(Me!Park_Code) Then
    Me!Visit_Year.RowSource = "SELECT Visit_Year FROM qry_sel_Infest_Year WHERE [Unit_Code] = '" & Me!Park_Code & "' ORDER BY Visit_Year"
    Me.Refresh
  End If
End Sub
Private Sub ButtonReport_Click()
On Error GoTo Err_ButtonReport_Click

    Dim stDocName As String
    Dim stWhere As String
    Dim stOpenArg As String

    If IsNull(Me!Park_Code) Or IsNull(Me!Visit_Year) Then
      MsgBox "You must select both park and year.", , "Infestation Listing"
      Exit Sub
    End If
    stOpenArg = Me!Park_Code & Me!Visit_Year
    stWhere = "[Unit_Code] = '" & Me!Park_Code & "' AND Visit_Year = " & Me!Visit_Year
    stDocName = "rpt_Infestation"
    DoCmd.OpenReport stDocName, acPreview, , stWhere, , stOpenArg

Exit_ButtonReport_Click:
    Exit Sub

Err_ButtonReport_Click:
    MsgBox Err.Description
    Resume Exit_ButtonReport_Click
    
End Sub
Private Sub ButtonQuery_Click()
On Error GoTo Err_ButtonQuery_Click

    Dim stDocName As String
    
    If IsNull(Me!Park_Code) Or IsNull(Me!Visit_Year) Then
      MsgBox "You must select both park and year.", , "Infestation Query"
      Exit Sub
    End If

    stDocName = "qry_Infest_Param"
    DoCmd.OpenQuery stDocName, acNormal, acEdit

Exit_ButtonQuery_Click:
    Exit Sub

Err_ButtonQuery_Click:
    MsgBox Err.Description
    Resume Exit_ButtonQuery_Click
    
End Sub
