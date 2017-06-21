Operation =1
Option =0
Begin InputTables
    Name ="Route_SpeciesCover_Crosstab_PctCover"
    Name ="Route_SpeciesCover_Crosstab_TCount"
    Name ="Route_SpeciesCover_Crosstab_SE"
End
Begin OutputColumns
    Expression ="Route_SpeciesCover_Crosstab_TCount.Unit_Code"
    Expression ="Route_SpeciesCover_Crosstab_TCount.Visit_Year"
    Expression ="Route_SpeciesCover_Crosstab_PctCover.Species"
    Expression ="Route_SpeciesCover_Crosstab_PctCover.Master_Common_Name"
    Expression ="Route_SpeciesCover_Crosstab_PctCover.[Alive?]"
    Expression ="Route_SpeciesCover_Crosstab_SE.[4WD Road 1 (3) TCount]"
    Expression ="Route_SpeciesCover_Crosstab_PctCover.[4WD Road 1 (3) PctCover]"
    Expression ="Route_SpeciesCover_Crosstab_TCount.[4WD Road 1 (3) SE]"
End
Begin Joins
    LeftTable ="Route_SpeciesCover_Crosstab_TCount"
    RightTable ="Route_SpeciesCover_Crosstab_PctCover"
    Expression ="Route_SpeciesCover_Crosstab_TCount.Unit_Code = Route_SpeciesCover_Crosstab_PctCo"
        "ver.Unit_Code"
    Flag =1
    LeftTable ="Route_SpeciesCover_Crosstab_TCount"
    RightTable ="Route_SpeciesCover_Crosstab_PctCover"
    Expression ="Route_SpeciesCover_Crosstab_TCount.Visit_Year = Route_SpeciesCover_Crosstab_PctC"
        "over.Visit_Year"
    Flag =1
    LeftTable ="Route_SpeciesCover_Crosstab_TCount"
    RightTable ="Route_SpeciesCover_Crosstab_PctCover"
    Expression ="Route_SpeciesCover_Crosstab_TCount.Species = Route_SpeciesCover_Crosstab_PctCove"
        "r.Species"
    Flag =1
    LeftTable ="Route_SpeciesCover_Crosstab_TCount"
    RightTable ="Route_SpeciesCover_Crosstab_PctCover"
    Expression ="Route_SpeciesCover_Crosstab_TCount.[Alive?] = Route_SpeciesCover_Crosstab_PctCov"
        "er.[Alive?]"
    Flag =1
    LeftTable ="Route_SpeciesCover_Crosstab_PctCover"
    RightTable ="Route_SpeciesCover_Crosstab_SE"
    Expression ="Route_SpeciesCover_Crosstab_PctCover.Unit_Code = Route_SpeciesCover_Crosstab_SE."
        "Unit_Code"
    Flag =1
    LeftTable ="Route_SpeciesCover_Crosstab_PctCover"
    RightTable ="Route_SpeciesCover_Crosstab_SE"
    Expression ="Route_SpeciesCover_Crosstab_PctCover.Visit_Year = Route_SpeciesCover_Crosstab_SE"
        ".Visit_Year"
    Flag =1
    LeftTable ="Route_SpeciesCover_Crosstab_PctCover"
    RightTable ="Route_SpeciesCover_Crosstab_SE"
    Expression ="Route_SpeciesCover_Crosstab_PctCover.Species = Route_SpeciesCover_Crosstab_SE.Sp"
        "ecies"
    Flag =1
    LeftTable ="Route_SpeciesCover_Crosstab_PctCover"
    RightTable ="Route_SpeciesCover_Crosstab_SE"
    Expression ="Route_SpeciesCover_Crosstab_PctCover.[Alive?] = Route_SpeciesCover_Crosstab_SE.["
        "Alive?]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x3f9ef670ce043f4f8390f9b81fccfa01
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Route_SpeciesCover_Crosstab_TCount.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_SpeciesCover_Crosstab_TCount.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_SpeciesCover_Crosstab_PctCover.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_SpeciesCover_Crosstab_PctCover.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_SpeciesCover_Crosstab_PctCover.[Alive?]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_SpeciesCover_Crosstab_SE.[4WD Road 1 (3) TCount]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_SpeciesCover_Crosstab_PctCover.[4WD Road 1 (3) PctCover]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_SpeciesCover_Crosstab_TCount.[4WD Road 1 (3) SE]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1603
    Bottom =876
    Left =-1
    Top =-1
    Right =1571
    Bottom =552
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =296
        Top =13
        Right =567
        Bottom =252
        Top =0
        Name ="Route_SpeciesCover_Crosstab_PctCover"
        Name =""
    End
    Begin
        Left =646
        Top =13
        Right =928
        Bottom =266
        Top =0
        Name ="Route_SpeciesCover_Crosstab_TCount"
        Name =""
    End
    Begin
        Left =28
        Top =18
        Right =260
        Bottom =553
        Top =0
        Name ="Route_SpeciesCover_Crosstab_SE"
        Name =""
    End
End
