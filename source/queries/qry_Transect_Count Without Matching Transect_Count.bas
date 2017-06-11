Operation =1
Option =0
Where ="(((Transect_Count.Route) Is Null))"
Begin InputTables
    Name ="qry_Transect_Count"
    Name ="Transect_Count"
End
Begin OutputColumns
    Expression ="qry_Transect_Count.Unit_Code"
    Expression ="qry_Transect_Count.Route"
    Expression ="qry_Transect_Count.Visit_Year"
    Expression ="qry_Transect_Count.CountOfTransect"
End
Begin Joins
    LeftTable ="qry_Transect_Count"
    RightTable ="Transect_Count"
    Expression ="qry_Transect_Count.[Route] = Transect_Count.[Route]"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x128876dd39afc0458111678adaaa0a6c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[qry_Transect_Count].[Unit_Code]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qry_Transect_Count].[Route]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qry_Transect_Count].[Visit_Year]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qry_Transect_Count].[CountOfTransect]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =117
    Top =184
    Right =959
    Bottom =823
    Left =-1
    Top =-1
    Right =810
    Bottom =271
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qry_Transect_Count"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Transect_Count"
        Name =""
    End
End
