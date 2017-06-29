Operation =1
Option =0
Where ="(((Route_Transects_With_Species.Transect) Is Null))"
Begin InputTables
    Name ="Route_Transects_SpeciesCover"
    Name ="Route_Transects_With_Species"
End
Begin OutputColumns
    Expression ="Route_Transects_SpeciesCover.Transect"
End
Begin Joins
    LeftTable ="Route_Transects_SpeciesCover"
    RightTable ="Route_Transects_With_Species"
    Expression ="Route_Transects_SpeciesCover.[Transect] = Route_Transects_With_Species.[Transect"
        "]"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe15828686198d144a704ef1af67d66fc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Route_Transects_SpeciesCover].[Transect]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =831
    Bottom =736
    Left =-1
    Top =-1
    Right =799
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
        Name ="Route_Transects_SpeciesCover"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Route_Transects_With_Species"
        Name =""
    End
End
