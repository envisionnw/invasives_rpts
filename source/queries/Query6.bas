Operation =1
Option =0
Begin InputTables
    Name ="Route_StdError"
    Name ="Route_Transects"
End
Begin OutputColumns
    Expression ="Route_StdError.Unit_Code"
    Expression ="Route_StdError.Visit_Year"
    Expression ="Route_StdError.Route"
    Expression ="Route_Transects.TransectCount"
    Expression ="Route_StdError.TransectsSampled"
    Expression ="Route_StdError.Species"
    Expression ="Route_StdError.Master_Common_Name"
    Expression ="Route_StdError.IsDead"
    Expression ="Route_StdError.AverageCover"
    Expression ="Route_StdError.StdDeviation"
    Expression ="Route_StdError.StdError"
End
Begin Joins
    LeftTable ="Route_Transects"
    RightTable ="Route_StdError"
    Expression ="Route_Transects.Unit_Code = Route_StdError.Unit_Code"
    Flag =3
    LeftTable ="Route_StdError"
    RightTable ="Route_Transects"
    Expression ="Route_StdError.Visit_Year = Route_Transects.Visit_Year"
    Flag =2
    LeftTable ="Route_Transects"
    RightTable ="Route_StdError"
    Expression ="Route_Transects.Route = Route_StdError.Route"
    Flag =3
End
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa978f7e45f3edf43a9872ea2c4e39639
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Route_StdError.StdDeviation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_StdError.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_StdError.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_StdError.Route"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Route_StdError.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_StdError.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_StdError.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_StdError.AverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_StdError.StdError"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route_StdError.TransectsSampled"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Route_Transects.TransectCount"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =38
    Top =-2
    Right =833
    Bottom =640
    Left =-1
    Top =-1
    Right =763
    Bottom =355
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =229
        Bottom =310
        Top =0
        Name ="Route_StdError"
        Name =""
    End
    Begin
        Left =281
        Top =11
        Right =425
        Bottom =155
        Top =0
        Name ="Route_Transects"
        Name =""
    End
End
