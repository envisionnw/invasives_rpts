Operation =6
Option =0
Begin InputTables
    Name ="Route_Species"
End
Begin OutputColumns
    Expression ="Route_Species.[Unit_Code]"
    GroupLevel =2
    Expression ="Route_Species.[Visit_Year]"
    GroupLevel =2
    Expression ="Route_Species.[Route]"
    GroupLevel =2
    Expression ="Route_Species.[ColRouteTransects]"
    GroupLevel =1
    Alias ="MinOfRouteAverageCover"
    Expression ="Min(Route_Species.[RouteAverageCover])"
End
Begin Groups
    Expression ="Route_Species.[Unit_Code]"
    GroupLevel =2
    Expression ="Route_Species.[Visit_Year]"
    GroupLevel =2
    Expression ="Route_Species.[Route]"
    GroupLevel =2
    Expression ="Route_Species.[ColRouteTransects]"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x79f7dd271f5e6e46a472bb0d63a0d379
End
Begin
    Begin
        dbText "Name" ="Route_Species.[Unit_Code]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1284
    Bottom =672
    Left =-1
    Top =-1
    Right =1252
    Bottom =348
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Route_Species"
        Name =""
    End
End
