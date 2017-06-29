dbMemo "SQL" ="SELECT Unit_Code, Visit_Year, Route, MIN(Area) AS Area, PlantCode, IsDead, COUNT"
    "(Transect) AS TransectsDetected\015\012FROM Route_Transect_AverageCover\015\012W"
    "HERE TotalCover IS NOT NULL\015\012GROUP BY Unit_Code, Visit_Year, Route, PlantC"
    "ode, IsDead\015\012ORDER BY Unit_Code, Visit_Year, Route, MIN(Area), PlantCode, "
    "IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x275c138dce50344b904d3f2af5d335ee
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsDetected"
        dbLong "AggregateType" ="-1"
    End
End
