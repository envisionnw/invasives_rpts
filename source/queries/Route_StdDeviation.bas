dbMemo "SQL" ="SELECT d.Unit_Code, d.Visit_Year, d.Route, d.Species, MIN(d.Master_Common_Name) "
    "AS Master_Common_Name, d.IsDead, MIN(d.TransectsSampled) AS TransectsSampled, MI"
    "N(d.TransectsDetected) AS TransectsDetected, MIN(d.TotalCover) AS TotalCover, MI"
    "N(d.RouteAverageCover) AS RouteAverageCover, MIN(d.TotalDevSquared) AS TotalDevS"
    "quared, MIN(IIF(d.TransectsSampled = 1, NULL, SQR(d.TotalDevSquared/(d.Transects"
    "Sampled -1)))) AS StdDeviation\015\012FROM Route_AverageCover_Deviations_Aggrega"
    "te AS d\015\012GROUP BY d.Unit_Code, d.Visit_Year, d.Route, d.Species, d.IsDead\015"
    "\012ORDER BY d.Unit_Code, d.Visit_Year, d.Route, d.Species, d.IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5e4a309e69c2f744976f57cf479392b7
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalDevSquared"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="StdDeviation"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectAverageCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="615"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="d.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsDetected"
        dbLong "AggregateType" ="-1"
    End
End
