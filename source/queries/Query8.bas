dbMemo "SQL" ="SELECT td.Unit_Code, td.Visit_Year, td.Route, td.Area, td.PlantCode, td.IsDead, "
    "MIN(td.RouteTotalCover) AS RouteTotalCover, MIN(td.TransectsDetected) AS Transec"
    "tsDetected, MIN(ts.TransectsSampled) AS SampledTransects, MIN(td.RouteTotalCover"
    " / ts.TransectsSampled) AS AverageCover\015\012FROM Transect_Select_SpeciesCover"
    "_Aggregate_TransectsDetected AS td INNER JOIN Route_TransectsSampled AS ts ON (t"
    "s.Route = td.Route) AND (ts.Visit_Year = td.Visit_Year) AND (ts.Unit_Code = td.U"
    "nit_Code)\015\012GROUP BY td.Unit_Code, td.Visit_Year, td.Route, td.Area, td.Pla"
    "ntCode, td.IsDead\015\012ORDER BY td.Unit_Code, td.Visit_Year, td.Route, td.Area"
    ", td.PlantCode, td.IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Query8].[Unit_Code]=\"CARE\"))) AND ([Query8].[Visit_Year]=2015)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd90bd4cfba6de7479915b8845ea39204
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="td.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RouteTotalCover"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsDetected"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SampledTransects"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageCover"
        dbLong "AggregateType" ="-1"
    End
End
