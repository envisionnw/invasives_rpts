dbMemo "SQL" ="SELECT MIN(d.Unit_Code) AS Unit_Code, MIN(d.Visit_Year) AS Visit_Year, MIN(d.Rou"
    "te) AS Route, MIN(d.Species) AS Species, MIN(d.IsDead) AS IsDead, MIN(d.SampledQ"
    "uadrats) AS SampledQuadrats, MIN(d.TotalCover) AS TotalCover, MIN(d.AverageCover"
    ") AS AverageCover, MIN(d.RouteAverageCover) AS RouteAverageCover, MIN(d.TotalDev"
    "Squared) AS TotalDevSquared, MIN(IIF(d.SampledQuadrats = 1, NULL, SQR(d.TotalDev"
    "Squared/(d.SampledQuadrats -1)))) AS StdDeviation\015\012FROM Route_AverageCover"
    "_Deviations_Aggregate AS d\015\012GROUP BY d.Unit_Code, d.Visit_Year, d.Route, d"
    ".Species\015\012ORDER BY d.Unit_Code, d.Visit_Year, d.Route, d.Species;\015\012"
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
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SampledQuadrats"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalDevSquared"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StdDeviation"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
