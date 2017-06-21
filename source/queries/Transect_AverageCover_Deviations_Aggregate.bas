dbMemo "SQL" ="SELECT MIN(d.Unit_Code) AS Unit_Code, MIN(d.Visit_Year) AS Visit_Year, MIN(d.Rou"
    "te) AS Route, MIN(d.Species) AS Species, MIN(d.Master_Common_Name) AS Master_Com"
    "mon_Name, MIN(d.IsDead) AS IsDead, MIN(d.TransectsSampled) AS TransectsSampled, "
    "MIN(d.TotalCover) AS TotalCover, MIN(d.TransectAverageCover) AS TransectAverageC"
    "over, MIN(d.RouteAverageCover) AS RouteAverageCover, SUM(d.DeviationSquared) AS "
    "TotalDevSquared\015\012FROM Transect_AverageCover_Deviations AS d\015\012GROUP B"
    "Y d.Unit_Code, d.Visit_Year, d.Route, d.Species\015\012ORDER BY d.Unit_Code, d.V"
    "isit_Year, d.Route, d.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x850af98add43c0468778af91387d8639
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
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectAverageCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TotalDevSquared"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
    End
End
