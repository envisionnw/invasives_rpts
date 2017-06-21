dbMemo "SQL" ="SELECT MIN(d.Unit_Code) AS Unit_Code, MIN(d.Visit_Year) AS Visit_Year, MIN(d.Rou"
    "te) AS Route, MIN(d.Species) AS Species, MIN(d.Master_Common_Name) AS Master_Com"
    "mon_Name, MIN(d.IsDead) AS IsDead, MIN(d.TransectsSampled) AS TransectsSampled, "
    "MIN(d.TotalCover) AS TotalCover, MIN(d.AverageCover) AS AverageCover, MIN(d.Rout"
    "eAverageCover) AS RouteAverageCover, SUM(d.DeviationSquared) AS TotalDevSquared\015"
    "\012FROM Route_AverageCover_Deviations AS d\015\012GROUP BY d.Unit_Code, d.Visit"
    "_Year, d.Route, d.Species\015\012ORDER BY d.Unit_Code, d.Visit_Year, d.Route, d."
    "Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4d7465baefec3941a1d7aae264055f3e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="((([Route_AverageCover_Deviations_Aggregate].[Unit_Code]=\"GOSP\"))) AND ([Route"
    "_AverageCover_Deviations_Aggregate].[Visit_Year]=2016)"
Begin
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2790"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2445"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TotalDevSquared"
        dbInteger "ColumnWidth" ="2475"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
End
