dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Area, ts.Species, MIN("
    "ts.Master_Common_Name) AS Master_Common_Name, MIN(ts.PlantCode) AS PlantCode, ts"
    ".IsDead, COUNT(ts.Transect) AS TransectsDetected, SUM(ts.PercentCover) AS RouteT"
    "otalCover\015\012FROM Transect_Select_SpeciesCover_Aggregate AS ts\015\012GROUP "
    "BY ts.Unit_Code, ts.Visit_Year, ts.Area, ts.Route, ts.Species, ts.IsDead\015\012"
    "ORDER BY ts.Unit_Code, ts.Visit_Year, ts.Area, ts.Route, ts.Species, ts.IsDead;\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Transect_Select_SpeciesCover_Aggregate_TransectsDetected].[Unit_Code]=\"CARE"
    "\"))) AND ([Transect_Select_SpeciesCover_Aggregate_TransectsDetected].[Visit_Yea"
    "r]=2015)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x02ffd638bfaa934ca1da82f2461db66d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="825"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ts.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="615"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TransectsDetected"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RouteTotalCover"
        dbInteger "ColumnWidth" ="1290"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="420"
        dbBoolean "ColumnHidden" ="0"
    End
End
