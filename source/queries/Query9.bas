dbMemo "SQL" ="SELECT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Area, ts.Transect, ts.Species, "
    "MIN(ts.Master_Common_Name) AS Master_Common_Name, MIN(ts.PlantCode) AS PlantCode"
    ", ts.IsDead, COUNT(ts.Transect) AS TransectsDetected, SUM(ts.PercentCover) AS Ro"
    "uteTotalCover\015\012FROM Transect_Select_SpeciesCover_Aggregate AS ts\015\012GR"
    "OUP BY ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Area, ts.Transect, ts.Species, "
    "ts.IsDead\015\012ORDER BY ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Area, ts.Tra"
    "nsect, ts.Species, ts.IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Query9].[Unit_Code]=\"CARE\"))) AND ([Query9].[Visit_Year]=2015)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe9b754b35d8b964e97778508ac607ea4
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
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
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
    End
    Begin
        dbText "Name" ="TransectsDetected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RouteTotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
    End
End
