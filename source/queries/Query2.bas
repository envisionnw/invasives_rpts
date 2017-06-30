dbMemo "SQL" ="SELECT DISTINCT se.Unit_Code, se.Visit_Year, se.Route, se.Species, se.Master_Com"
    "mon_Name, IIF(se.IsDead = 1,'N','Y') AS [Alive?], se.TransectsSampled, se.Transe"
    "ctsDetected, se.TotalCover, se.RouteAverageCover, se.StdDeviation, se.StdError, "
    "rt.RouteTruncated & \" (\" & rt.TransectCount & \") TCount\" AS ColRouteTransect"
    "s, rt.RouteTruncated & \" (\" & rt.TransectCount & \") PctCover\" AS ColRouteCov"
    "er, rt.RouteTruncated & \" (\" & rt.TransectCount & \") SE\" AS ColRouteStdError"
    "\015\012FROM Route_StdError AS se LEFT JOIN Route_Transects AS rt ON (rt.Route ="
    " se.Route) AND (rt.Visit_Year = se.Visit_Year) AND (rt.Unit_Code = se.Unit_Code)"
    "\015\012ORDER BY se.Unit_Code, se.Visit_Year, se.Route, se.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1377a2f600473645ac510e07372715dd
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tc.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalTransectAverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SampledTransects"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
    End
End
