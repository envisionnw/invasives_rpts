dbMemo "SQL" ="SELECT l.Unit_Code, l.Visit_Year, l.Route, l.Area, l.Transect, MIN(IIF(l.E_Coord"
    " IS NULL, ac.E_Coord, l.E_Coord)) AS E_Coord, MIN(IIF(l.N_Coord IS NULL, ac.N_Co"
    "ord, l.N_Coord)) AS N_Coord, l.PlantCode, l.IsDead, MIN(ac.TotalCover) AS TotalC"
    "over, MIN(ac.QuadratsSampled) AS QuadratsSampled, MIN(ac.TransectsSampled) AS Tr"
    "ansectsSampled, MIN(IIF(ac.TransectAverageCover IS NULL, 0, ac.TransectAverageCo"
    "ver)) AS TransectAverageCover\015\012FROM Route_Transect_Species_List AS l LEFT "
    "JOIN Transect_AverageCover AS ac ON (ac.IsDead = l.IsDead) AND (ac.PlantCode = l"
    ".PlantCode) AND (ac.Transect = l.Transect) AND (ac.Route = l.Route) AND (ac.Visi"
    "t_Year = l.Visit_Year) AND (ac.Unit_Code = l.Unit_Code)\015\012GROUP BY l.Unit_C"
    "ode, l.Visit_Year, l.Route, l.Area, l.Transect, l.PlantCode, l.IsDead\015\012ORD"
    "ER BY l.Unit_Code, l.Visit_Year, l.Route, l.Area, l.Transect, l.PlantCode, l.IsD"
    "ead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4a9de893e963e046a3d4b88e89eeb746
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="l.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuadratsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectAverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Area"
        dbLong "AggregateType" ="-1"
    End
End
