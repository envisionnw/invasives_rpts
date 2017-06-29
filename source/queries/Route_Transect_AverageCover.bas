dbMemo "SQL" ="SELECT l.Unit_Code, l.Visit_Year, l.Route, MIN(l.Area) AS Area, l.Transect, MIN("
    "IIF(l.E_Coord IS NULL, ac.E_Coord, l.E_Coord)) AS E_Coord, MIN(IIF(l.N_Coord IS "
    "NULL, ac.N_Coord, l.N_Coord)) AS N_Coord, l.PlantCode, l.IsDead, MIN(ac.TotalCov"
    "er) AS TotalCover, MIN(ac.QuadratsSampled) AS QuadratsSampled, MIN(ac.TransectsS"
    "ampled) AS TransectsSampled, MIN(IIF(ac.TransectAverageCover IS NULL, 0, ac.Tran"
    "sectAverageCover)) AS TransectAverageCover\015\012FROM Route_Transect_Species_Li"
    "st AS l LEFT JOIN Transect_AverageCover AS ac ON (ac.Unit_Code = l.Unit_Code) AN"
    "D (ac.Visit_Year = l.Visit_Year) AND (ac.Route = l.Route) AND (ac.Transect = l.T"
    "ransect) AND (ac.PlantCode = l.PlantCode) AND (ac.IsDead = l.IsDead)\015\012GROU"
    "P BY l.Unit_Code, l.Visit_Year, l.Route, l.Transect, l.PlantCode, l.IsDead\015\012"
    "ORDER BY l.Unit_Code, l.Visit_Year, l.Route, MIN(l.Area), l.Transect, l.PlantCod"
    "e, l.IsDead;\015\012"
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
    Begin
        dbText "Name" ="Area"
        dbLong "AggregateType" ="-1"
    End
End
