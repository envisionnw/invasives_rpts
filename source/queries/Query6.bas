dbMemo "SQL" ="SELECT DISTINCT ts.ID, ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect_ID, ts"
    ".Transect, ts.Area, ts.E_Coord, ts.N_Coord, ts.Position_m, ts.ColName, q.ID, ts."
    "Quadrat, ts.IsSampled, ts.NoExotics, sc.PlantCode, sc.IsDead, sc.PercentCover\015"
    "\012FROM (Transect_Select_LIMITED_NO_SPECIES AS ts LEFT JOIN Quadrat AS q ON q.T"
    "ransect_ID = ts.Transect_ID) LEFT JOIN SpeciesCover AS sc ON sc.Quadrat_ID = q.I"
    "D;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Query6].[Unit_Code]=\"GOSP\"))) AND ([Query6].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbb532edee2d17340932c206834fac441
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Position_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.ColName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Quadrat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.IsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.NoExotics"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.PercentCover"
        dbLong "AggregateType" ="-1"
    End
End
