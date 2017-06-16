dbMemo "SQL" ="SELECT ts.ID, ts.Unit_Code, ts.Visit_Year, ts.Location_ID, ts.Route, ts.Transect"
    "_ID, ts.Transect, ts.Area, ts.E_Coord, ts.N_Coord, ts.Quadrat_ID, ts.Quadrat, ts"
    ".IsSampled, ts.NoExotics, esp.Position_m, esp.ColName\015\012FROM Transect_Selec"
    "t_LIMITED_new AS ts LEFT JOIN EventSamplePosition AS esp ON esp.Location_ID = ts"
    ".Location_ID\015\012WHERE esp.Quadrat = ts.Quadrat\015\012ORDER BY ts.Route, ts."
    "Transect, ts.Quadrat_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Query8].[Unit_Code]=\"GOSP\"))) AND ([Query8].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x774708fb44add14f999f1be2188c1821
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
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
        dbText "Name" ="ts.Location_ID"
        dbLong "AggregateType" ="-1"
    End
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
        dbText "Name" ="ts.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Quadrat_ID"
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
        dbText "Name" ="esp.Position_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="esp.ColName"
        dbLong "AggregateType" ="-1"
    End
End
