dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Location_ID, ts.Route, ts.Transe"
    "ct_ID, ts.Transect, ts.Area, ts.E_Coord, ts.N_Coord, ts.Quadrat_ID, ts.Quadrat, "
    "ts.IsSampled, ts.NoExotics, esp.Position_m, esp.ColName\015\012FROM Transect_Sel"
    "ect AS ts LEFT JOIN EventSamplePosition AS esp ON (esp.Location_ID = ts.Location"
    "_ID) AND (esp.Event_ID = ts.Event_ID)\015\012WHERE esp.Quadrat = ts.Quadrat\015\012"
    "ORDER BY ts.Route, ts.Transect, ts.Quadrat;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBinary "GUID" = Begin
    0xd642c5240f81384c8f8d2ddbc4d378c1
End
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
