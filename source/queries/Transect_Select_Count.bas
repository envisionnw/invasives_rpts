dbMemo "SQL" ="SELECT ts.ID, ts.Unit_Code, ts.Visit_Year, ts.Plot_ID, ts.Transect_ID, ts.Transe"
    "ct, ts.Area, ts.E_Coord, ts.N_Coord, ts.Species, ts.Master_Common_Name, ts.IsDea"
    "d, Count(ts.IsSampled) AS QuadratsSampled, SUM(ts.PercentCover) AS TotalCover, ("
    "TotalCover/QuadratsSampled) AS AverageCover\015\012FROM Transect_Select AS ts\015"
    "\012GROUP BY ts.ID, ts.Unit_Code, ts.Visit_Year, ts.Plot_ID, ts.Transect_ID, ts."
    "Transect, ts.Area, ts.E_Coord, ts.N_Coord, ts.Species, ts.Master_Common_Name, ts"
    ".IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x020699d4ec5e7b4fa8e3581aa040db60
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="(((Transect_Select_Count.Unit_Code=\"ZION\"))) And (Transect_Select_Count.Visit_"
    "Year=2012)"
Begin
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="ts.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuadratsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc49a2c5c70b52347b70d98e4371e4d6e
        End
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x80677be8e736a548a68d0360c0609263
        End
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Transect_ID"
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
        dbText "Name" ="AverageCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9e2f7982c8522f43b94232e4e7f30729
        End
    End
End
