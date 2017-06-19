dbMemo "SQL" ="SELECT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, ts.E_Coord, ts.N_Coor"
    "d, ts.PlantCode, ts.IsDead, SUM(ts.PercentCover) AS TotalCover\015\012FROM Trans"
    "ect_Select_LIMITED_ESP_SpeciesCover_Species AS ts\015\012GROUP BY ts.Unit_Code, "
    "ts.Visit_Year, ts.Route, ts.Transect, ts.E_Coord, ts.N_Coord, ts.PlantCode, ts.I"
    "sDead\015\012ORDER BY ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, ts.E_C"
    "oord, ts.N_Coord, ts.PlantCode, ts.IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x610d7aaee573e94fad6b0427e2c85fe3
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="((([Transect_Select_TotalCover].[Unit_Code]=\"GOSP\"))) AND ([Transect_Select_To"
    "talCover].[Visit_Year]=2016)"
Begin
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCover"
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
End
