dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, ts.Area, ts."
    "E_Coord, ts.N_Coord, ts.Species, ts.Master_Common_Name, IIF(ts.IsDead = 1, \"N\""
    ", \"Y\") AS [Alive?], qs.SampledQuadrats, tc.TotalCover, (tc.TotalCover/qs.Sampl"
    "edQuadrats) AS AverageCover, ac.AverageCover AS RouteAverageCover, (RouteAverage"
    "Cover - AverageCover) AS Deviation\015\012FROM ((Transect_Select_LIMITED_ESP_Spe"
    "ciesCover_Species AS ts INNER JOIN Transect_Select_QuadratsSampled AS qs ON (qs."
    "Route = ts.Route) AND (qs.Visit_Year = ts.Visit_Year) AND (qs.Unit_Code = ts.Uni"
    "t_Code)) INNER JOIN Transect_Select_TotalCover AS tc ON (tc.IsDead = ts.IsDead) "
    "AND (tc.PlantCode = ts.PlantCode) AND (tc.N_Coord = ts.N_Coord) AND (tc.E_Coord "
    "= ts.E_Coord) AND (tc.Transect = ts.Transect) AND (tc.Route = ts.Route) AND (tc."
    "Visit_Year = ts.Visit_Year) AND (tc.Unit_Code = ts.Unit_Code)) LEFT JOIN Route_A"
    "verageCover AS ac ON (ac.IsDead = ts.IsDead) AND (ac.PlantCode = ts.PlantCode) A"
    "ND (ac.Route = ts.Route) AND (ac.Visit_Year = ts.Visit_Year) AND (ac.Unit_Code ="
    " ts.Unit_Code)\015\012ORDER BY ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transec"
    "t, ts.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x7c81657ccdc9954dac94fe116eac7a87
End
Begin
    Begin
        dbText "Name" ="ts.N_Coord"
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
        dbText "Name" ="Alive?"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qs.SampledQuadrats"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageCover"
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
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deviation"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
