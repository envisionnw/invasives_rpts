dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, ts.Area, ts."
    "E_Coord, ts.N_Coord, ts.Species, ts.Master_Common_Name, IIF(ts.IsDead = 1, \"N\""
    ", \"Y\") AS [Alive?], (tc.TotalCover/qs.SampledQuadrats) AS AverageCover\015\012"
    "FROM (Transect_Select_SpeciesCover AS ts INNER JOIN Transect_Select_QuadratsSamp"
    "led AS qs ON (qs.Unit_Code = ts.Unit_Code) AND (qs.Visit_Year = ts.Visit_Year) A"
    "ND (qs.Route = ts.Route) AND (qs.Transect = ts.Transect) AND (qs.E_Coord = ts.E_"
    "Coord) AND (qs.N_Coord = ts.N_Coord)) INNER JOIN Transect_Select_TotalCover AS t"
    "c ON (tc.Unit_Code = ts.Unit_Code) AND (tc.Visit_Year = ts.Visit_Year) AND (tc.R"
    "oute = ts.Route) AND (tc.Transect = ts.Transect) AND (tc.E_Coord = ts.E_Coord) A"
    "ND (tc.N_Coord = ts.N_Coord) AND (tc.PlantCode = ts.PlantCode) AND (tc.IsDead = "
    "ts.IsDead)\015\012ORDER BY ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, t"
    "s.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="((([Transect_Data].[Unit_Code]=\"GOSP\"))) AND ([Transect_Data].[Visit_Year]=201"
    "6)"
dbBinary "GUID" = Begin
    0xae80469a2679cb4bba2ec79d418ed46e
End
Begin
    Begin
        dbText "Name" ="Alive?"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa6075aa7909caa4d8b2e6d651a0104f3
        End
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
        dbText "Name" ="ts.N_Coord"
        dbLong "AggregateType" ="-1"
    End
End
