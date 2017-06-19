dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, ts.Area, ts."
    "E_Coord, ts.N_Coord, ts.Species, ts.Master_Common_Name, IIF(ts.IsDead = 1, \"N\""
    ", \"Y\") AS [Alive?], qs.SampledQuadrats, tc.TotalCover, (tc.TotalCover/qs.Sampl"
    "edQuadrats) AS AverageCover\015\012FROM (Transect_Select_LIMITED_ESP_SpeciesCove"
    "r_Species AS ts INNER JOIN Transect_Select_QuadratsSampled AS qs ON (qs.N_Coord "
    "= ts.N_Coord) AND (qs.E_Coord = ts.E_Coord) AND (qs.Transect = ts.Transect) AND "
    "(qs.Route = ts.Route) AND (qs.Visit_Year = ts.Visit_Year) AND (qs.Unit_Code = ts"
    ".Unit_Code)) INNER JOIN Transect_Select_TotalCover AS tc ON (tc.IsDead = ts.IsDe"
    "ad) AND (tc.PlantCode = ts.PlantCode) AND (tc.N_Coord = ts.N_Coord) AND (tc.E_Co"
    "ord = ts.E_Coord) AND (tc.Transect = ts.Transect) AND (tc.Route = ts.Route) AND "
    "(tc.Visit_Year = ts.Visit_Year) AND (tc.Unit_Code = ts.Unit_Code)\015\012ORDER B"
    "Y ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, ts.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Query4].[Unit_Code]=\"GOSP\"))) AND ([Query4].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8acef91ac3e5c1428052f4902617c8a1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "OrderBy" ="[Query4].[Transect]"
Begin
    Begin
        dbText "Name" ="ts.E_Coord"
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
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ts.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qs.SampledQuadrats"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SampledQuadrats"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.ID"
        dbInteger "ColumnWidth" ="5385"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuadratsSampled"
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
        dbText "Name" ="ts.Position_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.ColName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.PercentCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.SpeciesCover_ID"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.CountID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Alive?"
        dbLong "AggregateType" ="-1"
    End
End
