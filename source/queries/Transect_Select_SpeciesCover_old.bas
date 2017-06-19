dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Location_ID, ts.Route, ts.Transe"
    "ct, ts.Area, ts.E_Coord, ts.N_Coord, ts.Quadrat_ID, ts.Quadrat, ts.IsSampled, ts"
    ".NoExotics, ts.Position_m, ts.ColName, ts.Species, ts.Master_Common_Name, ts.Pla"
    "ntCode, ts.IsDead, sq.SampledQuadrats\015\012FROM Transect_Select_LIMITED_ESP_Sp"
    "eciesCover_Species AS ts LEFT JOIN Transect_Select_QuadratsSampled AS sq ON (sq."
    "Unit_Code = ts.Unit_Code) AND (sq.Visit_Year = ts.Visit_Year) AND (sq.Route = ts"
    ".Route) AND (sq.Transect = ts.Transect) AND (sq.Area = ts.Area) AND (sq.E_Coord "
    "= ts.E_Coord) AND (sq.N_Coord = ts.N_Coord)\015\012ORDER BY ts.Route, ts.Transec"
    "t, ts.Quadrat, ts.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xcaa63738fb8fa941a1016ee17d3665ee
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="((([Query4].[Unit_Code]=\"GOSP\"))) AND ([Query4].[Visit_Year]=2016)"
Begin
    Begin
        dbText "Name" ="ts.Location_ID"
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
        dbText "Name" ="ts.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Master_Common_Name"
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
        dbText "Name" ="sq.SampledQuadrats"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2295"
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
