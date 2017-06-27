dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, ts.Area, ts."
    "E_Coord, ts.N_Coord, ts.Species, ts.Master_Common_Name, ts.IsDead, st.TransectsS"
    "ampled, tc.TotalCover, (tc.TotalCover/st.TransectsSampled) AS AverageCover, ac.A"
    "verageCover AS RouteAverageCover, (ac.AverageCover - (tc.TotalCover/st.Transects"
    "Sampled)) AS Deviation, (ac.AverageCover - (tc.TotalCover/st.TransectsSampled))^"
    "2 AS DeviationSquared\015\012FROM ((Transect_Select_SpeciesCover AS ts INNER JOI"
    "N Route_TransectsSampled AS st ON (st.Route = ts.Route) AND (st.Visit_Year = ts."
    "Visit_Year) AND (st.Unit_Code = ts.Unit_Code)) INNER JOIN Transect_Select_TotalC"
    "over AS tc ON (tc.IsDead = ts.IsDead) AND (tc.PlantCode = ts.PlantCode) AND (tc."
    "N_Coord = ts.N_Coord) AND (tc.E_Coord = ts.E_Coord) AND (tc.Transect = ts.Transe"
    "ct) AND (tc.Route = ts.Route) AND (tc.Visit_Year = ts.Visit_Year) AND (tc.Unit_C"
    "ode = ts.Unit_Code)) LEFT JOIN Route_AverageCover AS ac ON (ac.IsDead = ts.IsDea"
    "d) AND (ac.PlantCode = ts.PlantCode) AND (ac.Route = ts.Route) AND (ac.Visit_Yea"
    "r = ts.Visit_Year) AND (ac.Unit_Code = ts.Unit_Code)\015\012ORDER BY ts.Unit_Cod"
    "e, ts.Visit_Year, ts.Route, ts.Transect, ts.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xac45e41519f41a45896ce321b8785ff6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x13c793e231197f4698af6de29ce0c9b6
        End
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0c9142bf7d06e74eb2eeef3d75ca49ae
        End
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x38467235cae07e44848db8205451b39d
        End
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xda68d87b8a4e154cb6742781e68e9e1d
        End
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcf2790103c416e4199e304cd476bde2a
        End
    End
    Begin
        dbText "Name" ="ts.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0a481bce4df121418937bf72beaf7069
        End
    End
    Begin
        dbText "Name" ="ts.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1fb531c9c75bbf44a91bb8070b9c1c43
        End
    End
    Begin
        dbText "Name" ="ts.Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2b9dd635d263cb44930c3a18d63145ac
        End
    End
    Begin
        dbText "Name" ="ts.Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfdfa8f1ba34afb44a8521e21b6caab4f
        End
    End
    Begin
        dbText "Name" ="tc.TotalCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdeb49e3f0252fd48af792d784d5d63c9
        End
    End
    Begin
        dbText "Name" ="AverageCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x46f7ad1cbb9a6240808ac19c1ddd6559
        End
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x73b3d887e294674b8df59fb06ec20a43
        End
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Deviation"
        dbInteger "ColumnWidth" ="2535"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbd05e19ae8380b4382d14f71da91ca38
        End
    End
    Begin
        dbText "Name" ="DeviationSquared"
        dbInteger "ColumnWidth" ="2685"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc565c62851ddba49889896f4e07908ce
        End
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="st.TransectsSampled"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
