dbMemo "SQL" ="SELECT td.Unit_Code, td.Visit_Year, td.Route, td.Area, td.Transect, MIN(td.E_Coo"
    "rd) AS E_Coord, MIN(td.N_Coord) AS N_Coord, tc.PlantCode, tc.IsDead, MIN(tc.Tota"
    "lCover) AS TotalCover, MIN(qs.SampledQuadrats) AS QuadratsSampled, MIN(ts.Transe"
    "ctsSampled) AS TransectsSampled, MIN(td.AverageCover) AS TransectAverageCover\015"
    "\012FROM ((Route_TransectsSampled AS ts INNER JOIN Transect_Data AS td ON (ts.Ro"
    "ute = td.Route) AND (ts.Visit_Year = td.Visit_Year) AND (ts.Unit_Code = td.Unit_"
    "Code)) INNER JOIN Transect_Select_TotalCover AS tc ON (tc.Transect = td.Transect"
    ") AND (tc.Route = td.Route) AND (tc.Visit_Year = td.Visit_Year) AND (tc.Unit_Cod"
    "e = td.Unit_Code)) INNER JOIN Transect_Select_QuadratsSampled AS qs ON (qs.Route"
    " = tc.Route) AND (qs.Visit_Year = tc.Visit_Year) AND (qs.Unit_Code = tc.Unit_Cod"
    "e)\015\012WHERE tc.PlantCode IS NOT NULL\015\012GROUP BY td.Unit_Code, td.Visit_"
    "Year, td.Route, td.Area, td.Transect, tc.PlantCode, tc.IsDead\015\012ORDER BY td"
    ".Unit_Code, td.Visit_Year, td.Route, td.Area, td.Transect, tc.PlantCode, tc.IsDe"
    "ad;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5c57b66d453ddf448d885461a3f21867
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="((([Transect_AverageCover].[Unit_Code]=\"CARE\"))) AND ([Transect_AverageCover]."
    "[Visit_Year]=2015)"
Begin
    Begin
        dbText "Name" ="tc.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x71b91873f6ee17488a4aa0bb809f5bab
        End
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3edd0ac2229ece4a822ca1fe239dfd7d
        End
    End
    Begin
        dbText "Name" ="td.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsSampled"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TransectAverageCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="QuadratsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N_Coord"
        dbLong "AggregateType" ="-1"
    End
End
