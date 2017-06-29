dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Area, ts.Transect, ts."
    "E_Coord, ts.N_Coord, ts.Species, ts.Master_Common_Name, IIF(ts.IsDead = 1, \"N\""
    ", \"Y\") AS [Alive?], ts.PlantCode, ts.IsDead, tc.TotalCover, qs.SampledQuadrats"
    ", rts.TransectsSampled, (tc.TotalCover/qs.SampledQuadrats) AS TransectAverageCov"
    "er\015\012FROM ((Transect_Select_SpeciesCover AS ts INNER JOIN Transect_Select_T"
    "otalCover AS tc ON (tc.IsDead = ts.IsDead) AND (tc.PlantCode = ts.PlantCode) AND"
    " (tc.Transect = ts.Transect) AND (tc.Route = ts.Route) AND (tc.Visit_Year = ts.V"
    "isit_Year) AND (tc.Unit_Code = ts.Unit_Code)) INNER JOIN Transect_Select_Quadrat"
    "sSampled AS qs ON (qs.Transect = ts.Transect) AND (qs.Route = ts.Route) AND (qs."
    "Visit_Year = ts.Visit_Year) AND (qs.Unit_Code = ts.Unit_Code)) INNER JOIN Route_"
    "TransectsSampled AS rts ON (rts.Route = ts.Route) AND (rts.Visit_Year = ts.Visit"
    "_Year) AND (rts.Unit_Code = ts.Unit_Code)\015\012ORDER BY ts.Unit_Code, ts.Visit"
    "_Year, ts.Route, ts.Transect, ts.Species, ts.IsDead;\015\012"
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
        dbText "Name" ="tc.TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qs.SampledQuadrats"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="rts.TransectsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.PlantCode"
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
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectAverageCover"
        dbLong "AggregateType" ="-1"
    End
End
