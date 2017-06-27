dbMemo "SQL" ="SELECT DISTINCT ta.Unit_Code, ta.Visit_Year, ta.Route, ta.Transect, ta.Area, ta."
    "E_Coord, ta.N_Coord, ta.PlantCode, sc.Species, sc.Master_Common_Name, ta.IsDead,"
    " ta.TransectsSampled, ta.TotalCover, ta.TransectAverageCover AS TransectAverageC"
    "over, ra.RouteAverageCover AS RouteAverageCover, (ra.RouteAverageCover - ta.Tran"
    "sectAverageCover) AS Deviation, (ra.RouteAverageCover - ta.TransectAverageCover)"
    "^2 AS DeviationSquared\015\012FROM (temp_Route_Transect_AverageCover AS ta INNER"
    " JOIN Route_AverageCover AS ra ON (ra.IsDead = ta.IsDead) AND (ra.PlantCode = ta"
    ".PlantCode) AND (ra.Route = ta.Route) AND (ra.Visit_Year = ta.Visit_Year) AND (r"
    "a.Unit_Code = ta.Unit_Code)) LEFT JOIN Transect_Select_SpeciesCover AS sc ON (sc"
    ".IsDead = ta.IsDead) AND (sc.PlantCode = ta.PlantCode) AND (sc.Transect = ta.Tra"
    "nsect) AND (sc.Route = ta.Route) AND (sc.Visit_Year = ta.Visit_Year) AND (sc.Uni"
    "t_Code = ta.Unit_Code)\015\012ORDER BY ta.Unit_Code, ta.Visit_Year, ta.Route, ta"
    ".Transect, sc.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf7ba8fa327314446a47427377e93bcdf
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ta.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ta.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ta.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ta.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ta.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ta.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ta.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ta.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ta.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ta.TransectsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ta.TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectAverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deviation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DeviationSquared"
        dbLong "AggregateType" ="-1"
    End
End
