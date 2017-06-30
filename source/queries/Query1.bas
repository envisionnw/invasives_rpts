dbMemo "SQL" ="SELECT DISTINCT ta.Unit_Code, ta.Visit_Year, ta.Route, ta.Transect, ta.Area, ta."
    "E_Coord, ta.N_Coord, ta.PlantCode, ta.IsDead, ta.TransectsSampled, ta.TotalCover"
    ", ta.TransectAverageCover AS TransectAverageCover, ra.RouteAverageCover AS Route"
    "AverageCover, (ra.RouteAverageCover - ta.TransectAverageCover) AS Deviation, (ra"
    ".RouteAverageCover - ta.TransectAverageCover)^2 AS DeviationSquared\015\012FROM "
    "temp_Route_Transect_AverageCover AS ta INNER JOIN Route_AverageCover AS ra ON (r"
    "a.IsDead = ta.IsDead) AND (ra.PlantCode = ta.PlantCode) AND (ra.Route = ta.Route"
    ") AND (ra.Visit_Year = ta.Visit_Year) AND (ra.Unit_Code = ta.Unit_Code)\015\012O"
    "RDER BY ta.Unit_Code, ta.Visit_Year, ta.Route, ta.Transect, sc.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x603a635aff32c64eaf0a8535a62cea9d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
