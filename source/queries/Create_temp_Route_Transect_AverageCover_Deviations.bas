dbMemo "SQL" ="SELECT * INTO temp_Route_Transect_AverageCover_Deviations\015\012FROM (SELECT DI"
    "STINCT ta.Unit_Code, ta.Visit_Year, ta.Route, ta.Transect, ta.Area, ta.E_Coord, "
    "ta.N_Coord, \015\012ta.PlantCode, sc.Species, sc.Master_Common_Name, ta.IsDead, "
    "\015\012ta.TransectsSampled, ta.TotalCover, \015\012ta.TransectAverageCover AS T"
    "ransectAverageCover, \015\012ra.RouteAverageCover AS RouteAverageCover, \015\012"
    "(ra.RouteAverageCover - ta.TransectAverageCover) AS Deviation, \015\012(ra.Route"
    "AverageCover - ta.TransectAverageCover)^2 AS DeviationSquared\015\012FROM ((temp"
    "_Route_Transect_AverageCover ta\015\012INNER JOIN Route_AverageCover ra\015\012O"
    "N (ra.Unit_Code = ta.Unit_Code) \015\012AND (ra.Visit_Year = ta.Visit_Year) \015"
    "\012AND (ra.Route = ta.Route) \015\012AND (ra.PlantCode = ta.PlantCode) \015\012"
    "AND (ra.IsDead = ta.IsDead)) \015\012LEFT JOIN Transect_Select_SpeciesCover sc\015"
    "\012ON (sc.Unit_Code = ta.Unit_Code) \015\012AND (sc.Visit_Year = ta.Visit_Year)"
    " \015\012AND (sc.Route = ta.Route) \015\012AND (sc.Transect = ta.Transect) \015\012"
    "AND (sc.PlantCode = ta.PlantCode) \015\012AND (sc.IsDead = ta.IsDead))\015\012OR"
    "DER BY ta.Unit_Code, ta.Visit_Year, ta.Route, ta.Transect, sc.Species\015\012)  "
    "AS [%$##@_Alias];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa6140c352803c14a8efcef1889de493f
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
