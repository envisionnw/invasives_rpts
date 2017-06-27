dbMemo "SQL" ="SELECT * INTO temp_Route_Transect_AverageCover\015\012FROM (SELECT\015\012l.Unit"
    "_Code, l.Visit_Year, l.Route, l.Area, l.Transect, \015\012MIN(IIF(l.E_Coord IS N"
    "ULL, ac.E_Coord, l.E_Coord)) AS E_Coord, \015\012MIN(IIF(l.N_Coord IS NULL, ac.N"
    "_Coord, l.N_Coord))  AS N_Coord, \015\012l.PlantCode, l.IsDead, \015\012MIN(ac.T"
    "otalCover) AS TotalCover, \015\012MIN(ac.QuadratsSampled) AS QuadratsSampled, \015"
    "\012MIN(ac.TransectsSampled) AS TransectsSampled, \015\012MIN(IIF(ac.TransectAve"
    "rageCover IS NULL, 0, ac.TransectAverageCover)) AS TransectAverageCover\015\012F"
    "ROM Route_Transect_Species_List l\015\012LEFT JOIN Transect_AverageCover ac\015\012"
    "ON ac.Unit_Code = l.Unit_Code\015\012AND ac.Visit_Year = l.Visit_Year\015\012AND"
    " ac.Route = l.Route\015\012AND ac.Transect = l.Transect\015\012AND ac.PlantCode "
    "= l.PlantCode\015\012AND ac.IsDead = l.IsDead\015\012GROUP BY l.Unit_Code, l.Vis"
    "it_Year, l.Route, l.Area, l.Transect, l.PlantCode, l.IsDead\015\012ORDER BY l.Un"
    "it_Code, l.Visit_Year, l.Route, l.Area, l.Transect, l.PlantCode, l.IsDead)  AS ["
    "%$##@_Alias];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x30539356f705324485708b1ff1baae3e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
