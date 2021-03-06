﻿dbMemo "SQL" ="SELECT DISTINCT ta.Unit_Code, ta.Visit_Year, ta.Route, ta.Transect, ta.Area, ta."
    "E_Coord, ta.N_Coord, ta.PlantCode, IIf(ta.Unit_Code In (\"CARE\",\"DINO\",\"GOSP"
    "\",\"ZION\"),p.[Utah_Species],\015\012IIf(ta.Unit_Code =\"FOBU\",p.[WY_Species],"
    "p.[Co_Species])) AS Species, p.Master_Common_Name, ta.IsDead, ra.SampledTransect"
    "s AS TransectsSampled, ta.TransectsDetected AS TransectsDetected, ta.TotalCover,"
    " ta.TransectAverageCover AS TransectAverageCover, ra.RouteAverageCover AS RouteA"
    "verageCover, (ra.RouteAverageCover - ta.TransectAverageCover) AS Deviation, (ra."
    "RouteAverageCover - ta.TransectAverageCover)^2 AS DeviationSquared\015\012FROM ("
    "(temp_Route_Transect_AverageCover AS ta INNER JOIN Route_AverageCover AS ra ON ("
    "ra.IsDead = ta.IsDead) AND (ra.PlantCode = ta.PlantCode) AND (ra.Route = ta.Rout"
    "e) AND (ra.Visit_Year = ta.Visit_Year) AND (ra.Unit_Code = ta.Unit_Code)) LEFT J"
    "OIN Transect_Select_SpeciesCover AS sc ON (sc.IsDead = ta.IsDead) AND (sc.PlantC"
    "ode = ta.PlantCode) AND (sc.Transect = ta.Transect) AND (sc.Route = ta.Route) AN"
    "D (sc.Visit_Year = ta.Visit_Year) AND (sc.Unit_Code = ta.Unit_Code)) LEFT JOIN t"
    "lu_NCPN_Plants AS p ON p.Master_PLANT_Code = ta.PlantCode\015\012ORDER BY ta.Uni"
    "t_Code, ta.Visit_Year, ta.Route, ta.Transect, IIf(ta.Unit_Code In (\"CARE\",\"DI"
    "NO\",\"GOSP\",\"ZION\"),p.[Utah_Species],\015\012IIf(ta.Unit_Code =\"FOBU\",p.[W"
    "Y_Species],p.[Co_Species]));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8743b21d10e1ae4581babd74e34575cf
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
        dbText "Name" ="ta.IsDead"
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
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deviation"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DeviationSquared"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsSampled"
        dbLong "AggregateType" ="-1"
    End
End
