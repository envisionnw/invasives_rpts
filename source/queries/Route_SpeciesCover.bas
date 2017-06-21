dbMemo "SQL" ="SELECT DISTINCT se.Unit_Code, se.Visit_Year, se.Route, se.Species, se.Master_Com"
    "mon_Name, IIF(se.IsDead = 1,'N','Y') AS [Alive?], se.TransectsSampled, se.TotalC"
    "over, se.RouteAverageCover, se.StdDeviation, se.StdError, rt.RouteTruncated & \""
    " (\" & rt.TransectCount & \") TCount\" AS ColRouteTransects, rt.RouteTruncated &"
    " \" (\" & rt.TransectCount & \") PctCover\" AS ColRouteCover, rt.RouteTruncated "
    "& \" (\" & rt.TransectCount & \") SE\" AS ColRouteStdError\015\012FROM Route_Std"
    "Error AS se LEFT JOIN Route_Transects AS rt ON (rt.Route = se.Route) AND (rt.Vis"
    "it_Year = se.Visit_Year) AND (rt.Unit_Code = se.Unit_Code)\015\012ORDER BY se.Un"
    "it_Code, se.Visit_Year, se.Route, se.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1142c6a053400948aa5c8b0e96ffd2bd
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="se.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc074869c2e9f7d4da49665e4208a2eeb
        End
    End
    Begin
        dbText "Name" ="se.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xff98d0629a61b74ea2f5f34d17c1142a
        End
    End
    Begin
        dbText "Name" ="se.Route"
        dbInteger "ColumnWidth" ="2820"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbddd08b9cf6cce4c9d6a2f864a47fc5a
        End
    End
    Begin
        dbText "Name" ="se.Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6aeb7036a02fc8429e7200c2eac42b0c
        End
    End
    Begin
        dbText "Name" ="se.Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5ea910ae1397014a8ce4a16f9428a4d2
        End
    End
    Begin
        dbText "Name" ="se.TransectsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe778e4c10f8528418ff403dc23fe7848
        End
        dbInteger "ColumnWidth" ="2535"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="se.TotalCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6d2c5abb75ab6244a0b6cbd78c2d0a94
        End
    End
    Begin
        dbText "Name" ="se.StdDeviation"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf45988f8cc13e145803458ed1cbc2fa5
        End
    End
    Begin
        dbText "Name" ="se.StdError"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x984f4515e6c04941914a61d83eb5ca9d
        End
    End
    Begin
        dbText "Name" ="ColRouteTransects"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdfd6b37db1047e46a463e76a1da73db2
        End
    End
    Begin
        dbText "Name" ="ColRouteCover"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x947a63d34baf7940b8cdb3fdd9471afb
        End
    End
    Begin
        dbText "Name" ="ColRouteStdError"
        dbInteger "ColumnWidth" ="2940"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9fd708ee7186794bb4f50b1a1ef3bdd7
        End
    End
    Begin
        dbText "Name" ="se.RouteAverageCover"
        dbLong "AggregateType" ="-1"
    End
End
