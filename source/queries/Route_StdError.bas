dbMemo "SQL" ="SELECT d.Unit_Code, d.Visit_Year, d.Route, d.Species, MIN(d.Master_Common_Name) "
    "AS Master_Common_Name, d.IsDead, MIN(d.TransectsSampled) AS TransectsSampled, MI"
    "N(d.TransectsDetected) AS TransectsDetected, MIN(d.TotalCover) AS TotalCover, MI"
    "N(d.TransectAverageCover) AS TransectAverageCover, MIN(d.RouteAverageCover) AS R"
    "outeAverageCover, MIN(d.TotalDevSquared) AS TotalDevSquared, MIN(d.StdDeviation)"
    " AS StdDeviation, MIN(IIF(d.TransectsSampled = 0, NULL, IIF(ISNULL(d.StdDeviatio"
    "n) = False,d.StdDeviation/SQR(d.TransectsSampled),NULL  ))) AS StdError\015\012F"
    "ROM Route_StdDeviation AS d\015\012GROUP BY d.Unit_Code, d.Visit_Year, d.Route, "
    "d.Species, d.IsDead\015\012ORDER BY d.Unit_Code, d.Visit_Year, d.Route, d.Specie"
    "s, d.IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x37883cd60961844884fc41038036c30e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf8d89a835716394985419ea9fc3d7a8f
        End
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8752c40e329fe94a95a5323c95e96eac
        End
    End
    Begin
        dbText "Name" ="TotalDevSquared"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x71aa10cf732b6643b442f193458f5dc9
        End
    End
    Begin
        dbText "Name" ="StdDeviation"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6bd662df2dafc74b9cb72fc8f78fcaec
        End
    End
    Begin
        dbText "Name" ="StdError"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xef6cb45e4e04b448a6035ccd7163532a
        End
    End
    Begin
        dbText "Name" ="Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsSampled"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TransectAverageCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="d.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectsDetected"
        dbLong "AggregateType" ="-1"
    End
End
