dbMemo "SQL" ="SELECT MIN(d.Unit_Code) AS Unit_Code, MIN(d.Visit_Year) AS Visit_Year, MIN(d.Rou"
    "te) AS Route, MIN(d.Species) AS Species, MIN(d.Master_Common_Name) AS Master_Com"
    "mon_Name, MIN(d.IsDead) AS IsDead, MIN(d.TransectsSampled) AS TransectsSampled, "
    "MIN(d.TotalCover) AS TotalCover, MIN(d.TransectAverageCover) AS TransectAverageC"
    "over, MIN(d.RouteAverageCover) AS RouteAverageCover, MIN(d.TotalDevSquared) AS T"
    "otalDevSquared, MIN(d.StdDeviation) AS StdDeviation, MIN(IIF(d.TransectsSampled "
    "= 0, NULL, IIF(ISNULL(d.StdDeviation) = False,d.StdDeviation/SQR(d.TransectsSamp"
    "led),NULL  ))) AS StdError\015\012FROM Route_StdDeviation AS d\015\012GROUP BY d"
    ".Unit_Code, d.Visit_Year, d.Route, d.Species\015\012ORDER BY d.Unit_Code, d.Visi"
    "t_Year, d.Route, d.Species;\015\012"
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
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x294efa2fd3d0e447a2f14c9374e96a54
        End
        dbInteger "ColumnWidth" ="870"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaffd388928903a419caa12e319af5b7b
        End
        dbInteger "ColumnWidth" ="825"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeb25c7ab34934147b3dfc037e40f7aa9
        End
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8bd8c23af1e62d48bcf4016d0bb8dc27
        End
    End
    Begin
        dbText "Name" ="IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x508c8472571293418855f32d869d7ece
        End
        dbInteger "ColumnWidth" ="615"
        dbBoolean "ColumnHidden" ="0"
    End
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
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TransectAverageCover"
        dbLong "AggregateType" ="-1"
    End
End
