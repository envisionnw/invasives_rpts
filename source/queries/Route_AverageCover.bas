dbMemo "SQL" ="SELECT tc.Unit_Code, tc.Visit_Year, tc.Route, tc.PlantCode, tc.IsDead, MIN(tc.To"
    "talTransectAverageCover) AS TotalTransectAverageCover, MIN(ts.TransectsSampled) "
    "AS SampledTransects, MIN(td.TransectsDetected) AS TransectsDetected, MIN(tc.Tota"
    "lTransectAverageCover / ts.TransectsSampled) AS RouteAverageCover\015\012FROM (R"
    "oute_TotalAverageCover AS tc INNER JOIN Route_TransectsSampled AS ts ON (ts.Unit"
    "_Code = tc.Unit_Code) AND (ts.Visit_Year = tc.Visit_Year) AND (ts.Route = tc.Rou"
    "te)) INNER JOIN Route_TransectsDetected AS td ON (td.Route = tc.Route) AND (td.V"
    "isit_Year = tc.Visit_Year) AND (td.Unit_Code = tc.Unit_Code) AND (td.PlantCode ="
    " tc.PlantCode) AND (td.IsDead = tc.IsDead)\015\012GROUP BY tc.Unit_Code, tc.Visi"
    "t_Year, tc.Route, tc.PlantCode, tc.IsDead\015\012ORDER BY tc.Unit_Code, tc.Visit"
    "_Year, tc.Route, tc.PlantCode, tc.IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbcb7a901773e50489c69bda28267bc8c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="((([Route_AverageCover].[Unit_Code]=\"CARE\"))) AND ([Route_AverageCover].[Visit"
    "_Year]=2015)"
Begin
    Begin
        dbText "Name" ="tc.PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe42e3f12ba89454a91d68089cef02ed8
        End
    End
    Begin
        dbText "Name" ="tc.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x622c6ca7b33eeb45bcacb2b059e940a9
        End
    End
    Begin
        dbText "Name" ="tc.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdc2d838f975ba74f82b382e55938247c
        End
    End
    Begin
        dbText "Name" ="tc.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x58d74b88edbd374982bb23c76609b116
        End
    End
    Begin
        dbText "Name" ="tc.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe0766d6bc68ceb48be56a8a0debdab0a
        End
    End
    Begin
        dbText "Name" ="SampledTransects"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalTransectAverageCover"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TransectsDetected"
        dbLong "AggregateType" ="-1"
    End
End
