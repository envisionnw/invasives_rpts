dbMemo "SQL" ="SELECT tc.Unit_Code, tc.Visit_Year, tc.Route, tc.PlantCode, tc.IsDead, MIN(tc.To"
    "talCover) AS TotalCover, MIN(ts.TransectsSampled) AS SampledTransects, MIN(tc.To"
    "talCover / ts.TransectsSampled) AS AverageCover\015\012FROM Route_TotalCover AS "
    "tc INNER JOIN Route_TransectsSampled AS ts ON (ts.Unit_Code = tc.Unit_Code) AND "
    "(ts.Visit_Year = tc.Visit_Year) AND (ts.Route = tc.Route)\015\012GROUP BY tc.Uni"
    "t_Code, tc.Visit_Year, tc.Route, tc.PlantCode, tc.IsDead\015\012ORDER BY tc.Unit"
    "_Code, tc.Visit_Year, tc.Route, tc.PlantCode, tc.IsDead;\015\012"
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
dbMemo "Filter" ="((([Route_AverageCover].[Unit_Code]=\"GOSP\"))) AND ([Route_AverageCover].[Visit"
    "_Year]=2016)"
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
        dbText "Name" ="AverageCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3ce6190874f1bd48b17b01302318e243
        End
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SampledTransects"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
