dbMemo "SQL" ="SELECT tc.Unit_Code, tc.Visit_Year, tc.Route, tc.PlantCode, tc.IsDead, MIN(tc.To"
    "talCover) AS TotalCover, MIN(qs.SampledQuadrats) AS SampledQuadrats, MIN(tc.Tota"
    "lCover / qs.SampledQuadrats) AS AverageCover\015\012FROM Route_TotalCover AS tc "
    "INNER JOIN Route_QuadratsSampled AS qs ON (qs.Route = tc.Route) AND (qs.Visit_Ye"
    "ar = tc.Visit_Year) AND (qs.Unit_Code = tc.Unit_Code)\015\012GROUP BY tc.Unit_Co"
    "de, tc.Visit_Year, tc.Route, tc.PlantCode, tc.IsDead\015\012ORDER BY tc.Unit_Cod"
    "e, tc.Visit_Year, tc.Route, tc.PlantCode, tc.IsDead;\015\012"
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
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x45afc71e9d36d74ea2b4c9ce42fbd509
        End
    End
    Begin
        dbText "Name" ="Expr1006"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x614dfe62720ce9448b63c1c4ee9da62b
        End
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SampledQuadrats"
        dbLong "AggregateType" ="-1"
    End
End
