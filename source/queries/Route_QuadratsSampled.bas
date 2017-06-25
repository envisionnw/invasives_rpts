dbMemo "SQL" ="SELECT ts.Unit_Code, ts.Visit_Year, ts.Route, SUM(ts.IsSampled) AS SampledQuadra"
    "ts\015\012FROM Transect_Select_Quadrat AS ts\015\012GROUP BY ts.Unit_Code, ts.Vi"
    "sit_Year, ts.Route\015\012ORDER BY ts.Unit_Code, ts.Visit_Year, ts.Route;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Route_QuadratsSampled].[Unit_Code]=\"CARE\"))) AND ([Route_QuadratsSampled]."
    "[Visit_Year]=2015)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x871ebfed32ff314fa7206fc2d0bc0612
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x08c798ccbb62a84db4cc3687a820e593
        End
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x45beadaa280f3143ad7ce7a4bda737d3
        End
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1ddb1f2afebbc44c876e2076f4809c21
        End
    End
    Begin
        dbText "Name" ="SampledQuadrats"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7adcd4b98ad30847901b116d8b0e0b0e
        End
    End
End
