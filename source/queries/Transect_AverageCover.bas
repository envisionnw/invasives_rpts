dbMemo "SQL" ="SELECT tc.Unit_Code, tc.Visit_Year, tc.Route, tc.Transect, tc.PlantCode, tc.IsDe"
    "ad, MIN(tc.TotalCover) AS TotalCover, MIN(ts.SampledQuadrats) AS SampledQuadrats"
    ", MIN(tc.TotalCover / ts.SampledQuadrats) AS AverageCover\015\012FROM Transect_S"
    "elect_TotalCover AS tc INNER JOIN Transect_Select_QuadratsSampled AS ts ON (ts.U"
    "nit_Code = tc.Unit_Code) AND (ts.Visit_Year = tc.Visit_Year) AND (ts.Route = tc."
    "Route)\015\012WHERE tc.PlantCode IS NOT NULL\015\012GROUP BY tc.Unit_Code, tc.Vi"
    "sit_Year, tc.Route, tc.Transect, tc.PlantCode, tc.IsDead\015\012ORDER BY tc.Unit"
    "_Code, tc.Visit_Year, tc.Route, tc.Transect, tc.PlantCode, tc.IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5c57b66d453ddf448d885461a3f21867
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tc.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xedda0aa4ee012e4588106d048d615851
        End
    End
    Begin
        dbText "Name" ="tc.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x039da84ddb9f1342baba9fc6e0ee893b
        End
    End
    Begin
        dbText "Name" ="tc.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdaf5003d32c9bb42add881e7a2cd84f0
        End
    End
    Begin
        dbText "Name" ="tc.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb2407ede89fdf641a33df740c19f984b
        End
    End
    Begin
        dbText "Name" ="tc.PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x20e653e8a15b8840a9f5995ef0b88a14
        End
    End
    Begin
        dbText "Name" ="tc.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x71b91873f6ee17488a4aa0bb809f5bab
        End
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3edd0ac2229ece4a822ca1fe239dfd7d
        End
    End
    Begin
        dbText "Name" ="AverageCover"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2edc6d21e4f8984193e5519f470dad9f
        End
    End
    Begin
        dbText "Name" ="SampledQuadrats"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xebe4580196d1904e8326a3b477433c3a
        End
    End
End
