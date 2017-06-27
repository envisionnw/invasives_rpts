dbMemo "SQL" ="SELECT MIN(d.Unit_Code) AS Unit_Code, MIN(d.Visit_Year) AS Visit_Year, MIN(d.Rou"
    "te) AS Route, MIN(d.Species) AS Species, MIN(d.Master_Common_Name) AS Master_Com"
    "mon_Name, MIN(d.IsDead) AS IsDead, MIN(d.TransectsSampled) AS TransectsSampled, "
    "MIN(d.TotalCover) AS TotalCover, MIN(d.TransectAverageCover) AS TransectAverageC"
    "over, MIN(d.RouteAverageCover) AS RouteAverageCover, SUM(d.DeviationSquared) AS "
    "TotalDevSquared\015\012FROM Route_AverageCover_Deviations AS d\015\012WHERE d.Sp"
    "ecies IS NOT NULL\015\012GROUP BY d.Unit_Code, d.Visit_Year, d.Route, d.Species\015"
    "\012ORDER BY d.Unit_Code, d.Visit_Year, d.Route, d.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2b931526cfe1714982b3145caff62da4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x30cd11173b9b6744b7c40ad8a055c911
        End
    End
    Begin
        dbText "Name" ="IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0e8c64b8f7aeb04e8ee5057729816f1d
        End
        dbInteger "ColumnWidth" ="465"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TransectsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1c87d6ba5e839d49a5be3884317226d4
        End
        dbInteger "ColumnWidth" ="615"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x90c2ad9221c6af499f0164c2763327bc
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbae8b6f521043c459cb5bd5e3d0ae2e8
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7044779357973b488ea7e20ffbdfc470
        End
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7e37d10d9df6d34595dc0efa6e534550
        End
    End
    Begin
        dbText "Name" ="TotalCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x367bd53dccd4fb40a154c65e9f4e0892
        End
    End
    Begin
        dbText "Name" ="TransectAverageCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x33b3e8fb02c59f4ca396b4e803fb8966
        End
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8392445dca1cef4aacbbdf19e41b47f3
        End
    End
    Begin
        dbText "Name" ="TotalDevSquared"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3ef55be73d4c1c4aaa9a00492dd8455c
        End
    End
End
