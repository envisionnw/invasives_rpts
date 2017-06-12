dbMemo "SQL" ="SELECT tsca.Unit_Code, tsca.Visit_Year, tsca.Route, tsca.Transect, tsca.Area, ts"
    "ca.E_Coord, tsca.N_Coord, tsca.Species, tsca.Master_Common_Name, tsca.IsDead, ts"
    "ca.QuadratsSampled, tsca.TotalCover, tsca.AverageCover, Dev_Q1_0_5m^2 AS DevQ1_0"
    "_5m, Dev_Q2_4_5m^2 AS DevQ2_4_5m, Dev_Q3_9_5m^2 AS DevQ3_9_5m, Dev_Q1_3m^2 AS De"
    "vQ1_3m, Dev_Q2_8m^2 AS DevQ2_8m, Dev_Q3_13m^2 AS DevQ3_13m, Dev_Q1^2 AS DevQ1, D"
    "ev_Q2^2 AS DevQ2, Dev_Q3^2 AS DevQ3\015\012FROM Transect_Select_Crosstab_with_Av"
    "erageCover_Deviations AS tsca\015\012ORDER BY tsca.Unit_Code, tsca.Visit_Year, t"
    "sca.Route, tsca.Transect, tsca.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x207c37369152cc4da3a72f6ce1014633
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tsca.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.QuadratsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StdDeviation"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deviations"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.AverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DevQ1_0_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9b0590cfe88a7845b7e3d0a5e7883ec3
        End
    End
    Begin
        dbText "Name" ="DevQ2_4_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x03cdb4d0f65e634b9e2900461bb0a4a7
        End
    End
    Begin
        dbText "Name" ="DevQ3_9_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x800c2b6504dad34f9b9febcfcad1a76a
        End
    End
    Begin
        dbText "Name" ="DevQ1_3m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x53f3606e1f589c4f94688b6cee8f2343
        End
    End
    Begin
        dbText "Name" ="DevQ2_8m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x795b39cbb310b64baccdf9ba8ae909be
        End
    End
    Begin
        dbText "Name" ="DevQ3_13m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x58d3d757bbbc4249ac1d47fbe464e6e5
        End
    End
    Begin
        dbText "Name" ="DevQ1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4fbf532cbb6e314ba1132f52e83d0e47
        End
    End
    Begin
        dbText "Name" ="DevQ2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe0ab3bd92881a54da531fcd14bf6b8f6
        End
    End
    Begin
        dbText "Name" ="DevQ3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xed74b0f0969a2f449fd7e87f9eca46ac
        End
    End
    Begin
        dbText "Name" ="tsca.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1019"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1020"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1021"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StdError"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
