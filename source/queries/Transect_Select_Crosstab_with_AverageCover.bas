dbMemo "SQL" ="SELECT DISTINCT xt.Unit_Code, xt.Visit_Year, xt.Plot_ID, xt.Transect, xt.Area, x"
    "t.E_Coord, xt.N_Coord, xt.Species, xt.Master_Common_Name, xt.IsDead, xt.Q1_0_5m,"
    " xt.Q2_4_5m, xt.Q3_9_5m, xt.Q1_3m, xt.Q2_8m, xt.Q3_13m, xt.Q1, xt.Q2, xt.Q3, tsc"
    ".QuadratsSampled, tsc.TotalCover, tsc.AverageCover\015\012FROM Transect_Select_C"
    "rosstab_with_ID AS xt INNER JOIN Transect_Select_Count AS tsc ON tsc.ID = xt.ID\015"
    "\012ORDER BY xt.Unit_Code, xt.Visit_Year, xt.Plot_ID, xt.Transect, xt.Area, xt.S"
    "pecies, xt.IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x054728e21c3eee488d8651fc1eb4dd57
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x85da5fee3f31244785bbf47115e6be6f
        End
    End
    Begin
        dbText "Name" ="l.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x12eaca863c3dd64f98a7ce20becd9f0b
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PercentCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6b7110ef092d7841a6c97335fd7481bd
        End
    End
    Begin
        dbText "Name" ="esp.Position_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="esp.ColName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q1_0_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q2_4_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q3_9_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q1_3m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.AverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Visit_Year"
        dbInteger "ColumnWidth" ="1290"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.TotalCover"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="xt.Q2_8m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q3_13m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="xt.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.QuadratsSampled"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
    End
End
