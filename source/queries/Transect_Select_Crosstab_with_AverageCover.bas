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
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
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
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
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
    Begin
        dbText "Name" ="Unit_Code"
        dbBinary "GUID" = Begin
            0xf24bffc7e880c74ba416926bc10f583f
        End
    End
    Begin
        dbText "Name" ="Plot_ID"
        dbBinary "GUID" = Begin
            0x970530537afcf94a8474802ba9962dbc
        End
    End
    Begin
        dbText "Name" ="Transect"
        dbBinary "GUID" = Begin
            0x409660c17eb78f418871e23aa84c4106
        End
    End
    Begin
        dbText "Name" ="Area"
        dbBinary "GUID" = Begin
            0x1245e51a455b274791328042b10a11a4
        End
    End
    Begin
        dbText "Name" ="E_Coord"
        dbBinary "GUID" = Begin
            0xb9cd4616f86f2f46881d73d999f818a5
        End
    End
    Begin
        dbText "Name" ="N_Coord"
        dbBinary "GUID" = Begin
            0x942da6a056736b458aa99ab0af908989
        End
    End
    Begin
        dbText "Name" ="Master_Common_Name"
        dbBinary "GUID" = Begin
            0xa1f604e481a9ca479e2c25ffd22ddacd
        End
    End
    Begin
        dbText "Name" ="IsDead"
        dbBinary "GUID" = Begin
            0x0eec529d0b57c54f99d131d68cd1f753
        End
    End
    Begin
        dbText "Name" ="Q1_0_5m"
        dbBinary "GUID" = Begin
            0x0d090629cefc194ab49076cc4ac5aa44
        End
    End
    Begin
        dbText "Name" ="Q2_4_5m"
        dbBinary "GUID" = Begin
            0x571dc7f2e467674cb9b3c01b6d8514a8
        End
    End
    Begin
        dbText "Name" ="Q3_9_5m"
        dbBinary "GUID" = Begin
            0x7d118faa2cf2e144a8f5b4579f844616
        End
    End
    Begin
        dbText "Name" ="Q1_3m"
        dbBinary "GUID" = Begin
            0xc20ddb606b32df4b955d4637395b3cd9
        End
    End
    Begin
        dbText "Name" ="Q2_8m"
        dbBinary "GUID" = Begin
            0x7662d61691e3364986724a595352c130
        End
    End
    Begin
        dbText "Name" ="Q3_13m"
        dbBinary "GUID" = Begin
            0x37cb6f51ab872741b3b694bd96818956
        End
    End
    Begin
        dbText "Name" ="Q1"
        dbBinary "GUID" = Begin
            0x0d1958f77931a14ebc6d770402f47493
        End
    End
    Begin
        dbText "Name" ="Q2"
        dbBinary "GUID" = Begin
            0xda0894af479cda409b1121c695f06abd
        End
    End
    Begin
        dbText "Name" ="Q3"
        dbBinary "GUID" = Begin
            0x66fdc5ca4e8d414ab281295ea463450e
        End
    End
    Begin
        dbText "Name" ="QuadratsSampled"
        dbBinary "GUID" = Begin
            0x95b3e5344768a44d872ec66a1c14d127
        End
    End
    Begin
        dbText "Name" ="TotalCover"
        dbBinary "GUID" = Begin
            0xc5c5d8d7b8d23043ac3cd41ecbed4b59
        End
    End
    Begin
        dbText "Name" ="AverageCover"
        dbBinary "GUID" = Begin
            0x472d483f1c22fd4481662e8aeb3d2ad6
        End
    End
End
