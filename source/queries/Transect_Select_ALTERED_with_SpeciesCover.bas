dbMemo "SQL" ="SELECT (l.Plot_ID  & \"_\" & t.Transect) AS ID, l.Unit_Code, Year([Start_Date]) "
    "AS Visit_Year, l.Plot_ID AS Route, t.Transect_ID, t.Transect, l.Area, t.E_Coord,"
    " t.N_Coord, esp.Position_m, esp.ColName, q.IsSampled, q.NoExotics, sc.PlantCode,"
    " sc.IsDead, sc.PercentCover\015\012FROM (((tbl_Locations AS l LEFT JOIN EventSam"
    "plePosition AS esp ON esp.Location_ID = l.Location_ID) LEFT JOIN Transect AS t O"
    "N t.Event_ID = esp.Event_ID) LEFT JOIN Quadrat AS q ON q.Transect_ID = t.Transec"
    "t_ID) LEFT JOIN SpeciesCover AS sc ON sc.Quadrat_ID = q.ID\015\012WHERE esp.Quad"
    "rat = q.Quadrat\015\012ORDER BY l.Plot_ID, t.Transect, q.Quadrat;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Transect_Select_ALTERED_with_SpeciesCover].[Unit_Code]=\"GOSP\"))) AND ([Tra"
    "nsect_Select_ALTERED_with_SpeciesCover].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x42ea3993cfd84a449937bdc111887b91
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x820176cd60e60e4cbb376e9357644306
        End
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x16168c7b636a61458793524ad1d64d11
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf3c0b9e157a2704da9a18cf1aa2f1fab
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x290d50edde6f0e47848baf200ee2fef3
        End
    End
    Begin
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x89cc49600be23b42ac0267ec9d61bc19
        End
    End
    Begin
        dbText "Name" ="t.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfcdfcf5532ba6d48811b03c12a98076e
        End
    End
    Begin
        dbText "Name" ="l.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfefa291811e7fe498f6bcaee7c1aa660
        End
    End
    Begin
        dbText "Name" ="t.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xafeb38e7ca275b408ccfed28ac985f95
        End
    End
    Begin
        dbText "Name" ="t.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x14631c3a3cc70842a7b9e802faf7f1c3
        End
    End
    Begin
        dbText "Name" ="esp.Position_m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb69bc6aebc3ddc4299d74f1c80e31aca
        End
    End
    Begin
        dbText "Name" ="esp.ColName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaeac6bf3583b2449815a0f8963e75809
        End
    End
    Begin
        dbText "Name" ="q.IsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xab61ea3382c3ea44b9904d6b0ab49950
        End
    End
    Begin
        dbText "Name" ="q.NoExotics"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8dd40a1c21204f43ad1221b7184dccff
        End
    End
    Begin
        dbText "Name" ="sc.PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9a0e5c5af8c2f34e9c27b28dd4efa970
        End
    End
    Begin
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x843419c794fe8349a6dd08313388bb47
        End
    End
    Begin
        dbText "Name" ="sc.PercentCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x296d445ca229664eab9679197bd69584
        End
    End
End
