dbMemo "SQL" ="TRANSFORM Min(ts.PercentCover) AS MinOfPercentCover\015\012SELECT ID, ts.Unit_Co"
    "de AS Unit_Code, ts.Plot_ID AS Plot_ID, ts.Transect AS Transect, ts.Area AS Area"
    ", ts.E_Coord AS E_Coord, ts.N_Coord AS N_Coord, ts.Species AS Species, ts.Master"
    "_Common_Name AS Master_Common_Name, ts.IsDead\015\012FROM Transect_Select AS ts\015"
    "\012GROUP BY ID, ts.Unit_Code, ts.Plot_ID, ts.Transect, ts.Area, ts.Species, ts."
    "Master_Common_Name, ts.E_Coord, ts.N_Coord, ts.IsDead\015\012PIVOT ts.ColName;\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x968adc8ce77e7b4da3a18726c2dee921
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7c3c97b086def542bd8582729504a676
        End
    End
    Begin
        dbText "Name" ="Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4241fe908313cb4c9bbb5ea44f4d733a
        End
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x019eb602d70bf44a9da16c8b72a1df05
        End
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1_0_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q1_3m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2_4_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q2_8m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3_13m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q3_9_5m"
        dbLong "AggregateType" ="3"
    End
    Begin
        dbText "Name" ="ts.ID"
        dbInteger "ColumnWidth" ="8115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="5835"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="MinOfPercentCover"
        dbBinary "GUID" = Begin
            0x3e9fa129860bb04c9f17e813789f71b8
        End
    End
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x527b00fe9c80814999da049d0d73a110
        End
    End
    Begin
        dbText "Name" ="Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd7bca6e027edd242a69e96c9a9a3aef2
        End
    End
    Begin
        dbText "Name" ="E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xee0c189afd51c24eb7da61692ca4960c
        End
    End
    Begin
        dbText "Name" ="N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb80c812e00b6864386750ce0237cc9bc
        End
    End
    Begin
        dbText "Name" ="Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb1a9cca1b0f3fd4b8a6b0671f0bf431f
        End
    End
End
