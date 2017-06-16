dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect_ID, ts.Transe"
    "ct, ts.Area, ts.E_Coord, ts.N_Coord, ts.Quadrat, ts.IsSampled, ts.NoExotics\015\012"
    "FROM Transect_Select_LIMITED_ESP_SpeciesCover_Species AS ts\015\012ORDER BY ts.U"
    "nit_Code, ts.Visit_Year, ts.Route, ts.Transect_ID, ts.Transect, ts.Area, ts.E_Co"
    "ord, ts.N_Coord, ts.Quadrat;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Transect_Select_Quadrats].[Unit_Code]=\"GOSP\"))) AND ([Transect_Select_Quad"
    "rats].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf27889785ac12041a847645760e82ae8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ts.IsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe00a36160730bd4789006f058742671d
        End
    End
    Begin
        dbText "Name" ="ts.NoExotics"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x70da1a28283eb64e8142bf9808af5ed0
        End
    End
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x52d900ff5c15954faca3be2ec3f3bdd1
        End
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x09d3051a8a068c42b23583f7802cd407
        End
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x145751787fc6764c8dfc08770107bfb4
        End
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe121ed0bc2bc9e44abfd2d0ee5d7241c
        End
    End
    Begin
        dbText "Name" ="ts.Transect_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9ed56e51f2e2ea4ea3ed8d83aad7dc62
        End
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xecdb5f49dd5e7b4780ccaebd087850f7
        End
    End
    Begin
        dbText "Name" ="ts.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2327a755f0e81e4ca9d589d717d8cedd
        End
    End
    Begin
        dbText "Name" ="ts.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x509fc3abd6bd6549b2d7d18750645120
        End
    End
    Begin
        dbText "Name" ="ts.Quadrat"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x423c370c1dba8640a7ff390e84b4f3fb
        End
    End
End
