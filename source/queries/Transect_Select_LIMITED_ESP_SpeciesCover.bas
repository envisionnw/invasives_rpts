dbMemo "SQL" ="SELECT DISTINCT ts.ID, ts.Unit_Code, ts.Visit_Year, ts.Location_ID, ts.Route, ts"
    ".Transect_ID, ts.Transect, ts.Area, ts.E_Coord, ts.N_Coord, ts.Quadrat_ID, ts.Qu"
    "adrat, ts.IsSampled, ts.NoExotics, ts.Position_m, ts.ColName, sc.PlantCode, sc.I"
    "sDead, sc.PercentCover\015\012FROM Transect_Select_LIMITED_ESP AS ts LEFT JOIN S"
    "peciesCover AS sc ON sc.Quadrat_ID = ts.Quadrat_ID\015\012ORDER BY ts.Route, ts."
    "Transect, ts.Quadrat;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Query4].[Unit_Code]=\"GOSP\"))) AND ([Query4].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfb26c579cd0d92418a3dc7304d9ca95e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ts.ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1d19bbfba7eb3d418a8e37ae6b9177db
        End
    End
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2b08a3c49ad1ea439f6ff6f79e5106b0
        End
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x04f4b6f878b7ea4fa656fa485e714057
        End
    End
    Begin
        dbText "Name" ="ts.Location_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x67a0c3830e92b24581ced3805d2d2ec8
        End
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb03c8a9ab6e5d944bd1531ffe18177a0
        End
    End
    Begin
        dbText "Name" ="ts.Transect_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd4142b104887524a8c85e7535561aa93
        End
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3a532811eb54ff4086a89987304d9472
        End
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x22c4a2c12bff4c4a8a5caeed80657ef6
        End
    End
    Begin
        dbText "Name" ="ts.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x44e97b755e98b743877c99eeb0020531
        End
    End
    Begin
        dbText "Name" ="ts.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xda36b8b09b062d4a8f0ead599136852b
        End
    End
    Begin
        dbText "Name" ="ts.Quadrat_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfaebb0ce06777f42b68fdf558e85c9d1
        End
    End
    Begin
        dbText "Name" ="ts.Quadrat"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0accfe9dae29d04e96b34a9d29d05b25
        End
    End
    Begin
        dbText "Name" ="ts.IsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x60fe03cf0f0806449909891fb2f8ca58
        End
    End
    Begin
        dbText "Name" ="ts.NoExotics"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe6fb8904f3bef14b8bc2934642eee202
        End
    End
    Begin
        dbText "Name" ="ts.Position_m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfb76623f28b234469d62b98d926ae7b0
        End
    End
    Begin
        dbText "Name" ="ts.ColName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcff77550b86af3428f4af0ebdcda4b85
        End
    End
    Begin
        dbText "Name" ="sc.PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfc7aedf984ccfd48aeab0b4497d78b27
        End
    End
    Begin
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x08140d0825eb7147930f99d763bd4737
        End
    End
    Begin
        dbText "Name" ="sc.PercentCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe3d2a744dca94e4093969532565fd2e6
        End
    End
End
