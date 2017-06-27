dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Area, ts.Transect, ts."
    "Species, MIN(ts.Master_Common_Name) AS Master_Common_Name, MIN(ts.PlantCode) AS "
    "PlantCode, ts.IsDead, SUM(ts.PercentCover) AS PercentCover\015\012FROM Transect_"
    "Select_SpeciesCover AS ts\015\012GROUP BY ts.Unit_Code, ts.Visit_Year, ts.Route,"
    " ts.Area, ts.Transect, ts.Species, ts.IsDead\015\012ORDER BY ts.Unit_Code, ts.Vi"
    "sit_Year, ts.Route, ts.Area, ts.Transect, ts.Species, ts.IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0013dcce293bf74198457d916c795c7d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="((([Transect_Select_SpeciesCover_Aggregate].[Unit_Code]=\"CARE\"))) AND ([Transe"
    "ct_Select_SpeciesCover_Aggregate].[Visit_Year]=2015)"
Begin
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa9c48d2c99e4c34988fa264dd19a0bb6
        End
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd4b64e7121007e448946e8871f957a15
        End
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9acb42f437f2d544a4dde8829abb7b16
        End
        dbInteger "ColumnWidth" ="2445"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeb194451ac4e6e4289fef03073ba999a
        End
        dbInteger "ColumnWidth" ="705"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ts.Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8410fac1a646a246b0067e6003bf636f
        End
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x369737e1f0797d45887e4502e417e369
        End
    End
    Begin
        dbText "Name" ="Quadrat"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbcb1256c6b93ba48ae43b50e745b9c13
        End
    End
    Begin
        dbText "Name" ="IsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3a6f8b8a9b5115489ea9e5abcc9258c7
        End
    End
    Begin
        dbText "Name" ="NoExotics"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0c5703659e717f4ebad8260e9072b758
        End
    End
    Begin
        dbText "Name" ="ColName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbbbe21673c89254eb38130541cd55ef1
        End
    End
    Begin
        dbText "Name" ="Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x97d4b86999d17e40ae8a2b07eb735733
        End
    End
    Begin
        dbText "Name" ="PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeefbea009a64db4c8faab18072d0eb3c
        End
        dbInteger "ColumnWidth" ="750"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="PercentCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x56626422bc2eb24994704675e95d4215
        End
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
    End
End
