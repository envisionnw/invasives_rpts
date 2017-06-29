dbMemo "SQL" ="SELECT * INTO temp_Route_Transect_AverageCover_NEW\015\012FROM (SELECT ac.*, td."
    "TransectsDetected\015\012FROM temp_Update_Table td\015\012INNER JOIN temp_Route_"
    "Transect_AverageCover ac\015\012ON\015\012td.Unit_Code = ac.Unit_Code\015\012AND"
    " td.Visit_Year = ac.Visit_Year\015\012AND td.Route = ac.Route\015\012AND td.Tran"
    "sect = ac.Transect\015\012AND td.PlantCode = ac.PlantCode\015\012AND td.IsDead ="
    " ac.IsDead\015\012)  AS [%$##@_Alias];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6f761986bffb5a48bbf5defcd451647e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ac.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7a56d754b76d6346b2ac4000d6ae0821
        End
    End
    Begin
        dbText "Name" ="ac.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdfed580099a6bd4391b6a98e931ddd3a
        End
    End
    Begin
        dbText "Name" ="ac.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9bdcd50d3d96ca45b8b6a86d27cb638b
        End
    End
    Begin
        dbText "Name" ="ac.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x430f98ff9c99ff41bb07e53e0b1ed874
        End
    End
    Begin
        dbText "Name" ="ac.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc8e50d26cac02f468460469ed1e05776
        End
    End
    Begin
        dbText "Name" ="ac.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa8c6cac688859041843c32ee7d8678da
        End
    End
    Begin
        dbText "Name" ="ac.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x26d4174fa75f7f4091cb7fe7e909d441
        End
    End
    Begin
        dbText "Name" ="ac.PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x10f937608329d24abd1f86124e38bb7f
        End
    End
    Begin
        dbText "Name" ="ac.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x01badc894763d74daf9995a625bff0f2
        End
    End
    Begin
        dbText "Name" ="ac.TotalCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x17dafee9692b43448b6aaf72854af0a4
        End
    End
    Begin
        dbText "Name" ="ac.QuadratsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x68d7e7dfe13d754a9179e65ee88a994b
        End
    End
    Begin
        dbText "Name" ="ac.TransectsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9104ccfd74aa2b43947c1f1fe2785aa3
        End
    End
    Begin
        dbText "Name" ="ac.TransectAverageCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9f7183d28bdb9b4c9f78d46a8c936085
        End
    End
    Begin
        dbText "Name" ="ac.TransectsDetected"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7b28f1c55f525c44bb11e69369c60fbf
        End
    End
    Begin
        dbText "Name" ="td.TransectsDetected"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb1d079da31c69948af1946625412d811
        End
    End
End
