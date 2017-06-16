dbMemo "SQL" ="SELECT DISTINCT Unit_Code, Visit_Year\015\012FROM Transect_Select_LIMITED_ESP_Sp"
    "eciesCover_Species\015\012ORDER BY Unit_Code, Visit_Year;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9065182354331d4691a1be3edc76a606
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbc1543836ed0904c9da2423df12c5b84
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x52f6571f28ff80488ce51398017258de
        End
    End
End
