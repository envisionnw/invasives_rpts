dbMemo "SQL" ="SELECT DISTINCT Unit_Code, Visit_Year, Route, Area, PlantCode, IsDead\015\012FRO"
    "M Transect_AverageCover\015\012ORDER BY Unit_Code, Visit_Year, Route, Area, Plan"
    "tCode;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9ab3e1fb37cd6e46a5da1382c789ead1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd53c9af635d70d4a87eb32a05d212645
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0886ec835a3e3c46991a557c91bb3e83
        End
    End
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbda574700e3b874386905092467aaccb
        End
    End
    Begin
        dbText "Name" ="Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbafd3e5406e3f94a8eadf4b5558e545b
        End
    End
    Begin
        dbText "Name" ="PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3ed858361d8d114dab5df0614f2dfc5b
        End
    End
    Begin
        dbText "Name" ="IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7fc6ee01d99edd4e8fb5fc400687a6a4
        End
    End
End
