dbMemo "SQL" ="SELECT DISTINCT lu.ParkCode, lu.ParkName\015\012FROM tbl_Locations AS l INNER JO"
    "IN tlu_Parks AS lu ON lu.ParkCode = l.Unit_Code\015\012ORDER BY ParkName;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9dcdf5122222764f9e4df312924fe4b7
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="lu.ParkCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2316fb9c3fe6e2419f54a905e490449c
        End
    End
    Begin
        dbText "Name" ="lu.ParkName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x77da1d79c7288f4780151c1bb524dfe0
        End
    End
End
