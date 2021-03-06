﻿dbMemo "SQL" ="SELECT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.PlantCode, ts.IsDead, SUM(ts.Tr"
    "ansectAverageCover) AS TotalTransectAverageCover\015\012FROM Transect_AverageCov"
    "er AS ts\015\012GROUP BY ts.Unit_Code, ts.Visit_Year, ts.Route, ts.PlantCode, ts"
    ".IsDead\015\012ORDER BY ts.Unit_Code, ts.Visit_Year, ts.Route, ts.PlantCode, ts."
    "IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb359c963405c464399f5459c23d403a1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3501ac158e2b9349809022ef9d44c95d
        End
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf46c019398fe8e48b1073c611d76797f
        End
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9e578773affb32478037ab669586f349
        End
    End
    Begin
        dbText "Name" ="ts.PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7c99d011aa5dad43ad73e2812650b0dc
        End
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa3fdd27c486bff43bbe175732b5cd1a1
        End
    End
    Begin
        dbText "Name" ="TotalTransectAverageCover"
        dbLong "AggregateType" ="-1"
    End
End
