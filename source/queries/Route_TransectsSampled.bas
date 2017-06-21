dbMemo "SQL" ="SELECT rt.Unit_Code, rt.Visit_Year, rt.Route, MIN(t.TransectCount) AS TransectCo"
    "unt, SUM(rt.TransectSampled) AS TransectsSampled\015\012FROM Route_Transect AS r"
    "t LEFT JOIN Route_Transects AS t ON (t.Route = rt.Route) AND (t.Visit_Year = rt."
    "Visit_Year) AND (t.Unit_Code = rt.Unit_Code)\015\012GROUP BY rt.Unit_Code, rt.Vi"
    "sit_Year, rt.Route\015\012ORDER BY rt.Unit_Code, rt.Visit_Year, rt.Route;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xcf27e96222e5a14c90a868d3d11f8662
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="rt.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xccbfb7f9519c784bbd555502748c92a1
        End
    End
    Begin
        dbText "Name" ="rt.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x919426fe4ee6484e9684cc3558d73ffc
        End
    End
    Begin
        dbText "Name" ="rt.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4684b3238b733e408493c60fcb076c40
        End
    End
    Begin
        dbText "Name" ="TransectsSampled"
        dbInteger "ColumnWidth" ="2475"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3aad3e9150d38140ab8b70749ad06c2e
        End
    End
End
