dbMemo "SQL" ="SELECT l.Unit_Code, Year(e.Start_Date) AS Visit_Year, l.Plot_ID AS Route, LEFT(l"
    ".Plot_ID, 48) AS RouteTruncated, COUNT(t.Transect) AS TransectCount\015\012FROM "
    "(Transect AS t LEFT JOIN tbl_Events AS e ON e.Event_ID = t.Event_ID) LEFT JOIN t"
    "bl_Locations AS l ON l.Location_ID = e.Location_ID\015\012GROUP BY l.Unit_Code, "
    "Year(e.Start_Date), l.Plot_ID\015\012ORDER BY l.Unit_Code, Year(e.Start_Date), l"
    ".Plot_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x458ba3feb891d04fbe6c70df5a5f965f
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdf444e3d1a50124d8420bc513d3e60f3
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x48980e8d1022ed47bc6582bfd2d37880
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x18fe706557f48c49ac3146f507820773
        End
        dbInteger "ColumnWidth" ="4665"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TransectCount"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa6cd110f6bce5840bba7caf04fa7cea7
        End
    End
    Begin
        dbText "Name" ="RouteTruncated"
        dbInteger "ColumnWidth" ="3990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
