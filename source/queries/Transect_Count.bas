dbMemo "SQL" ="SELECT l.Unit_Code, l.Plot_ID AS Route, Year([Start_Date]) AS Visit_Year, Count("
    "t.Transect) AS TransectCount\015\012FROM (tbl_Locations AS l LEFT JOIN tbl_Event"
    "s AS e ON e.Location_ID = l.Location_ID) LEFT JOIN Transect AS t ON t.Event_ID ="
    " e.Event_ID\015\012GROUP BY l.Unit_Code, l.Plot_ID, Year([Start_Date])\015\012HA"
    "VING (((Year([Start_Date])) Is Not Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x6238a39bc2031b4c8b58c0d219f125bc
End
Begin
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x602e2b5d3de459448b72a12ca1b344bf
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x19e5cf8ff7b6164c939ff99e02816d9a
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransectCount"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7f2292987b6e6f4f96d5251fc36ce28a
        End
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
