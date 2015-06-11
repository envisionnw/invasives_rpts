dbMemo "SQL" ="SELECT Last_Modified\015\012FROM tbl_Target_List\015\012WHERE Park_Code LIKE 'BL"
    "CA' AND Target_Year = 2017;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4980d777ef6b214b8c480cb18ed43a4f
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Last_Modified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_List.Last_Modified"
        dbLong "AggregateType" ="-1"
    End
End
