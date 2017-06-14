dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, Year([Start_Date]) AS Visit_Year, tbl_Locations."
    "Plot_ID AS Route, COUNT(Transect) AS TransectsPerRoute\015\012FROM tbl_Locations"
    " LEFT JOIN (tbl_Events LEFT JOIN Transect ON tbl_Events.Event_ID = Transect.Even"
    "t_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID\015\012GROUP BY tbl_"
    "Locations.Unit_Code, Year([Start_Date]), tbl_Locations.Plot_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "OrderBy" ="[Query4].[Unit_Code], [Query4].[Route]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8bba4a2fe1c8ac409f74fd6bde7f7e29
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x19dcd13fce0bbe48bf9f8710aafc1933
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x49411dcc2f37604b97231811e49aa24d
        End
    End
    Begin
        dbText "Name" ="TransectsPerRoute"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7b03bb16e8bca744a9a9311513ae35ea
        End
    End
    Begin
        dbText "Name" ="Transect.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
End
