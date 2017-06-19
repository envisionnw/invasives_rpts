dbMemo "SQL" ="SELECT l.Unit_Code, l.Plot_ID AS Route, Year([Start_Date]) AS Visit_Year, Count("
    "t.Transect) AS TransectCount\015\012FROM (tbl_Locations AS l LEFT JOIN tbl_Event"
    "s AS e ON e.Location_ID = l.Location_ID) LEFT JOIN Transect AS t ON t.Event_ID ="
    " e.Event_ID\015\012GROUP BY l.Unit_Code, l.Plot_ID, Year([Start_Date])\015\012HA"
    "VING (((l.Unit_Code)=Forms!frm_Select_Transect_Counts!Park_Code) \015\012AND ((Y"
    "ear([Start_Date]))=Forms!frm_Select_Transect_Counts!Visit_Year));\015\012"
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
    0x044edeb3832a3442822f123f78d6dc91
End
Begin
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x469fc7ec59f955488dee3f63701ae5a9
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8751ad407fa76f4fb95487a77b356c9b
        End
    End
    Begin
        dbText "Name" ="TransectCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x000050c8feca4744b434dcb1dc952b1b
        End
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
