dbMemo "SQL" ="SELECT l.Unit_Code, Year([Start_Date]) AS Visit_Year, l.Plot_ID AS Route, COUNT("
    "Transect) AS TransectsPerRoute\015\012FROM tbl_Locations AS l LEFT JOIN (tbl_Eve"
    "nts AS e LEFT JOIN Transect AS t ON e.Event_ID = t.Event_ID) ON l.Location_ID = "
    "e.Location_ID\015\012GROUP BY l.Unit_Code, Year([Start_Date]), l.Plot_ID\015\012"
    "ORDER BY l.Unit_Code, Year([Start_Date]), l.Plot_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Query7].[Unit_Code]=\"GOSP\"))) AND ([Query7].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5e72509f1d79694c8536a4c26b12075d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xedb1d31cd56e334aa209b5084d87f172
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5abac1d343f6ed439c0588f62f328fa6
        End
    End
    Begin
        dbText "Name" ="TransectsPerRoute"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4a2fe6bbaf8f0148a3855328fbeb7cbb
        End
    End
End
