dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, Year([Start_Date]) AS Visit_Year, tbl_Locations."
    "Plot_ID AS Route, Transect.Transect, Transect.Transect_ID\015\012FROM tbl_Locati"
    "ons LEFT JOIN (tbl_Events LEFT JOIN Transect ON tbl_Events.Event_ID = Transect.E"
    "vent_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID\015\012GROUP BY t"
    "bl_Locations.Unit_Code, Year([Start_Date]), tbl_Locations.Plot_ID, Transect.Tran"
    "sect, Transect.Transect_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBinary "GUID" = Begin
    0xe6c5411112cc114bae83af5ac81756e9
End
Begin
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x764fe0c3bd115841bd33399a46fc08f4
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa70dec5d513bb14fba613af2c65944e9
        End
    End
    Begin
        dbText "Name" ="Transect.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
End
