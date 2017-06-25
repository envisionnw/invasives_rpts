dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, Year([Start_Date]) AS Visit_Year, tbl_Locations."
    "Plot_ID, Transect.Transect, tbl_Locations.Area, SpeciesCover.PlantCode, SpeciesC"
    "over.IsDead, SpeciesCover.PercentCover\015\012FROM (tbl_Locations INNER JOIN (tb"
    "l_Events INNER JOIN Transect ON tbl_Events.Event_ID = Transect.Event_ID) ON tbl_"
    "Locations.Location_ID = tbl_Events.Location_ID) INNER JOIN (Quadrat INNER JOIN S"
    "peciesCover ON Quadrat.ID =\015\012SpeciesCover.Quadrat_ID) ON Transect.Transect"
    "_ID = Quadrat.Transect_ID\015\012WHERE (((tbl_Locations.Unit_Code)=\"care\") AND"
    " ((Year([Start_Date]))=2015))\015\012ORDER BY tbl_Locations.Plot_ID, Transect.Tr"
    "ansect, SpeciesCover.PlantCode;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x7c07f8c582665c439b4e8a0ac0eb55d1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x513f2ac1f759d5428bdba6c3d9647c5a
        End
        dbInteger "ColumnWidth" ="915"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7ab504f5ffc99840aea6cc5176a3bbbf
        End
        dbInteger "ColumnWidth" ="1035"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6becc557e1b4bf498db9b958d27ccad8
        End
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Transect.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x54e6090f56379b40baca671cb374d209
        End
        dbInteger "ColumnWidth" ="855"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfdc3151082638141aa9e41f637a4fb41
        End
    End
    Begin
        dbText "Name" ="SpeciesCover.PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb99273f42f38b641b108b9c6153907e4
        End
    End
    Begin
        dbText "Name" ="SpeciesCover.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa3507e291eb3e840b18319a4d4840088
        End
        dbInteger "ColumnWidth" ="510"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SpeciesCover.PercentCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc3933e294c85964eb8c3773a8028178b
        End
        dbInteger "ColumnWidth" ="825"
        dbBoolean "ColumnHidden" ="0"
    End
End
