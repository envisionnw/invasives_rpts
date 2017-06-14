dbMemo "SQL" ="SELECT DISTINCT (l.Plot_ID  & \"_\" & t.Transect) AS ID, l.Unit_Code, Year([Star"
    "t_Date]) AS Visit_Year, l.Plot_ID AS Route, t.Transect_ID, t.Transect, l.Area, t"
    ".E_Coord, t.N_Coord, q.Quadrat, esp.Position_m, esp.ColName, q.IsSampled, q.NoEx"
    "otics, sc.PlantCode, sc.IsDead, sc.PercentCover\015\012FROM (((tbl_Locations AS "
    "l LEFT JOIN EventSamplePosition AS esp ON esp.Location_ID = l.Location_ID) LEFT "
    "JOIN Transect AS t ON t.Event_ID = esp.Event_ID) LEFT JOIN Quadrat AS q ON q.Tra"
    "nsect_ID = t.Transect_ID) LEFT JOIN SpeciesCover AS sc ON sc.Quadrat_ID = q.ID\015"
    "\012WHERE esp.Quadrat = q.Quadrat\015\012ORDER BY l.Plot_ID, t.Transect, q.Quadr"
    "at;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbf5b1601b4ea244e88560a5f0f2b5222
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="((([Transect_Select_LIMITED_SpeciesCover].[Unit_Code]=\"GOSP\"))) AND ([Transect"
    "_Select_LIMITED_SpeciesCover].[Visit_Year]=2016)"
Begin
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5f46438f6768ff4f885da5434e0c776d
        End
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6758b335f44566469cf1431afc919d2e
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x44e88e83fb66934fb2811759310b6c7e
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x47437139e14e7c4292eab82d99dfa36a
        End
    End
    Begin
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1ec6846e275828459b0fff281e350e55
        End
    End
    Begin
        dbText "Name" ="t.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0ed00b674129eb4789a372565d45b601
        End
    End
    Begin
        dbText "Name" ="l.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x23af6e88ee798b46bfc3903abe671dcb
        End
    End
    Begin
        dbText "Name" ="t.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa5b8a9397db8ee4d843248390ff40521
        End
    End
    Begin
        dbText "Name" ="t.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4d968cb664890d46b3b2e0178579f821
        End
    End
    Begin
        dbText "Name" ="esp.Position_m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf15b4a1e95f60444af04f8ef868a00c9
        End
    End
    Begin
        dbText "Name" ="esp.ColName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd5cc454bd57dfb46a96cc971b6bec5d9
        End
    End
    Begin
        dbText "Name" ="q.IsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeda516d3fd719440a43d6d44607895ff
        End
    End
    Begin
        dbText "Name" ="q.NoExotics"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaf3c94c2f7f8bf468fd1c002df1d36ce
        End
    End
    Begin
        dbText "Name" ="sc.PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x429e0f099466454e8aa3389f9a063ddb
        End
    End
    Begin
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa47dfb4dc3fb9a4d896dec889a63b81e
        End
    End
    Begin
        dbText "Name" ="sc.PercentCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x33b63dfe473c8146ab5728a4dad3ef90
        End
    End
    Begin
        dbText "Name" ="q.Quadrat"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe2605e7dab365b4a852e548678b59a9c
        End
    End
End
