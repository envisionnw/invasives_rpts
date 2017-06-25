dbMemo "SQL" ="SELECT qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Unit_Code, qry_rpt_Mo_Tr_Listing_Qua"
    "drat_Detail.Visit_Year, qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Plot_ID, qry_rpt_Mo"
    "_Tr_Listing_Quadrat_Detail.Transect, qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Area, "
    "qry_rpt_Mo_Tr_Listing_Quadrat_Detail.E_Coord, qry_rpt_Mo_Tr_Listing_Quadrat_Deta"
    "il.N_Coord, Sum(qry_rpt_Mo_Tr_Listing_Quadrat_Detail.IsSampled) AS Quadrats_Samp"
    "led\015\012FROM qry_rpt_Mo_Tr_Listing_Quadrat_Detail\015\012GROUP BY qry_rpt_Mo_"
    "Tr_Listing_Quadrat_Detail.Unit_Code, qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Visit_"
    "Year, qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Plot_ID, qry_rpt_Mo_Tr_Listing_Quadra"
    "t_Detail.Transect, qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Area, qry_rpt_Mo_Tr_List"
    "ing_Quadrat_Detail.E_Coord, qry_rpt_Mo_Tr_Listing_Quadrat_Detail.N_Coord\015\012"
    "ORDER BY qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Plot_ID, qry_rpt_Mo_Tr_Listing_Qua"
    "drat_Detail.Transect;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x3fc0b3a36a021941b86548775275dfa4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="((([qry_rpt_Mo_Tr_Listing_Quadrats_Sampled].[Unit_Code]=\"CARE\"))) AND ([qry_rp"
    "t_Mo_Tr_Listing_Quadrats_Sampled].[Visit_Year]=2015)"
Begin
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x84c09046ae9fbc43bb3a6c497ce0f8df
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x35e56f8bb438664c813b2f8aa669d720
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcbb50c90c57a29489b692eca66dec8e2
        End
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x127a66d693d07f4bb2e3fae013558d48
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrat_Detail.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9d466c3407a116429af3ae536ea0c1f6
        End
        dbInteger "ColumnWidth" ="435"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrat_Detail.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf17eae790a335f4f8cb8134c0dee1fe8
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrat_Detail.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf5a1d1ae914bc14b83b6c5efdb857b07
        End
    End
    Begin
        dbText "Name" ="Quadrats_Sampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4be9dfc665dae947bbffd495f5b7476a
        End
    End
End
