dbMemo "SQL" ="SELECT qry_rpt_Mo_Tr_Listing_Species_Detail.Unit_Code, qry_rpt_Mo_Tr_Listing_Spe"
    "cies_Detail.Visit_Year, qry_rpt_Mo_Tr_Listing_Species_Detail.Plot_ID, qry_rpt_Mo"
    "_Tr_Listing_Species_Detail.Transect, qry_rpt_Mo_Tr_Listing_Species_Detail.Area, "
    "qry_rpt_Mo_Tr_Listing_Species_Detail.PlantCode, qry_rpt_Mo_Tr_Listing_Species_De"
    "tail.IsDead, Sum(qry_rpt_Mo_Tr_Listing_Species_Detail.PercentCover) AS Percent_C"
    "over_Sum\015\012FROM qry_rpt_Mo_Tr_Listing_Species_Detail\015\012GROUP BY qry_rp"
    "t_Mo_Tr_Listing_Species_Detail.Unit_Code, qry_rpt_Mo_Tr_Listing_Species_Detail.V"
    "isit_Year, qry_rpt_Mo_Tr_Listing_Species_Detail.Plot_ID, qry_rpt_Mo_Tr_Listing_S"
    "pecies_Detail.Transect, qry_rpt_Mo_Tr_Listing_Species_Detail.Area, qry_rpt_Mo_Tr"
    "_Listing_Species_Detail.PlantCode, qry_rpt_Mo_Tr_Listing_Species_Detail.IsDead\015"
    "\012ORDER BY qry_rpt_Mo_Tr_Listing_Species_Detail.Plot_ID, qry_rpt_Mo_Tr_Listing"
    "_Species_Detail.Transect, qry_rpt_Mo_Tr_Listing_Species_Detail.PlantCode, qry_rp"
    "t_Mo_Tr_Listing_Species_Detail.IsDead;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc001cbd0e292aa4abfaf02112665d06c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Species_Detail.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x32182250c2e42940ad8d24666ceb935a
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Species_Detail.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7206ce13d002bf4a9350ad72d33e7dd3
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Species_Detail.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9a13c7ede573094ea6a01087ba6dcca9
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Species_Detail.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc2234861c654754db38e7abb3693b3f7
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Species_Detail.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb37a60df2464bc4c8c68d31b4abd64c2
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Species_Detail.PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x31db3e99488e214aa417cfa5fc574d42
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Species_Detail.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xee3778b8433fa040a0549253483767b1
        End
    End
    Begin
        dbText "Name" ="Percent_Cover_Sum"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x91287f65cfc3574083bf2916970db047
        End
    End
End
