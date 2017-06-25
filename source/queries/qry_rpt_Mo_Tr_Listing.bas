dbMemo "SQL" ="SELECT qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.Unit_Code, qry_rpt_Mo_Tr_Listing_Q"
    "uadrats_Sampled.Visit_Year, qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.Plot_ID, qry_"
    "rpt_Mo_Tr_Listing_Quadrats_Sampled.Transect AS Transect_Number, qry_rpt_Mo_Tr_Li"
    "sting_Quadrats_Sampled.Area, qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.E_Coord, qry"
    "_rpt_Mo_Tr_Listing_Quadrats_Sampled.N_Coord, tlu_NCPN_Plants.Utah_Species, tlu_N"
    "CPN_Plants.Master_Common_Name, qry_rpt_Mo_Tr_Listing_Species_Cover_Sums.IsDead, "
    "qry_rpt_Mo_Tr_Listing_Species_Cover_Sums.Percent_Cover_Sum, qry_rpt_Mo_Tr_Listin"
    "g_Quadrats_Sampled.Quadrats_Sampled, [qry_rpt_Mo_Tr_Listing_Species_Cover_Sums]!"
    "[Percent_Cover_Sum]/[qry_rpt_Mo_Tr_Listing_Quadrats_Sampled]![Quadrats_Sampled] "
    "AS Average_Cover\015\012FROM (qry_rpt_Mo_Tr_Listing_Quadrats_Sampled LEFT JOIN q"
    "ry_rpt_Mo_Tr_Listing_Species_Cover_Sums ON (qry_rpt_Mo_Tr_Listing_Quadrats_Sampl"
    "ed.Unit_Code = qry_rpt_Mo_Tr_Listing_Species_Cover_Sums.Unit_Code) AND (qry_rpt_"
    "Mo_Tr_Listing_Quadrats_Sampled.Visit_Year = qry_rpt_Mo_Tr_Listing_Species_Cover_"
    "Sums.Visit_Year) AND (qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.Plot_ID = qry_rpt_M"
    "o_Tr_Listing_Species_Cover_Sums.Plot_ID) AND (qry_rpt_Mo_Tr_Listing_Quadrats_Sam"
    "pled.Transect = qry_rpt_Mo_Tr_Listing_Species_Cover_Sums.Transect)) LEFT JOIN tl"
    "u_NCPN_Plants ON qry_rpt_Mo_Tr_Listing_Species_Cover_Sums.PlantCode =\015\012tlu"
    "_NCPN_Plants.Master_PLANT_Code\015\012WHERE (((tlu_NCPN_Plants.Utah_Species) Is "
    "Not Null))\015\012ORDER BY qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.Plot_ID, qry_r"
    "pt_Mo_Tr_Listing_Quadrats_Sampled.Transect, tlu_NCPN_Plants.Utah_Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0c5b3022aac41c4bb11deff23bb26dbd
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x477d0c5f6cdd534886c0db52fbe83eed
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xccbd6be580982047bd7c5d4549b6bb7f
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5a65856525f2a4499b46bd03e25c6995
        End
    End
    Begin
        dbText "Name" ="Transect_Number"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5dde3cda717c3340b5e6239c5a9345d0
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x601b9cac0eb30c4bb829194d940b2e14
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd98e2a4a98fdd1409773800e3d60fbb0
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7a4993b52c573f45b52e8f5b022509e0
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Utah_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x97101c99f7e0c241bc6c3d61bc78e8ba
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6bdcd07788641743aec85f1b3f5ea767
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Species_Cover_Sums.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8a91a8b3cfe4f045bfc4d58a029fedb7
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Species_Cover_Sums.Percent_Cover_Sum"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd91421e5a1b5a0469a821cccc01d116f
        End
    End
    Begin
        dbText "Name" ="qry_rpt_Mo_Tr_Listing_Quadrats_Sampled.Quadrats_Sampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xde3b205845df4043844fe3f9c934f156
        End
    End
    Begin
        dbText "Name" ="Average_Cover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf8a6a4aff34281438de14baaffd9ba1f
        End
    End
End
