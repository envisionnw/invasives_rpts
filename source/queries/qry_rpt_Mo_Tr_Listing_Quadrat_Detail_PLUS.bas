dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, Year([Start_Date]) AS Visit_Year, tbl_Locations."
    "Plot_ID AS Route, Transect.Transect, tbl_Locations.Area, Transect.E_Coord, Trans"
    "ect.N_Coord, Quadrat.Quadrat, Quadrat.IsSampled, Quadrat.NoExotics, IIf(tbl_Loca"
    "tions.Unit_Code In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],\015\012"
    "IIf(tbl_Locations.Unit_Code=\"FOBU\",[WY_Species],[Co_Species])) AS Species, sc."
    "IsDead\015\012FROM tbl_Locations INNER JOIN ((((tbl_Events INNER JOIN Transect O"
    "N tbl_Events.Event_ID = Transect.Event_ID) INNER JOIN Quadrat ON Transect.Transe"
    "ct_ID = Quadrat.Transect_ID) LEFT JOIN SpeciesCover AS sc ON sc.Quadrat_ID = Qua"
    "drat.ID) LEFT JOIN tlu_NCPN_Plants ON tlu_NCPN_Plants.Master_PLANT_Code = sc.Pla"
    "ntCode) ON tbl_Locations.Location_ID = tbl_Events.Location_ID\015\012WHERE (((tb"
    "l_Locations.Unit_Code)=\"gosp\") AND ((Year([Start_Date]))=2016))\015\012ORDER B"
    "Y tbl_Locations.Plot_ID, Transect.Transect, Quadrat.Quadrat;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBinary "GUID" = Begin
    0xd2ff503a6bb2b94d976cf859514b1620
End
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x304f64ea81046c4790131ad462a7cc41
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd077f7927d94464fb5dd8949a1535456
        End
    End
    Begin
        dbText "Name" ="Transect.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xed1eed0b0bfc9d418316d1cd958a6dcd
        End
    End
    Begin
        dbText "Name" ="tbl_Locations.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x389a360999d383409a0db5aa1c47a894
        End
    End
    Begin
        dbText "Name" ="Transect.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x96a98a8f153273499fca3395d5701546
        End
    End
    Begin
        dbText "Name" ="Transect.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf60250b9079cc74cb9a99431fd8b7dae
        End
    End
    Begin
        dbText "Name" ="Quadrat.Quadrat"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe5dfbdc9cef2644c98bfed70cac02877
        End
    End
    Begin
        dbText "Name" ="Quadrat.IsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xeea4f3879e5c004fbf80985c904d787b
        End
    End
    Begin
        dbText "Name" ="Quadrat.NoExotics"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x59c9d1284bc02548b2fe28d4ae787e1d
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3f2e2ea650bf5641ba0da4b5d3675068
        End
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2432d351ab1ec54f8bb6729100aeb40a
        End
    End
    Begin
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
    End
End
