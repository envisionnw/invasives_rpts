dbMemo "SQL" ="SELECT (l.Plot_ID  & \"_\" & t.Transect  & \"_\" &  Species  & \"_\" &  sc.IsDea"
    "d) AS ID, l.Unit_Code, Year([Start_Date]) AS Visit_Year, l.Plot_ID AS Route, t.T"
    "ransect_ID, t.Transect, l.Area, IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\""
    "ZION\"),[Utah_Species],\015\012IIf([Unit_Code]=\"FOBU\",[WY_Species],[Co_Species"
    "])) AS Species, tlu_NCPN_Plants.Master_Common_Name, t.E_Coord, t.N_Coord, IIF(Is"
    "Null(sc.PercentCover),0,sc.PercentCover) AS PercentCover, esp.Position_m, esp.Co"
    "lName, sc.IsDead, q.IsSampled\015\012FROM ((((tbl_Locations AS l LEFT JOIN Event"
    "SamplePosition AS esp ON esp.Location_ID = l.Location_ID) LEFT JOIN Transect AS "
    "t ON t.Event_ID = esp.Event_ID) LEFT JOIN Quadrat AS q ON q.Transect_ID = t.Tran"
    "sect_ID) LEFT JOIN SpeciesCover AS sc ON sc.Quadrat_ID = q.ID) LEFT JOIN tlu_NCP"
    "N_Plants ON tlu_NCPN_Plants.Master_PLANT_Code = sc.PlantCode\015\012WHERE esp.Qu"
    "adrat = q.Quadrat\015\012AND\015\012(((\015\012IIf([Unit_Code] In (\"CARE\",\"DI"
    "NO\",\"GOSP\",\"ZION\"),[Utah_Species],\015\012IIf([Unit_Code]=\"FOBU\",[WY_Spec"
    "ies],[Co_Species]))) Is Not Null))\015\012ORDER BY l.Plot_ID, t.Transect, IIf([U"
    "nit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],\015\012IIf([U"
    "nit_Code]=\"FOBU\",[WY_Species],[Co_Species]));\015\012"
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
    0xd2b1b62b4fe1504b852ccf794b82ad8e
End
dbMemo "Filter" ="(((Transect_Select.Unit_Code=\"GOSP\"))) And (Transect_Select.Visit_Year=2016)"
Begin
    Begin
        dbText "Name" ="Visit_Year"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x62c44a96f682974eb1392781ba0c2b05
        End
    End
    Begin
        dbText "Name" ="Species"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1a35b91843572b4585fc45dc92c4987a
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PercentCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x67b7e637c2071741b977720786d3bc11
        End
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="esp.Position_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="esp.ColName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x604e959de516474e915ea1873e49f8f1
        End
    End
    Begin
        dbText "Name" ="q.IsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbc5f812ffdd6f142958450f982019fc2
        End
    End
End
