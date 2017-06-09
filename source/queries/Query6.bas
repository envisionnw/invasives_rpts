dbMemo "SQL" ="SELECT l.Unit_Code, Year([Start_Date]) AS Visit_Year, l.Plot_ID, t.Transect_ID, "
    "t.Transect, l.Area, IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Ut"
    "ah_Species],\015\012IIf([Unit_Code]=\"FOBU\",[WY_Species],[Co_Species])) AS Spec"
    "ies, tlu_NCPN_Plants.Master_Common_Name, t.E_Coord, t.N_Coord, IIF(IsNull(sc.Per"
    "centCover),0,sc.PercentCover) AS PercentCover, esp.Position_m, esp.ColName, sc.I"
    "sDead\015\012FROM ((((tbl_Locations AS l LEFT JOIN EventSamplePosition AS esp ON"
    " esp.Location_ID = l.Location_ID) LEFT JOIN Transect AS t ON t.Event_ID = esp.Ev"
    "ent_ID) LEFT JOIN Quadrat AS q ON q.Transect_ID = t.Transect_ID) LEFT JOIN Speci"
    "esCover AS sc ON sc.Quadrat_ID = q.ID) LEFT JOIN tlu_NCPN_Plants ON tlu_NCPN_Pla"
    "nts.Master_PLANT_Code = sc.PlantCode\015\012WHERE esp.Quadrat = q.Quadrat\015\012"
    "AND\015\012(((\015\012IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),["
    "Utah_Species],\015\012IIf([Unit_Code]=\"FOBU\",[WY_Species],[Co_Species]))) Is N"
    "ot Null))\015\012ORDER BY l.Plot_ID, t.Transect, IIf([Unit_Code] In (\"CARE\",\""
    "DINO\",\"GOSP\",\"ZION\"),[Utah_Species],\015\012IIf([Unit_Code]=\"FOBU\",[WY_Sp"
    "ecies],[Co_Species]));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x054728e21c3eee488d8651fc1eb4dd57
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
            0x85da5fee3f31244785bbf47115e6be6f
        End
    End
    Begin
        dbText "Name" ="l.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Transect_ID"
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
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x12eaca863c3dd64f98a7ce20becd9f0b
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Common_Name"
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
        dbText "Name" ="PercentCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6b7110ef092d7841a6c97335fd7481bd
        End
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
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
    End
End
