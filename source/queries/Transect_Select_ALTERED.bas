dbMemo "SQL" ="SELECT (l.Plot_ID  & \"_\" & t.Transect  & \"_\" &  Species  & \"_\" &  sc.IsDea"
    "d) AS ID, l.Unit_Code, Year([Start_Date]) AS Visit_Year, l.Plot_ID AS Route, t.T"
    "ransect_ID, t.Transect, l.Area, IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\""
    "ZION\"),[Utah_Species],\015\012IIf([Unit_Code]=\"FOBU\",[WY_Species],[Co_Species"
    "])) AS Species, tlu_NCPN_Plants.Master_Common_Name, t.E_Coord, t.N_Coord, IIF(Is"
    "Null(sc.PercentCover),0,sc.PercentCover) AS PercentCover, esp.Position_m, esp.Co"
    "lName, sc.IsDead, q.IsSampled, q.NoExotics\015\012FROM ((((tbl_Locations AS l LE"
    "FT JOIN EventSamplePosition AS esp ON esp.Location_ID = l.Location_ID) LEFT JOIN"
    " Transect AS t ON t.Event_ID = esp.Event_ID) LEFT JOIN Quadrat AS q ON q.Transec"
    "t_ID = t.Transect_ID) LEFT JOIN SpeciesCover AS sc ON sc.Quadrat_ID = q.ID) LEFT"
    " JOIN tlu_NCPN_Plants ON tlu_NCPN_Plants.Master_PLANT_Code = sc.PlantCode\015\012"
    "WHERE esp.Quadrat = q.Quadrat\015\012AND\015\012(((\015\012IIf([Unit_Code] In (\""
    "CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],\015\012IIf([Unit_Code]=\"FOBU"
    "\",[WY_Species],[Co_Species]))) Is Not Null))\015\012ORDER BY l.Plot_ID, t.Trans"
    "ect, IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],\015"
    "\012IIf([Unit_Code]=\"FOBU\",[WY_Species],[Co_Species]));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Query5].[Unit_Code]=\"GOSP\"))) AND ([Query5].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xeb4f8358debde342b3f666751bb224ca
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6e634bcbb60bc24483c973750d80a36c
        End
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcdde00dc670c2a47a549fa81530d5d52
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdbed79e3c5948c4bbc0eb73c5fc8a3a8
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3caae319f462784cae9ba1f58cbebb2b
        End
    End
    Begin
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf1dcf88799b5224f8edc3a614abd5eea
        End
    End
    Begin
        dbText "Name" ="t.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb025f3446ed5954baa29fc2ccc0fd63e
        End
    End
    Begin
        dbText "Name" ="q.IsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe8d9810a0c6ede43a63abe11591a8e91
        End
    End
    Begin
        dbText "Name" ="q.NoExotics"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe9cff3ece97e02489c6688a30e5a69d8
        End
    End
    Begin
        dbText "Name" ="l.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaf1e1668302a0442af7c5e9d655b40ce
        End
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe3133656cb795f4e91230c6cdf69a733
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbaf56bcb48abaf49bee9c78ebdc8b6ce
        End
    End
    Begin
        dbText "Name" ="t.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfe76269673b47940af61bf144651c3ab
        End
    End
    Begin
        dbText "Name" ="t.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3c7bd1d8129974469511c36de44944a6
        End
    End
    Begin
        dbText "Name" ="PercentCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa5c118a81e85ca42bcf3fb5b4f2d22b4
        End
    End
    Begin
        dbText "Name" ="esp.Position_m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x533f3d58056fd648a9f56d1b9d82ef53
        End
    End
    Begin
        dbText "Name" ="esp.ColName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6972311d587dbf4ea2986848219ae010
        End
    End
    Begin
        dbText "Name" ="sc.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x921491a99991914a9e6a183b4884187a
        End
    End
End
