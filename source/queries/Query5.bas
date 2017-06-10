dbMemo "SQL" ="SELECT (l.Plot_ID  & \"_\" & t.Transect  & \"_\" &  Species  & \"_\" &  sc.IsDea"
    "d) AS ID, l.Unit_Code, Year([Start_Date]) AS Visit_Year, l.Plot_ID, t.Transect_I"
    "D, t.Transect, l.Area, IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),"
    "[Utah_Species],\015\012IIf([Unit_Code]=\"FOBU\",[WY_Species],[Co_Species])) AS S"
    "pecies, tlu_NCPN_Plants.Master_Common_Name, t.E_Coord, t.N_Coord, IIF(IsNull(sc."
    "PercentCover),0,sc.PercentCover) AS PercentCover, esp.Position_m, esp.ColName, s"
    "c.IsDead, q.IsSampled\015\012FROM ((((tbl_Locations AS l LEFT JOIN EventSamplePo"
    "sition AS esp ON esp.Location_ID = l.Location_ID) LEFT JOIN Transect AS t ON t.E"
    "vent_ID = esp.Event_ID) LEFT JOIN Quadrat AS q ON q.Transect_ID = t.Transect_ID)"
    " LEFT JOIN SpeciesCover AS sc ON sc.Quadrat_ID = q.ID) LEFT JOIN tlu_NCPN_Plants"
    " ON tlu_NCPN_Plants.Master_PLANT_Code = sc.PlantCode\015\012WHERE esp.Quadrat = "
    "q.Quadrat\015\012AND\015\012(((\015\012IIf([Unit_Code] In (\"CARE\",\"DINO\",\"G"
    "OSP\",\"ZION\"),[Utah_Species],\015\012IIf([Unit_Code]=\"FOBU\",[WY_Species],[Co"
    "_Species]))) Is Not Null))\015\012ORDER BY l.Plot_ID, t.Transect, IIf([Unit_Code"
    "] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],\015\012IIf([Unit_Code"
    "]=\"FOBU\",[WY_Species],[Co_Species]));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe5df587720525647a664d7dfa7809edc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="t.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Vert_Prec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Horz_Prec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Std_Dev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Protocol_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.ID"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Quadrat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.IsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.NoExotics"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Rcvr_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Elevation"
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
        dbText "Name" ="t.UTM_Zone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Datum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Feat_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Unfilt_Pos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Observer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Transect_ID"
        dbInteger "ColumnWidth" ="5115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.GPS_File_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Max_PDOP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Corr_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.GPS_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Update_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Recorder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Start_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Stop_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Observer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.version_key_number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Start_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3d2f085bb158944dba9cc5daeac9a949
        End
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8ff98a646e5a6246a434f1d54ccc2134
        End
    End
    Begin
        dbText "Name" ="l.Plot_ID"
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
            0xf5b61a8e51e1d245b17b1d74425e1416
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
            0xc3284a9c68604c43af2e5f109dad7550
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
