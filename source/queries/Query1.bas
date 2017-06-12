dbMemo "SQL" ="SELECT tsca.Unit_Code, tsca.Visit_Year, tsca.Plot_ID, tsca.Transect, tsca.Area, "
    "tsca.E_Coord, tsca.N_Coord, tsca.Species, tsca.Master_Common_Name AS Common_Name"
    ", tsca.IsDead, tsca.Q1_0_5m, tsca.Q2_4_5m, tsca.Q3_9_5m, tsca.Q1_3m, tsca.Q2_8m,"
    " tsca.Q3_13m, tsca.Q1, tsca.Q2, tsca.Q3, tsca.QuadratsSampled, tsca.TotalCover, "
    "tsca.AverageCover, tc.TransectCount\015\012FROM Transect_Select_Crosstab_with_Av"
    "erageCover AS tsca INNER JOIN Transect_Count AS tc ON (tc.Unit_Code = tsca.Unit_"
    "Code) AND (tc.Visit_Year = tsca.Visit_Year) AND (tc.Route = tsca.Plot_ID);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBinary "GUID" = Begin
    0xb44113b794912c4fb2bcf2c74d45260c
End
Begin
    Begin
        dbText "Name" ="e.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbInteger "ColumnWidth" ="4110"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Protocol_Name"
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
        dbText "Name" ="e.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Observer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qp.SamplingYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qp.Quadrat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qp.Position_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ColName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x28c799490d8e31468948bfb2547b9b59
        End
    End
    Begin
        dbText "Name" ="tsca.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Q1_0_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Q2_4_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Q3_9_5m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Q2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.TotalCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Q1_3m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1a0ab7c98df8614fa70b1309a14e2d18
        End
    End
    Begin
        dbText "Name" ="tsca.Q3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Q2_8m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.QuadratsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.TransectCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Q3_13m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.Q1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsca.AverageCover"
        dbLong "AggregateType" ="-1"
    End
End
