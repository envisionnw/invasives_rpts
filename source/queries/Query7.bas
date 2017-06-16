dbMemo "SQL" ="SELECT (l.Plot_ID  & \"_\" & t.Transect & \"_\" &  sc.IsDead) AS ID, l.Unit_Code"
    ", Year([Start_Date]) AS Visit_Year, l.Plot_ID AS Route, t.Transect_ID, t.Transec"
    "t, l.Area, t.E_Coord, t.N_Coord, IIF(IsNull(sc.PercentCover),0,sc.PercentCover) "
    "AS PercentCover, q.ID AS Quadrat_ID, q.Quadrat, sc.IsDead, q.IsSampled, q.NoExot"
    "ics\015\012FROM (((tbl_Locations AS l LEFT JOIN tbl_Events AS e ON e.Location_ID"
    " = l.Location_ID) LEFT JOIN Transect AS t ON t.Event_ID = e.Event_ID) LEFT JOIN "
    "Quadrat AS q ON q.Transect_ID = t.Transect_ID) LEFT JOIN SpeciesCover AS sc ON s"
    "c.Quadrat_ID = q.ID\015\012ORDER BY l.Plot_ID, t.Transect;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Query7].[Unit_Code]=\"GOSP\"))) AND ([Query7].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5e72509f1d79694c8536a4c26b12075d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Route"
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
    End
    Begin
        dbText "Name" ="Quadrat_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Quadrat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.IsDead"
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
End
