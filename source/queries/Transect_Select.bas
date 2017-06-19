dbMemo "SQL" ="SELECT l.Unit_Code, Year([Start_Date]) AS Visit_Year, e.Event_ID, l.Location_ID,"
    " l.Plot_ID AS Route, t.Transect_ID, t.Transect, l.Area, t.E_Coord, t.N_Coord, q."
    "ID AS Quadrat_ID, q.Quadrat, q.IsSampled, q.NoExotics\015\012FROM ((tbl_Location"
    "s AS l LEFT JOIN tbl_Events AS e ON e.Location_ID = l.Location_ID) LEFT JOIN Tra"
    "nsect AS t ON t.Event_ID = e.Event_ID) LEFT JOIN Quadrat AS q ON q.Transect_ID ="
    " t.Transect_ID\015\012ORDER BY l.Plot_ID, t.Transect;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Transect_Select_LIMITED].[Unit_Code]=\"GOSP\"))) AND ([Transect_Select_LIMIT"
    "ED].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc2dee0d27fd5bb4896a7932686be7a96
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x175c495d2f2dcb408a1572fda3973208
        End
    End
    Begin
        dbText "Name" ="Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3a0f4fdf76029547a729b751455b8515
        End
    End
    Begin
        dbText "Name" ="t.Transect_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00056997e8add942842fc87db5d8e7c8
        End
    End
    Begin
        dbText "Name" ="t.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbc8d6d70517f90428ad870e9118d508f
        End
    End
    Begin
        dbText "Name" ="l.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2cd7aa5bb07ca645a4fa6644c2fba253
        End
    End
    Begin
        dbText "Name" ="t.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x813c2ae2ec524c44bf70bf3e6633f68e
        End
    End
    Begin
        dbText "Name" ="t.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaab660d8da85104ab1e9bea1c5861bc6
        End
    End
    Begin
        dbText "Name" ="Quadrat_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd8c4bd17fe63b6429a58a7976f7760ae
        End
    End
    Begin
        dbText "Name" ="q.Quadrat"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3a7f08015e6a924ab87231745644ac7d
        End
    End
    Begin
        dbText "Name" ="q.IsSampled"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x41fd7a1d6a1e234aaf2291e7d960a079
        End
    End
    Begin
        dbText "Name" ="q.NoExotics"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfa37e8b2e7c93e4590871c18b84f4348
        End
    End
    Begin
        dbText "Name" ="l.Location_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9bc3cdd2420f6542a8e54e85610f946f
        End
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
End
