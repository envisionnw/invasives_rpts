dbMemo "SQL" ="SELECT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, ts.Area, ts.E_Coord, "
    "ts.N_Coord, SUM(ts.IsSampled) AS SampledQuadrats\015\012FROM Transect_Select_Qua"
    "drat AS ts\015\012GROUP BY ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, t"
    "s.Area, ts.E_Coord, ts.N_Coord\015\012ORDER BY ts.Unit_Code, ts.Visit_Year, ts.R"
    "oute, ts.Transect, ts.Area, ts.E_Coord, ts.N_Coord;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="((([Query6].[Unit_Code]=\"GOSP\"))) AND ([Query6].[Visit_Year]=2016)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9165fbf4b587f349add22b24ff0ff91e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1e53e4e9ec630a4a9ea1894ea4939e71
        End
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x85bfd908b52d9c42b706d40202aebcda
        End
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8bd409eb54f3814d892a323e643a11c1
        End
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1d542bcfb40e12489e43c738d82bf8a2
        End
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb6b1e6b1a9e7774da9ba6336cd512cce
        End
    End
    Begin
        dbText "Name" ="ts.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd534e44b761df440a9d31a6f46331a34
        End
    End
    Begin
        dbText "Name" ="ts.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe8068db3dbfba244bc355fdea957b69a
        End
    End
    Begin
        dbText "Name" ="SampledQuadrats"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf7a8cc3f3ce6f144adbab5af8b4965b0
        End
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
    End
End
