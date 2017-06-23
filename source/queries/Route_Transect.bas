dbMemo "SQL" ="SELECT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Area, ts.Transect, IIF(SUM(ts.I"
    "sSampled)>0,1,0) AS TransectSampled, SUM(ts.IsSampled) AS QuadratsSampled\015\012"
    "FROM Transect_Select_Quadrat AS ts\015\012GROUP BY ts.Unit_Code, ts.Visit_Year, "
    "ts.Route, ts.Area, ts.Transect\015\012ORDER BY ts.Unit_Code, ts.Visit_Year, ts.R"
    "oute, ts.Area, ts.Transect;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa6c213a146059544a3f1352ed69df017
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x92bca51a8b376440bb0428ac450a5924
        End
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xff0bbe08b360fc41beb440fcc0935249
        End
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x15eb0d2f00d5d94595e9d9a11e7468a6
        End
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb22d9b1218e5ed49ad90e6be5f6462f4
        End
    End
    Begin
        dbText "Name" ="TransectSampled"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="QuadratsSampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
    End
End
