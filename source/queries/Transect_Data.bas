dbMemo "SQL" ="SELECT tsc.Unit_Code, tsc.Visit_Year, tsc.Plot_ID, tsc.Transect, tsc.Area, tsc.E"
    "_Coord, tsc.N_Coord, tsc.Species, tsc.Master_Common_Name, IIF(tsc.IsDead = 1, \""
    "No\", \"Yes\") AS [Alive?], tsc.AverageCover\015\012FROM Transect_Select_Count A"
    "S tsc\015\012ORDER BY tsc.Unit_Code, tsc.Visit_Year, tsc.Plot_ID, tsc.Transect, "
    "tsc.Area, tsc.Species;\015\012"
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
dbMemo "Filter" ="([Transect_Data].Visit_Year=2012)"
Begin
    Begin
        dbText "Name" ="tsc.AverageCover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Alive?"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa6075aa7909caa4d8b2e6d651a0104f3
        End
    End
    Begin
        dbText "Name" ="tsc.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsc.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
End
