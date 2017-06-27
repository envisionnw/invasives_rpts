dbMemo "SQL" ="SELECT DISTINCT rt.Unit_Code, rt.Visit_Year, rt.Route, rt.Area, rt.Transect, rt."
    "E_Coord, rt.N_Coord, rs.PlantCode, rs.IsDead, rt.QuadratsSampled AS QuadratsSamp"
    "led\015\012FROM Route_Transect AS rt INNER JOIN Route_Species AS rs ON (rs.Route"
    " = rt.Route) AND (rs.Visit_Year = rt.Visit_Year) AND (rs.Unit_Code = rt.Unit_Cod"
    "e)\015\012ORDER BY rt.Unit_Code, rt.Visit_Year, rt.Route, rt.Area, rt.Transect, "
    "rs.PlantCode;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe05aee80481b24408c72781cce05e5f3
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="rt.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8c553c307cee5147acbe5f2c04aa69df
        End
    End
    Begin
        dbText "Name" ="rt.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7da5025e1203c748bcc9fdda1b3fbdc6
        End
    End
    Begin
        dbText "Name" ="rt.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7e3ec8b88e7aa24abb8f8257e2bddfac
        End
    End
    Begin
        dbText "Name" ="rt.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf7399dfe18cf1043876d19e7984be1b4
        End
    End
    Begin
        dbText "Name" ="rt.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x09f47178964bca44bdce58efd01d7f06
        End
    End
    Begin
        dbText "Name" ="rs.PlantCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4ac7c9256e0ade46ae20db11ec0ee974
        End
    End
    Begin
        dbText "Name" ="rs.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3bb12ee094a48945b17ad149313bb8a4
        End
    End
    Begin
        dbText "Name" ="rt.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="rt.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QuadratsSampled"
        dbLong "AggregateType" ="-1"
    End
End
