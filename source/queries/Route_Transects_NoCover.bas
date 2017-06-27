dbMemo "SQL" ="SELECT DISTINCT Transect_Select_SpeciesCover.Transect\015\012FROM Transect_Selec"
    "t_SpeciesCover LEFT JOIN Transect_Data ON Transect_Select_SpeciesCover.[Transect"
    "] = Transect_Data.[Transect]\015\012WHERE (((Transect_Data.Transect) Is Null));\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2eecf105b249f14eb8028b2a4e33c4e6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="sc.Transect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Select_SpeciesCover.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf756b4793481ed4f9c88196ad5ec67d9
        End
    End
End
