dbMemo "SQL" ="UPDATE temp_Route_Transect_AverageCover AS ac SET ac.TransectsDetected = (\015\012"
    "SELECT td.TransectsDetected\015\012FROM temp_Update_Table td\015\012INNER JOIN t"
    "emp_Route_Transect_AverageCover ac\015\012ON\015\012td.Unit_Code = ac.Unit_Code\015"
    "\012AND td.Visit_Year = ac.Visit_Year\015\012AND td.Route = ac.Route\015\012AND "
    "td.Transect = ac.Transect\015\012AND td.PlantCode = ac.PlantCode\015\012AND td.I"
    "sDead = ac.IsDead\015\012);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x28b4d21c1d25a14c822652ad4968ca84
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ac.TransectsDetected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.TransectsDetected"
        dbLong "AggregateType" ="-1"
    End
End
