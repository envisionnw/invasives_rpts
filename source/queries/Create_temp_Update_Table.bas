dbMemo "SQL" ="SELECT * INTO temp_Update_Table\015\012FROM (SELECT DISTINCT \015\012ac.Unit_Cod"
    "e,\015\012ac.Visit_Year,\015\012ac.Route,\015\012ac.Transect,\015\012ac.PlantCod"
    "e,\015\012ac.IsDead,\015\012td.TransectsDetected\015\012FROM temp_Route_Transect"
    "_AverageCover ac\015\012LEFT JOIN temp_Route_TransectsDetected td\015\012ON td.U"
    "nit_Code = ac.Unit_Code\015\012AND td.Visit_Year = ac.Visit_Year\015\012AND td.R"
    "oute = ac.Route\015\012AND td.PlantCode = ac.PlantCode\015\012AND td.IsDead = ac"
    ".IsDead\015\012)  AS [%$##@_Alias];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa497ae81e6d82641a9ed822de647853d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="([Query1].[IsDead]=1)"
Begin
    Begin
        dbText "Name" ="td.TransectsDetected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ac.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ac.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ac.Route"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ac.PlantCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ac.IsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ac.TransectsDetected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ac.Transect"
        dbLong "AggregateType" ="-1"
    End
End
