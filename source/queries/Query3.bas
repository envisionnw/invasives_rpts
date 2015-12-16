dbMemo "SQL" ="SELECT DISTINCT Master_Plant_Code_FK, LU_Code, Family, Species_Name, utah_specie"
    "s, Co_Species, Wy_Species, Master_Common_Name, ConcatRelated(\"ParkYearPriority\""
    ",\"qry_Annual_Complete_Tgt_Species_Lists\",\"Park = '\" & [Tempvars]![Park] & \""
    "' AND Species_Name='\"+Species_Name+\"'\",'',\"|\") AS ParkYearPriorities, (C) A"
    "S MinYear, (SELECT Max(TgtYear) FROM qry_Annual_Complete_Tgt_Species_Lists WHERE"
    " Park = \"'\" & [Tempvars]![Park] & \"'\") AS MaxYear\015\012FROM (SELECT * FROM"
    " qry_Annual_Complete_Tgt_Species_Lists WHERE Park = \"'\" & [Tempvars]![Park] & "
    "\"'\")  AS [%$##@_Alias]\015\012GROUP BY Park, Master_Plant_Code_FK, LU_Code, Fa"
    "mily, Species_Name, Priority, Transect_Only, Target_Area_ID, Tgt_Area, utah_spec"
    "ies, Co_Species, Wy_Species, Master_Common_Name, PriorityTarget, ParkYearPriorit"
    "y, SpeciesYear;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xab0a2fdd743f0348bfc36f85fc2a2af8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Co_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Wy_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkYearPriorities"
        dbInteger "ColumnWidth" ="3144"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5fe491dcb756e141a63bfc4d3a2f04aa
        End
    End
    Begin
        dbText "Name" ="MinYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf85e7c8f96d3b141a536da67eeb272a4
        End
    End
    Begin
        dbText "Name" ="MaxYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x90ee3e7732671b4bafc4a40fb7a6c169
        End
    End
End
