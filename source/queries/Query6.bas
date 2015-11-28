dbMemo "SQL" ="SELECT DISTINCT Master_Plant_Code_FK, LU_Code, Family, Species_Name, utah_specie"
    "s, Co_Species, Wy_Species, Master_Common_Name, ConcatRelated(\"ParkYearPriority\""
    ", \"qry_Annual_Complete_Tgt_Species_Lists\",\"Park= 'ZION' AND Species_Name='\"+"
    "Species_Name+\"'\",'',\"|\") AS ParkYearPriorities, (SELECT Min(TgtYear) FROM qr"
    "y_Annual_Complete_Tgt_Species_Lists WHERE Park = 'ZION') AS MinYear, (SELECT Max"
    "(TgtYear) FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park = 'ZION') AS Max"
    "Year\015\012FROM (SELECT * FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park"
    " = 'ZION')  AS [%$##@_Alias]\015\012GROUP BY Park, Master_Plant_Code_FK, LU_Code"
    ", Family, Species_Name, Priority, Transect_Only, Target_Area_ID, Tgt_Area, utah_"
    "species, Co_Species, Wy_Species, Master_Common_Name, PriorityTarget, SpeciesYear"
    ";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xcd8a5cddd5173047a5b3664ba74fabd3
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
        dbInteger "ColumnWidth" ="3795"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xde2df9d68faab943a6484a041894c5d5
        End
    End
    Begin
        dbText "Name" ="MinYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x92e445d5cf62464b89318b01d3b84632
        End
    End
    Begin
        dbText "Name" ="MaxYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x90c380e6f6c2b0499f57f8934bfda966
        End
    End
End
