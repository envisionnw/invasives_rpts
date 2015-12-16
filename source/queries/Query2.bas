dbMemo "SQL" ="SELECT Park, TgtYear, Master_Plant_Code_FK, LU_Code, Family, Species_Name, Prior"
    "ity, Transect_Only, Target_Area_ID, Tgt_Area, utah_species, Co_Species, Wy_Speci"
    "es, Master_Common_Name, PriorityTarget, ParkPriority, ParkYearPriority, SpeciesY"
    "ear, (SELECT Max(TgtYear) FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park "
    "= 'CURE') AS MaxYear, (SELECT Min(TgtYear) FROM qry_Annual_Complete_Tgt_Species_"
    "Lists WHERE Park = 'CURE') AS MinYear\015\012FROM qry_Annual_Complete_Tgt_Specie"
    "s_Lists\015\012WHERE Park = 'CURE'\015\012GROUP BY Park, TgtYear, Master_Plant_C"
    "ode_FK, LU_Code, Family, Species_Name, Priority, Transect_Only, Target_Area_ID, "
    "Tgt_Area, utah_species, Co_Species, Wy_Species, Master_Common_Name, PriorityTarg"
    "et, ParkPriority, ParkYearPriority, SpeciesYear;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5791ecf026ecd347a53b2b9c7ed397e2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="MaxYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x92503bb4b1ac2c40821a6cddf40bec09
        End
    End
    Begin
        dbText "Name" ="Park"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TgtYear"
        dbLong "AggregateType" ="-1"
    End
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
        dbText "Name" ="Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Only"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Target_Area_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tgt_Area"
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
        dbText "Name" ="PriorityTarget"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkPriority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkYearPriority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.Park"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb41a39d808b8ea4e84121a2bc4313e20
        End
    End
End
