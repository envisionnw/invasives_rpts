Operation =1
Option =2
Begin InputTables
    Name ="qryAnnualCompleteTgtSpeciesLists"
End
Begin OutputColumns
    Expression ="qryAnnualCompleteTgtSpeciesLists.TgtYear"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Master_Plant_Code_FK"
    Expression ="qryAnnualCompleteTgtSpeciesLists.LU_Code"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Family"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Species_Name"
    Expression ="qryAnnualCompleteTgtSpeciesLists.utah_species"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Co_Species"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Wy_Species"
    Expression ="qryAnnualCompleteTgtSpeciesLists.Master_Common_Name"
    Alias ="ParkPriorities"
    Expression ="ConcatRelated(\"ParkPriority\",\"qryAnnualCompleteTgtSpeciesLists\",\"Species_Na"
        "me='\"+Species_Name+\"'\",'',\"|\")"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x407218643ce7154bbf8f499f0b2c6124
End
Begin
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Co_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Wy_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ParkPriorities"
        dbInteger "ColumnWidth" ="5625"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd9889c2e6375c048961504abd97141f5
        End
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryAnnualCompleteTgtSpeciesLists.TgtYear"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1332
    Bottom =625
    Left =-1
    Top =-1
    Right =1294
    Bottom =359
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =15
        Right =296
        Bottom =402
        Top =0
        Name ="qryAnnualCompleteTgtSpeciesLists"
        Name =""
    End
End
