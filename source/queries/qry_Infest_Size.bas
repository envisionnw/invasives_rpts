Operation =1
Option =2
Where ="(((qry_Infest_Size_Select.Size_Class) Is Not Null))"
Begin InputTables
    Name ="qry_Infest_Size_Select"
    Name ="qry_Annual_Complete_Tgt_Species_Lists"
End
Begin OutputColumns
    Expression ="qry_Infest_Size_Select.*"
    Expression ="qry_Annual_Complete_Tgt_Species_Lists.Priority"
End
Begin Joins
    LeftTable ="qry_Infest_Size_Select"
    RightTable ="qry_Annual_Complete_Tgt_Species_Lists"
    Expression ="qry_Infest_Size_Select.Unit_Code = qry_Annual_Complete_Tgt_Species_Lists.Park"
    Flag =2
    LeftTable ="qry_Infest_Size_Select"
    RightTable ="qry_Annual_Complete_Tgt_Species_Lists"
    Expression ="qry_Infest_Size_Select.Visit_Year = qry_Annual_Complete_Tgt_Species_Lists.TgtYea"
        "r"
    Flag =2
    LeftTable ="qry_Infest_Size_Select"
    RightTable ="qry_Annual_Complete_Tgt_Species_Lists"
    Expression ="qry_Infest_Size_Select.Master_Code = qry_Annual_Complete_Tgt_Species_Lists.Maste"
        "r_Plant_Code_FK"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x29e685a9421a354390ecb3bad41b7881
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="((([qry_Infest_Size].[Unit_Code]=\"GOSP\"))) AND ([qry_Infest_Size].[Visit_Year]"
    "=2012)"
Begin
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Infestation.Pulled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Infestation.Growth_Stage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Infestation.N_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Infestation.E_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tlu_Size_Class.Size_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Infest_Size_Select.tbl_Infestation.Master_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Annual_Complete_Tgt_Species_Lists.Priority"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =81
    Top =119
    Right =842
    Bottom =729
    Left =-1
    Top =-1
    Right =729
    Bottom =343
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =222
        Bottom =259
        Top =0
        Name ="qry_Infest_Size_Select"
        Name =""
    End
    Begin
        Left =270
        Top =12
        Right =574
        Bottom =344
        Top =0
        Name ="qry_Annual_Complete_Tgt_Species_Lists"
        Name =""
    End
End
