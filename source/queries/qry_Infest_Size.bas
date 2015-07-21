Operation =1
Option =0
Where ="(((qry_Infest_Size_Select.Size_Class) Is Not Null))"
Begin InputTables
    Name ="qry_Infest_Size_Select"
    Name ="tbl_Target_Species"
    Name ="tbl_Target_List"
End
Begin OutputColumns
    Expression ="qry_Infest_Size_Select.*"
    Expression ="tbl_Target_Species.Priority"
End
Begin Joins
    LeftTable ="tbl_Target_Species"
    RightTable ="tbl_Target_List"
    Expression ="tbl_Target_Species.Tgt_List_ID_FK = tbl_Target_List.Tgt_List_ID"
    Flag =1
    LeftTable ="qry_Infest_Size_Select"
    RightTable ="tbl_Target_List"
    Expression ="qry_Infest_Size_Select.Unit_Code = tbl_Target_List.Park_Code"
    Flag =1
    LeftTable ="qry_Infest_Size_Select"
    RightTable ="tbl_Target_List"
    Expression ="qry_Infest_Size_Select.Visit_Year = tbl_Target_List.Target_Year"
    Flag =1
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
End
Begin
    State =0
    Left =2
    Top =14
    Right =1299
    Bottom =739
    Left =-1
    Top =-1
    Right =1259
    Bottom =-1
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =233
        Bottom =210
        Top =0
        Name ="qry_Infest_Size_Select"
        Name =""
    End
    Begin
        Left =584
        Top =-5
        Right =764
        Bottom =190
        Top =0
        Name ="tbl_Target_Species"
        Name =""
    End
    Begin
        Left =359
        Top =6
        Right =539
        Bottom =186
        Top =0
        Name ="tbl_Target_List"
        Name =""
    End
End
