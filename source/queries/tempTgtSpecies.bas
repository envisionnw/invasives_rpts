Operation =1
Option =0
Where ="(((tbl_Target_Species.Target_Year)=CInt(2016)) And ((LCase(tbl_Target_Species.Pa"
    "rk_Code))=LCase('BLCA')))"
Begin InputTables
    Name ="tbl_Target_Species"
End
Begin OutputColumns
    Alias ="Park"
    Expression ="tbl_Target_Species.Park_Code"
    Alias ="TgtYear"
    Expression ="tbl_Target_Species.Target_Year"
    Expression ="tbl_Target_Species.Master_Plant_Code_FK"
    Expression ="tbl_Target_Species.Species_Name"
    Expression ="tbl_Target_Species.LU_Code"
    Expression ="tbl_Target_Species.Priority"
    Expression ="tbl_Target_Species.Transect_Only"
    Expression ="tbl_Target_Species.Target_Area_ID"
End
Begin OrderBy
    Expression ="tbl_Target_Species.Species_Name"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xcb09984879db7b44a0f1496b1eed8ee0
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Park"
        dbBinary "GUID" = Begin
            0x75272ebc8a14244ba60bba9c13abe894
        End
    End
    Begin
        dbText "Name" ="TgtYear"
        dbBinary "GUID" = Begin
            0x76c86558f3b54e498c4931c5078371d9
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =835
    Bottom =821
    Left =-1
    Top =-1
    Right =803
    Bottom =498
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Target_Species"
        Name =""
    End
End
