dbMemo "SQL" ="SELECT tbl_Target_Species.Park_Code AS Park, tbl_Target_Species.Target_Year AS T"
    "gtYear, tbl_Target_Species.Master_Plant_Code_FK, tbl_Target_Species.Species_Name"
    ", tbl_Target_Species.LU_Code, tbl_Target_Species.Priority, tbl_Target_Species.Tr"
    "ansect_Only, tbl_Target_Species.Target_Area_ID, tbl_Target_Species.Comments\015\012"
    "FROM tbl_Target_Species\015\012WHERE (((tbl_Target_Species.Target_Year)=CInt(201"
    "6)) And ((LCase(tbl_Target_Species.Park_Code))=LCase('BLCA')))\015\012ORDER BY t"
    "bl_Target_Species.Species_Name;\015\012"
dbMemo "Connect" =""
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
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x75272ebc8a14244ba60bba9c13abe894
        End
    End
    Begin
        dbText "Name" ="TgtYear"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x76c86558f3b54e498c4931c5078371d9
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Transect_Only"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Target_Area_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Comments"
        dbLong "AggregateType" ="-1"
    End
End
