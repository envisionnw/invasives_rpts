dbMemo "SQL" ="INSERT INTO tbl_Target_Species ( Master_Plant_Code_FK, Park_Code, Target_Year, P"
    "riority, Transect_Only, Comments, LU_Code, Species_Name )\015\012SELECT tbl_Targ"
    "et_Plant_Lists.Master_Plant_Code, Unit_Code, Visit_Year, IIF(Priority=1,Priority"
    ",0) AS Pri, IIF(LCase(Comments) Like \"*transects only*\", 1,0) AS Transect_Only"
    ", Comments, tlu_NCPN_Plants.LU_Code, tlu_NCPN_Plants.Master_Species\015\012FROM "
    "tbl_Target_Plant_Lists LEFT JOIN tlu_NCPN_Plants ON tbl_Target_Plant_Lists.Maste"
    "r_Plant_Code = tlu_NCPN_Plants.Master_PLANT_Code;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa283222453940f45af2e6b9ebbde1df1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tbl_Target_Plant_Lists.Visit_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Plant_Lists.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Plant_Lists.Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Plant_Lists.Report_Scientific_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Plant_Lists.Report_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Plant_Lists.Master_Plant_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdaad05096ccead49ab894bb52bde2017
        End
    End
    Begin
        dbText "Name" ="tbl_Target_Plant_Lists.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Master_Plant_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xce748e0dacc3b646bc673f5d843390c8
        End
    End
    Begin
        dbText "Name" ="Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbe7d6f961e99994b84829965db2f2803
        End
    End
    Begin
        dbText "Name" ="Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transect_Only"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfbf049719de3694ba89edb087d188d57
        End
    End
    Begin
        dbText "Name" ="Comments"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2090d028f57a1945a57755130298527e
        End
    End
    Begin
        dbText "Name" ="Expr1003"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Pri"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2a027b997c772c43bc7cafccabd9df1c
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.LU_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x197c6851e0abf14dbdb87ea4cf117e27
        End
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd08ba54a0f3fc94a9d6447354d5d0c0e
        End
    End
End
