dbMemo "SQL" ="SELECT q.*, (\015\012IIF(q.Q1 > 0, 1,0) +\015\012IIF(q.Q2 > 0, 1,0) +\015\012IIF"
    "(q.Q3 > 0, 1,0) +\015\012IIF(q.Q1_3m > 0, 1,0) +\015\012IIF(q.Q2_8m > 0, 1,0) +\015"
    "\012IIF(q.Q3_13m > 0, 1,0) +\015\012IIF(q.Q1_hm > 0, 1,0) +\015\012IIF(q.Q2_5m >"
    " 0, 1,0) +\015\012IIF(q.Q3_10m > 0, 1,0)\015\012) AS QCount\015\012FROM qry_Tran"
    "sect_Select AS q;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf44db26519d9474d8f5910601625e2a4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="q.tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00c4a73e8c3f7741bdfcd7018f9a3095
        End
    End
    Begin
        dbText "Name" ="q.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc4bf7a880a80d04abcebed5a75e5da0b
        End
    End
    Begin
        dbText "Name" ="q.tbl_Locations.Plot_ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6f25d0a8f1fc5341b61e5442fdae8d74
        End
    End
    Begin
        dbText "Name" ="q.tbl_Quadrat_Transect.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xacefdb706025b84cbc53d0dc5dd4c3ea
        End
    End
    Begin
        dbText "Name" ="q.tbl_Locations.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x476738a79c82cd4c88b50fec9e324509
        End
    End
    Begin
        dbText "Name" ="q.Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x06ce9f386cb42846a7b4019d2092a750
        End
    End
    Begin
        dbText "Name" ="q.tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x561fb5eb47754749876e9f803a8da0d6
        End
    End
    Begin
        dbText "Name" ="q.tbl_Quadrat_Transect.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8e9b84c3200bc644aee9193000a39cbc
        End
    End
    Begin
        dbText "Name" ="q.tbl_Quadrat_Transect.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7d11e904e63bd6478530bf4d0ad7276d
        End
    End
    Begin
        dbText "Name" ="q.Q1_hm"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xff5c62635cb98044bd83d592ed6fbffe
        End
    End
    Begin
        dbText "Name" ="q.Q2_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe1b7814b2379cc4e98e8f3d36abfb3e8
        End
    End
    Begin
        dbText "Name" ="q.Q3_10m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0ba7019d6cad204598660ff41e45815c
        End
    End
    Begin
        dbText "Name" ="q.Q1_3m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ea4f61a21166b408e0054bbda71f9a9
        End
    End
    Begin
        dbText "Name" ="q.Q2_8m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x673b5cd041aebe4b910b23f9a27fbfa2
        End
    End
    Begin
        dbText "Name" ="q.Q3_13m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe14dc43d268a184dbdfe492a5bcc4937
        End
    End
    Begin
        dbText "Name" ="q.Q1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x97a9257c0fc84546a4dde47bf2b7d7ec
        End
    End
    Begin
        dbText "Name" ="q.Q2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1f2a17b116c191409b36c7eabd00beaf
        End
    End
    Begin
        dbText "Name" ="q.Q3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x147dfc6c28783c489eb86228b9aa715f
        End
    End
    Begin
        dbText "Name" ="QCount"
        dbLong "AggregateType" ="0"
        dbBinary "GUID" = Begin
            0x0c7dc1eeeeb0df44baf1fd4a01fc46a0
        End
    End
End
