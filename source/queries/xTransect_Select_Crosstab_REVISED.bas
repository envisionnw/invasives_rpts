dbMemo "SQL" ="TRANSFORM Min(ts.PercentCover) AS MinOfPercentCover\015\012SELECT ts.Route, ts.T"
    "ransect, ts.Species, ts.IsDead\015\012FROM Transect_Select_LIMITED_ESP_SpeciesCo"
    "ver_Species AS ts\015\012GROUP BY ts.Route, ts.Transect, ts.Species, ts.IsDead\015"
    "\012PIVOT ts.ColName;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x7d82eda03d98524c834c768075e9d0f8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x766538fb7919a24c8f204b1a3efcb73f
        End
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x95b8b789dff075479b1327c3f3a8798b
        End
    End
    Begin
        dbText "Name" ="ts.Species"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x71e54786ec491140b64c84be7da5828e
        End
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x94a51eaa403b574182ca04e5086a6b43
        End
    End
    Begin
        dbText "Name" ="Q1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x40d164be5349f84caa1996c5bb5d3349
        End
    End
    Begin
        dbText "Name" ="Q1_0m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x90a189c0643225449a0602c5a348c9ad
        End
    End
    Begin
        dbText "Name" ="Q1_3m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x88b0cbfe22e6204ba6b38e431a232e50
        End
    End
    Begin
        dbText "Name" ="Q2_8m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbb46878be83cc14daac669452945314a
        End
    End
    Begin
        dbText "Name" ="Q3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x33516762658fdc4ca222200968010a88
        End
    End
    Begin
        dbText "Name" ="Q3_13m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3000bcad718bc04ebbe1b8644f7d5ac5
        End
    End
    Begin
        dbText "Name" ="Q3_9_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x337f977cb7cc5347abd600772e114080
        End
    End
    Begin
        dbText "Name" ="Q2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x28b68a93d519d448ab09b3dd90c0ee0a
        End
    End
    Begin
        dbText "Name" ="Q2_4_5m"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x660c2fc64b6ecc4aabeb559c9d254f28
        End
    End
    Begin
        dbText "Name" ="MinOfPercentCover"
        dbBinary "GUID" = Begin
            0x7e91a729e3b6b148948ddfd792ac2af6
        End
    End
End
