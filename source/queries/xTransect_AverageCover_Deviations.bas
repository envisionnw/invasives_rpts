dbMemo "SQL" ="SELECT DISTINCT ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, ts.Area, ts."
    "E_Coord, ts.N_Coord, ts.Species, ts.Master_Common_Name, ts.IsDead, st.TransectsS"
    "ampled, tc.TotalCover, ta.AverageCover AS TransectAverageCover, ac.AverageCover "
    "AS RouteAverageCover, (ac.AverageCover - ta.AverageCover) AS Deviation, (ac.Aver"
    "ageCover - ta.AverageCover)^2 AS DeviationSquared\015\012FROM (((Transect_Select"
    "_SpeciesCover AS ts INNER JOIN Transect_AverageCover AS ta ON (ta.Unit_Code = ts"
    ".Unit_Code) AND (ta.Visit_Year = ts.Visit_Year) AND (ta.Route = ts.Route) AND (t"
    "a.Transect = ts.Transect) AND (ta.PlantCode = ts.PlantCode) AND (ta.IsDead = ts."
    "IsDead)) INNER JOIN Route_TransectsSampled AS st ON (st.Unit_Code = ts.Unit_Code"
    ") AND (st.Visit_Year = ts.Visit_Year) AND (st.Route = ts.Route)) INNER JOIN Tran"
    "sect_Select_TotalCover AS tc ON (tc.Unit_Code = ts.Unit_Code) AND (tc.Visit_Year"
    " = ts.Visit_Year) AND (tc.Route = ts.Route) AND (tc.Transect = ts.Transect) AND "
    "(tc.E_Coord = ts.E_Coord) AND (tc.N_Coord = ts.N_Coord) AND (tc.PlantCode = ts.P"
    "lantCode) AND (tc.IsDead = ts.IsDead)) LEFT JOIN Route_AverageCover AS ac ON (ac"
    ".Unit_Code = ts.Unit_Code) AND (ac.Visit_Year = ts.Visit_Year) AND (ac.Route = t"
    "s.Route) AND (ac.PlantCode = ts.PlantCode) AND (ac.IsDead = ts.IsDead)\015\012OR"
    "DER BY ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect, ts.Species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x573a879819f4224aa200cc43554bb542
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="TransectAverageCover"
        dbInteger "ColumnWidth" ="2430"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfaca46e98b962648a16f1e1f91e29a28
        End
    End
    Begin
        dbText "Name" ="ts.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xba19f8749776584e87a09576cb59433b
        End
    End
    Begin
        dbText "Name" ="ts.Visit_Year"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2a97bc9fc9ae3d4090d7762204fd9729
        End
    End
    Begin
        dbText "Name" ="ts.Route"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa3bedeaf5f61f54587beb413f1c292ae
        End
        dbInteger "ColumnWidth" ="2385"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ts.Transect"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6c650a88b14c104d9e306f55a3c6af7e
        End
    End
    Begin
        dbText "Name" ="ts.Area"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe7a4d024f0e87d4794585dc8b2f1c911
        End
    End
    Begin
        dbText "Name" ="ts.E_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x90e4d1591fd24941a05bf9f4de815eca
        End
    End
    Begin
        dbText "Name" ="ts.N_Coord"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8f26925b88d8cc428d74ffc278896247
        End
    End
    Begin
        dbText "Name" ="ts.Species"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4dc15cb93915f346975f012ca2d59574
        End
    End
    Begin
        dbText "Name" ="ts.Master_Common_Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf1adebe696d85e4a851414b49834dcf9
        End
    End
    Begin
        dbText "Name" ="ts.IsDead"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa8b2d796828395438ab96955da65277c
        End
    End
    Begin
        dbText "Name" ="st.TransectsSampled"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7367b7cca6c2d04b96aa5934cd329be2
        End
    End
    Begin
        dbText "Name" ="tc.TotalCover"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8276cd6901151540a8c1793af3077f16
        End
    End
    Begin
        dbText "Name" ="RouteAverageCover"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc8af717404e6fa4bb1cc9dc4c204c760
        End
    End
    Begin
        dbText "Name" ="Deviation"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5daea19f27925141b271303a3e750f05
        End
    End
    Begin
        dbText "Name" ="DeviationSquared"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc6b79b0c43850943a75a3a2797cc253e
        End
    End
End
