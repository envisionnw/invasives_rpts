dbMemo "SQL" ="SELECT ts.ID, ts.Unit_Code, ts.Visit_Year, ts.Route, ts.Transect_ID, ts.Transect"
    ", ts.Area, ts.E_Coord, ts.N_Coord, ts.Position_m, ts.ColName, ts.IsSampled, ts.N"
    "oExotics, ts.PlantCode, IIf([Unit_Code] In (\"CARE\",\"DINO\",\"GOSP\",\"ZION\")"
    ",[Utah_Species],\015\012IIf([Unit_Code]=\"FOBU\",[WY_Species],[Co_Species])) AS "
    "Species, tlu_NCPN_Plants.Master_Common_Name, ts.IsDead, ts.PercentCover\015\012F"
    "ROM Transect_Select_LIMITED AS ts LEFT JOIN tlu_NCPN_Plants ON tlu_NCPN_Plants.M"
    "aster_PLANT_Code = ts.PlantCode\015\012WHERE (((\015\012IIf([Unit_Code] In (\"CA"
    "RE\",\"DINO\",\"GOSP\",\"ZION\"),[Utah_Species],\015\012IIf([Unit_Code]=\"FOBU\""
    ",[WY_Species],[Co_Species]))) Is Not Null))\015\012ORDER BY ts.Unit_Code, ts.Vis"
    "it_Year, ts.Plot_ID, ts.Transect, ts.Quadrat, IIf([Unit_Code] In (\"CARE\",\"DIN"
    "O\",\"GOSP\",\"ZION\"),[Utah_Species],\015\012IIf([Unit_Code]=\"FOBU\",[WY_Speci"
    "es],[Co_Species]));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1567b102df74ac4d9aa60b3e8aa567d3
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Species"
        dbBinary "GUID" = Begin
            0x0417f4ef9b2f7446a93994842b001227
        End
    End
End
